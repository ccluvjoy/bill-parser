
# -*- coding: utf-8 -*-
"""
账单流水解析工具（Streamlit 网页应用）

功能：
    粘贴或上传微信 / 支付宝等账单流水文本，调用 DeepSeek API 自动解析为
    结构化账单记录（时间、类型、分类、金额、商家、备注），支持表格展示与 CSV 导出。

启动方法：
    1. 安装依赖：pip install streamlit pandas requests
    2. 运行：    streamlit run bill_parser_app.py
    3. 浏览器打开提示的地址（默认 http://localhost:8501）
    4. 在左侧边栏填入 DeepSeek API Key，粘贴流水文本，点击"开始解析"
"""

import datetime
import json
import os
import re
import subprocess
import tempfile

import pandas as pd
import requests
import streamlit as st

# ======================= 常量区（Prompt 集中管理，方便修改） =======================

# DeepSeek 对话接口地址（OpenAI 兼容格式）
DEEPSEEK_API_URL = "https://api.deepseek.com/chat/completions"

# 可选模型列表：deepseek-chat 为通用对话模型，deepseek-reasoner 为推理模型
MODEL_OPTIONS = ["deepseek-chat", "deepseek-reasoner"]

# 单次请求超时时间（秒）：流水较长时推理耗时较多，预留充足时间
API_TIMEOUT = 180

# 允许的账单类型与分类（与 Prompt 中的约束保持一致，用于结果规整）
ALLOWED_TYPES = {"income", "expense", "transfer"}
ALLOWED_CATEGORIES = {
    "餐饮", "交通", "购物", "房租住宿", "学习资料",
    "医药", "娱乐", "转账", "红包", "充值", "其他",
}

# ---------- 账单解析 Prompt（固定模板，模型按此格式输出严格 JSON） ----------
BILL_PARSE_SYSTEM_PROMPT = """你是一个账单解析器，任务：
解析支付流水，输出严格JSON。
可选分类：餐饮，交通，购物，房租住宿，学习资料，医药，娱乐，转账，红包，充值，其他。
type只能是income/expense/transfer。
时间格式：year-month-day，或者year-month-day hour:min。
忽略广告、无关提示、解释。
格式：{"records":[{"time":"","type":"","category":"","amount":数字,"merchant":"","remark":""}]}
只输出 JSON 本身，不要输出任何解释、代码块标记或多余文字。"""

# 用户消息模板：{text} 会被替换为原始流水文本
BILL_PARSE_USER_TEMPLATE = "流水:{text}"

# ---------- 修复重试 Prompt（模型第一次输出非法 JSON 时使用） ----------
REPAIR_SYSTEM_PROMPT = """你是一个账单解析器。你上一次的输出不是合法的 JSON，无法被程序解析。
请根据原始流水文本重新解析，输出严格 JSON。
要求：
- 可选分类：餐饮，交通，购物，房租住宿，学习资料，医药，娱乐，转账，红包，充值，其他。
- type只能是income/expense/transfer。
- 时间格式：year-month-day，或者year-month-day hour:min。
- 格式：{"records":[{"time":"","type":"","category":"","amount":数字,"merchant":"","remark":""}]}
只输出 JSON 本身，不要输出任何解释、代码块标记或多余文字。"""

REPAIR_USER_TEMPLATE = """原始流水文本如下：
{text}

你上一次的错误输出如下：
{raw_output}

请重新输出修复后的严格 JSON："""

# ---------- 图片识别脚本（Windows 内置 OCR，程序运行时写入临时 .ps1 再执行） ----------
# 用途：上传 .jpg/.jpeg/.png 账单截图时，先把图片识别成文字，再交给大模型解析。
# 依赖 Windows 的 WinRT OCR 引擎与中文语言包（zh-Hans-CN），无需安装任何第三方 OCR 库。
OCR_POWERSHELL_SCRIPT = """param([string]$ImagePath)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine,Windows.Foundation,ContentType=WindowsRuntime]
$null = [Windows.Foundation.IAsyncOperation`1,Windows.Foundation,ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.SoftwareBitmap,Windows.Foundation,ContentType=WindowsRuntime]
$null = [Windows.Globalization.Language,Windows.Globalization,ContentType=WindowsRuntime]
function Await($t, $rt) {
    $asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
    $task = $asTask.MakeGenericMethod($rt).Invoke($null, @($t))
    $task.Wait(-1) | Out-Null
    $task.Result
}
$file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($ImagePath)) ([Windows.Storage.StorageFile])
$stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
$decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
$bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
$lang = New-Object Windows.Globalization.Language 'zh-Hans-CN'
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
if ($null -eq $engine) { Write-Error 'OCR_LANG_MISSING'; exit 2 }
$result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
foreach ($line in $result.Lines) { Write-Output $line.Text }
"""


# ======================= 自定义异常 =======================

class BillParseError(Exception):
    """面向用户的友好错误：message 内容直接展示给用户，不抛出堆栈。"""


# ======================= 工具函数 =======================

def call_deepseek(api_key, model, messages):
    """调用 DeepSeek 对话接口，返回模型输出文本；失败时抛出 BillParseError（中文友好提示）。"""
    try:
        resp = requests.post(
            DEEPSEEK_API_URL,
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            json={"model": model, "messages": messages, "stream": False},
            timeout=API_TIMEOUT,
        )
    except requests.exceptions.Timeout:
        raise BillParseError("请求超时：模型响应时间过长，请稍后重试。")
    except requests.exceptions.ConnectionError:
        raise BillParseError("无法连接 DeepSeek API，请检查网络是否正常。")
    except requests.exceptions.RequestException as e:
        raise BillParseError(f"网络请求失败：{e}")

    # HTTP 状态码 → 对应的中文友好提示
    if resp.status_code == 401:
        raise BillParseError("API Key 无效或已过期，请在左侧边栏检查后重试。")
    if resp.status_code == 402:
        raise BillParseError("API 账户余额不足，请前往 DeepSeek 平台充值。")
    if resp.status_code == 429:
        raise BillParseError("请求过于频繁或额度受限，请稍后重试。")
    if resp.status_code >= 500:
        raise BillParseError(f"DeepSeek 服务器暂时不可用（错误码 {resp.status_code}），请稍后重试。")
    if resp.status_code != 200:
        raise BillParseError(f"API 返回异常状态码 {resp.status_code}，请稍后重试。")

    # 解析响应体（服务器偶尔会返回非 JSON 的错误页，需要容错）
    try:
        data = resp.json()
    except ValueError:
        raise BillParseError("API 返回内容无法解析，请稍后重试。")

    content = data.get("choices", [{}])[0].get("message", {}).get("content", "")
    if not content or not content.strip():
        raise BillParseError("模型返回内容为空，请重试。")
    return content


def extract_json(text):
    """从模型输出中提取 JSON 对象：去掉 Markdown 代码块围栏后，取第一个 { 到最后一个 }。"""
    if not text:
        return None
    # 去掉 ```json ... ``` 之类的围栏标记
    text = re.sub(r"```[a-zA-Z]*", "", text).strip()
    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1 or end <= start:
        return None
    try:
        return json.loads(text[start:end + 1])
    except (json.JSONDecodeError, ValueError):
        return None


def normalize_record(rec):
    """把单条记录规整成统一字段与类型，缺省 / 非法值回退到安全默认值。"""
    # 时间：统一转字符串
    time_str = str(rec.get("time", "")).strip()

    # type：只允许 income/expense/transfer，非法值回退为 expense（账单多为支出）
    type_str = str(rec.get("type", "expense")).strip().lower()
    if type_str not in ALLOWED_TYPES:
        type_str = "expense"

    # 分类：为空回退为"其他"，其余保留模型判断结果
    category = str(rec.get("category", "")).strip()
    if category == "":
        category = "其他"

    # 金额：尽量转成数字，转不了回退 0.0
    try:
        amount = float(rec.get("amount", 0))
    except (TypeError, ValueError):
        amount = 0.0

    return {
        "time": time_str,
        "type": type_str,
        "category": category,
        "amount": amount,
        "merchant": str(rec.get("merchant", "")).strip(),
        "remark": str(rec.get("remark", "")).strip(),
    }


def validate_records(data):
    """校验解析结果结构：必须是 {"records": [...]} 形式，否则返回 None（触发修复重试）。"""
    if not isinstance(data, dict):
        return None
    records = data.get("records")
    if not isinstance(records, list):
        return None
    # 只保留字典类型的条目，并逐条规整字段
    return [normalize_record(r) for r in records if isinstance(r, dict)]


def parse_bill_text(api_key, model, text):
    """
    完整解析流程：
    1. 用账单解析 Prompt 调用模型；
    2. 若返回非法 JSON，自动用修复 Prompt 重试一次；
    3. 仍失败则抛出 BillParseError（前端展示友好提示）。
    """
    # 第一轮：正常解析
    content = call_deepseek(api_key, model, [
        {"role": "system", "content": BILL_PARSE_SYSTEM_PROMPT},
        {"role": "user", "content": BILL_PARSE_USER_TEMPLATE.format(text=text)},
    ])
    records = validate_records(extract_json(content))
    if records is not None:
        return records

    # 第二轮：修复重试（把原始文本和上一次的错误输出一起交给模型，截断防止超长）
    repair_content = call_deepseek(api_key, model, [
        {"role": "system", "content": REPAIR_SYSTEM_PROMPT},
        {"role": "user", "content": REPAIR_USER_TEMPLATE.format(
            text=text, raw_output=content[:3000]
        )},
    ])
    records = validate_records(extract_json(repair_content))
    if records is not None:
        return records

    raise BillParseError("解析失败：模型两次返回的内容都不是合法 JSON，请检查流水文本后重试。")


def read_uploaded_file(uploaded_file):
    """读取上传的文本文件，自动尝试 utf-8 / gbk 等常见编码。"""
    raw = uploaded_file.getvalue()
    for encoding in ("utf-8-sig", "utf-8", "gbk"):
        try:
            return raw.decode(encoding)
        except (UnicodeDecodeError, LookupError):
            continue
    # 全部失败则忽略无法解码的字节
    return raw.decode("utf-8", errors="ignore")


def ocr_image(image_bytes, file_name):
    """用 Windows 内置 OCR 识别账单截图中的文字（需 Windows 系统 + 中文 OCR 语言包）。

    流程：图片字节写入临时文件 → 把 PowerShell OCR 脚本写入临时 .ps1 → 执行 → 返回识别文字。
    识别失败时抛出 BillParseError（中文友好提示）。
    """
    # 1. 图片写入临时文件（保留原后缀，方便解码器识别格式）；先关闭句柄再交给 PowerShell 读取
    suffix = os.path.splitext(file_name)[1].lower() or ".png"
    tmp_img_f = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp_img = tmp_img_f.name
    tmp_img_f.close()
    # 2. OCR 脚本写入临时文件（utf-8-sig 带 BOM，防止 PowerShell 5.1 按 ANSI 误读）
    ps_f = tempfile.NamedTemporaryFile(delete=False, suffix=".ps1")
    ps_path = ps_f.name
    ps_f.close()
    try:
        with open(tmp_img, "wb") as f:
            f.write(image_bytes)
        with open(ps_path, "w", encoding="utf-8-sig") as f:
            f.write(OCR_POWERSHELL_SCRIPT)
        # 3. 执行 OCR：优先用 System32 下的绝对路径（兼容 PATH 不含 System32 的环境）
        ps_exe = r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
        if not os.path.exists(ps_exe):
            ps_exe = "powershell"
        try:
            proc = subprocess.run(
                [ps_exe, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps_path, tmp_img],
                capture_output=True,
                timeout=180,
            )
        except FileNotFoundError:
            raise BillParseError("图片识别不可用：当前系统没有 PowerShell，请改用文本方式输入。")
        except subprocess.TimeoutExpired:
            raise BillParseError("图片识别超时，请重试，或改用文本方式输入。")
        # 4. 校验执行结果，返回码非 0 说明识别失败，逐项映射为友好提示
        if proc.returncode != 0:
            err = (proc.stderr or b"").decode("utf-8", errors="replace")
            if "OCR_LANG_MISSING" in err:
                raise BillParseError(
                    "图片识别不可用：系统缺少中文 OCR 语言包"
                    "（Windows 设置 → 时间和语言 → 语言和区域 → 添加中文语言）。"
                )
            raise BillParseError("图片识别失败：请确认上传的是清晰的账单截图，或改用文本方式输入。")
        text = proc.stdout.decode("utf-8", errors="replace").strip()
        if not text:
            raise BillParseError("图片中没有识别出文字：请确认上传的是清晰的账单截图。")
        return text
    finally:
        # 5. 无论成功失败都清理临时文件
        for p in (tmp_img, ps_path):
            try:
                os.remove(p)
            except OSError:
                pass


# ======================= 页面 UI =======================

def main():
    # 页面基本设置（必须放在所有 Streamlit 组件之前）
    st.set_page_config(page_title="账单流水解析工具", page_icon="🧾", layout="wide")

    # ---------- 侧边栏：设置面板 ----------
    with st.sidebar:
        st.header("⚙️ 设置")
        # API Key 与模型保存在 session_state：仅当前会话有效，刷新页面后重新填写即可
        api_key = st.text_input("DeepSeek API Key", type="password", placeholder="sk-...", key="api_key")
        model = st.selectbox("模型", MODEL_OPTIONS, key="model")
        st.caption("设置保存在当前会话中，刷新页面后失效，重新填写即可。")

    # ---------- 主区域 ----------
    st.title("🧾 账单流水解析工具")
    st.caption("粘贴微信 / 支付宝等账单流水文本，一键解析为结构化记录")

    # 隐私提示：提示用户不要上传敏感内容
    st.info("🔒 请勿上传隐私敏感内容（身份证号、银行卡号、住址等）。流水文本会发送至 DeepSeek API 用于解析。")

    # 输入方式选择：粘贴文本 / 上传文件
    input_mode = st.radio("输入方式", ["粘贴文本", "上传文件"], horizontal=True)

    raw_text = ""
    if input_mode == "粘贴文本":
        # 大文本框，支持大段粘贴
        raw_text = st.text_area(
            "流水文本",
            height=260,
            placeholder="在这里粘贴微信 / 支付宝的账单流水文本，支持大段粘贴……",
            key="input_text",
        )
        st.caption(f"已输入 {len(raw_text)} 字符")
    else:
        # 上传文件：支持文本文件（.txt 等）与账单截图（.jpg/.jpeg/.png）
        uploaded_file = st.file_uploader(
            "上传账单流水文件（.txt / .jpg / .png）",
            type=["txt", "md", "csv", "log", "jpg", "jpeg", "png"],
            key="uploaded_file",
        )
        if uploaded_file is not None:
            name_lower = uploaded_file.name.lower()
            if name_lower.endswith((".jpg", ".jpeg", ".png")):
                # 图片文件：新上传时才执行一次 OCR（同一文件多次交互不重复识别）
                file_id = getattr(uploaded_file, "file_id", uploaded_file.name)
                if st.session_state.get("last_image_id") != file_id:
                    st.session_state["last_image_id"] = file_id
                    with st.spinner("正在识别图片文字，请稍候…"):
                        try:
                            ocr_text = ocr_image(uploaded_file.getvalue(), uploaded_file.name)
                        except BillParseError as e:
                            # 识别失败只展示友好提示，文本框留空供手动粘贴
                            st.session_state["flash"] = ("error", str(e))
                            ocr_text = ""
                    st.session_state["ocr_text"] = ocr_text
                # OCR 结果放进可编辑文本框，用户可修正识别错误后再解析
                raw_text = st.text_area(
                    "识别出的流水文字（可修改后解析）",
                    height=260,
                    key="ocr_text",
                )
                st.caption("图片识别可能有误，建议核对修改后再点击解析。")
            else:
                # 文本文件：直接读取全部内容
                raw_text = read_uploaded_file(uploaded_file)
                st.caption(f"已读取文件：{uploaded_file.name}（{len(raw_text)} 字符）")

    # ---------- 解析按钮 ----------
    if st.button("🚀 开始解析", type="primary", use_container_width=True):
        # 限制 1：API Key 未填写时直接提示，不发起请求
        if not api_key or not api_key.strip():
            st.session_state["flash"] = ("error", "请先在左侧边栏填写 DeepSeek API Key。")
        # 限制 2：输入为空时直接提示用户输入流水
        elif not raw_text or not raw_text.strip():
            st.session_state["flash"] = ("warning", "输入内容为空：请先粘贴流水文本或上传文件。")
        else:
            # 加载动画：网络请求期间显示"解析中，请稍候"
            with st.spinner("解析中，请稍候…"):
                try:
                    records = parse_bill_text(api_key.strip(), model, raw_text.strip())
                except BillParseError as e:
                    # 友好报错：只展示提示语，不抛堆栈
                    st.session_state["flash"] = ("error", str(e))
                else:
                    if records:
                        st.session_state["records"] = records
                        st.session_state["flash"] = (
                            "success",
                            f"解析成功，共 {len(records)} 笔记录。",
                        )
                    else:
                        # 模型返回空记录列表也算异常情况，提示用户检查输入
                        st.session_state.pop("records", None)
                        st.session_state["flash"] = (
                            "warning",
                            "未解析出任何账单记录，请检查流水内容是否正确。",
                        )

    # ---------- 渲染提示消息（成功 / 警告 / 错误，展示一次后清除） ----------
    flash = st.session_state.pop("flash", None)
    if flash is not None:
        kind, message = flash
        if kind == "success":
            st.success(message)
        elif kind == "warning":
            st.warning(message)
        else:
            st.error(message)

    # ---------- 结果区域：表格展示 + CSV 导出 ----------
    records = st.session_state.get("records", [])
    if records:
        st.subheader("📋 解析结果")
        df = pd.DataFrame(records)
        # 统一列顺序与中文表头（type 字段按需求保持原名）
        display_df = df[["time", "type", "category", "amount", "merchant", "remark"]].rename(
            columns={
                "time": "时间",
                "type": "type",
                "category": "分类",
                "amount": "金额",
                "merchant": "商家",
                "remark": "备注",
            }
        )
        st.dataframe(display_df, use_container_width=True, hide_index=True)

        # 导出 CSV：utf-8-sig 带 BOM，Excel 打开不乱码
        csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
        file_name = datetime.datetime.now().strftime("账单记录_%Y%m%d_%H%M.csv")
        st.download_button(
            "⬇️ 导出 CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="text/csv",
            type="primary",
        )


if __name__ == "__main__":
    main()
