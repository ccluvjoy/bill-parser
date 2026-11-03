# 🧾 账单流水解析工具

粘贴微信 / 支付宝账单流水，或上传账单截图，调用 DeepSeek API 一键解析为结构化记录（时间、类型、分类、金额、商家、备注），表格展示，一键导出 CSV。

## ✨ 功能

- **双输入方式**：文本框大段粘贴 / 上传 `.txt` 等文本文件
- **截图直接解析**：上传 `.jpg` / `.jpeg` / `.png` 账单截图，自动 OCR 识别文字（Windows 内置引擎，无需安装第三方 OCR），识别结果可修改后再解析
- **智能分类**：自动归类为餐饮 / 交通 / 购物 / 房租住宿 / 学习资料 / 医药 / 娱乐 / 转账 / 红包 / 充值 / 其他，type 标记 income / expense / transfer
- **自动修复**：模型返回非法 JSON 时，自动调用修复 Prompt 重试一次
- **导出 CSV**：utf-8 带 BOM 编码，Excel 打开不乱码
- **友好报错**：网络失败、API Key 无效、超时等均给出中文提示，不抛堆栈

## 👥 适合人群

- 想记账但懒得逐条手动录入的学生、租房室友
- 需要快速整理微信 / 支付宝流水的人

## 🚀 部署启动

```bash
git clone <仓库地址>
cd bill-parser
pip install -r requirements.txt
streamlit run bill_parser_app.py
```

浏览器打开 **http://localhost:8501** 即可使用。

> 环境要求：Python 3.10+；图片 OCR 功能仅支持 Windows（文本解析跨平台）。

## 🔑 获取 DeepSeek API Key

1. 打开 [platform.deepseek.com](https://platform.deepseek.com)，注册并登录
2. 左侧栏 **API Keys** → **创建 API Key**，复制 `sk-` 开头的密钥
3. 回到工具页面，粘贴到左侧边栏的「DeepSeek API Key」

> API 按量计费，费用很低，新用户有赠送额度可先试用。

## 📖 使用教程

1. 左侧边栏填写 API Key，选择模型（`deepseek-chat` / `deepseek-reasoner`）
2. 粘贴流水文本，或切换到「上传文件」上传 `.txt` 或账单截图
3. 点击 **开始解析**，等待几秒
4. 核对结果表格（截图识别建议先核对 OCR 文字），点击 **导出 CSV** 下载

## ⚠️ 局限

- 图片 OCR 仅限 Windows，且需系统已安装中文语言包；截图越清晰识别越准
- 解析结果依赖大模型，偶尔可能分类不准，建议导出前核对
- API Key 仅保存在当前会话，刷新页面后需重新填写
- 超长流水（数千条）单次解析可能超时，建议分批粘贴

## 🔒 隐私提醒

流水文本会发送至 DeepSeek API 用于解析。**请勿上传**包含身份证号、银行卡号、住址等敏感内容的流水，必要时先脱敏处理。

## 📁 项目结构

```
bill-parser/
├── bill_parser_app.py   # 主程序（Prompt 集中在文件顶部常量区，便于修改）
├── requirements.txt     # 依赖清单
└── README.md
```
