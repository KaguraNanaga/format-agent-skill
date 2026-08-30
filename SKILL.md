---
name: format-agent
description: 通用文档格式排版。当用户给一份"格式规范"（自然语言规范文字，或排好版的 Word 模板）加一份格式混乱的 docx，要求按规范重排时使用。AI 只产出两个 JSON（格式规则 FormatSpec、段落角色 RoleMap），所有文档修改由确定性代码完成，输出可继续编辑的命名样式文档、修订模式文档和修改对照报告。
---

# 格式排版 Agent

把格式混乱的 Word 文档（.docx）按给定规范自动重排。**理解归 AI，动手归代码，中间用 JSON ⋅交接。**

## 能力边界

- 支持两种格式来源：自然语言规范文字（.txt）、排版正确的 Word 模板（.docx）
- 支持无结构文档（靠模型语义判断段落角色）和有序号文档（一、（一）、1. 手工编号或 Word 自动编号，确定性识别）
- 页面级：页边距、行网格；段落级：字体/字号/加粗/对齐/行距/首行缩进/编号体例
- 表格内段落不重排；正文文字内容永不改动
- 模板未规定的角色自动与正文保持一致

## 环境要求

- Python 3.10+，依赖：`pip install -r requirements.txt`（python-docx、requests、PyMuPDF；Windows 渲染另需 pywin32）
- 模型配置（二选一）：
  - 环境变量 / `.env` 文件：`LLM_BASE_URL`、`LLM_API_KEY`、`LLM_MODEL`（可选 `LLM_VISION_MODEL`、`LLM_TIMEOUT`、`LLM_TEMPERATURE`）
  - 或完全不用模型：直接提供 `spec_std.json` 与 `rolemap_std.json` 走确定性降级链路

## 模型选择（重要：建议多模态）

本技能不绑定任何特定模型，任何 OpenAI 兼容端点均可。**但建议配置多模态模型**
（同时支持文本与图像输入，如 GPT-4o、Kimi K3、Qwen-VL、GLM-4V 等）：

- 排版主流程（规则抽取、角色标注）只需文本能力
- **视觉自检（--verify）必须把渲染图交给模型质检**，纯文本模型无法使用这一环节

**给调用方 Agent 的指示**：在开始执行前，先确认当前使用的模型是什么、
是否具备多模态（图像输入）能力。如果当前模型不具备多模态能力，请主动提示用户：
"当前模型不支持图像输入，视觉自检将不可用；建议切换为多模态模型（如 XXX），
可以获得排版后的自动质检，效果更佳。" 用户确认后再继续执行。

## 使用方法

在本 skill 目录下运行：

### 规范文字 → 排版

```bash
python main.py --spec 规范文字.txt --target 待排版.docx --out 输出目录/排版后.docx
```

### Word 模板 → 排版

```bash
python main.py --template 模板.docx --target 待排版.docx --out 输出目录/排版后.docx
```

### 无模型降级（全部跳过 LLM）

```bash
python main.py --spec-json examples/spec_std.json --rolemap-json examples/rolemap_std.json \
    --target examples/messy.docx --out 输出目录/排版后.docx
```

### 视觉自检（可选，需要多模态模型）

加 `--verify`：排版后渲染成图，由视觉模型对照规范质检，发现偏差定向修复并重排一次（只一轮，不做开放循环）。

## 输出产物

每次运行产出（同目录同名前缀）：

- `排版后.docx` —— 命名样式写入的干净稿
- `排版后_tracked.docx` —— **修订模式**：Word 审阅视图可见每处格式改动，可逐条接受/拒绝
- `排版后_report.docx` / `.md` —— 修改对照报告（页面设置、规则、段落明细）
- `排版后_formatspec.json` / `_rolemap.json` / `_stylemap.json` —— 中间产物（给评审看的"理解证据"）

## 排错速查

- `缺少 LLM 配置` → 检查 .env 或环境变量
- 角色标注反复失败 → 模型 JSON 输出不稳定，改用 `--rolemap-json` 手工标注
- 视觉复核 400 temperature → 该模型只许 temperature= 1，在 .env 设 `LLM_TEMPERATURE= 1`
- 视觉复核报图片错误 → 端点不支持 data:base64 内联图片，或当前模型无图像输入能力；换多模态模型
- 渲染失败 → Windows 需装 Word（走 COM），macOS/Linux 需 LibreOffice

## examples/ 目录

- `messy.docx`：格式全乱的示例公文（通知）
- `spec.txt`：公文排版规范文字示例
- `spec_std.json` / `rolemap_std.json`：上述示例的标准答案（验收基准 + 降级输入）
- `模板-党委会议题样表.docx`：模板模式的格式来源示例
