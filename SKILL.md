---
name: format-agent
description: 通用文档与论文格式排版。当用户给一份格式规范（自然语言规范文字，或排好版的 Word 模板）加一份格式混乱的 docx，要求按规范重排时使用。支持论文摘要、关键词、章标题、参考文献、公式和附录语义；AI 只产出 FormatSpec 与 RoleMap，所有 Word 修改由确定性代码完成。
---

# 格式排版 Agent

把格式混乱的 Word 文档（.docx）按给定规范自动重排。**理解归 AI，动手归代码，中间用 JSON ⋅交接。**

## 能力边界

- 支持两种格式来源：自然语言规范文字（.txt）、排版正确的 Word 模板（.docx）
- 支持无结构文档（靠模型语义判断段落角色）和有序号文档（一、（一）、1. 手工编号或 Word 自动编号，确定性识别）
- 页面级：页边距、行网格；段落级：字体/字号/粗斜体/下划线/颜色/对齐/行距/首行或悬挂缩进/编号体例/分页约束
- 表格内段落不重排；正文文字内容永不改动
- 模板未规定的角色自动与正文保持一致
- 论文模板出现“摘要 + 关键词 + 参考文献”等高置信信号时启用 thesis profile：
  - `strict` 清掉源文档的异常斜体、下划线、颜色、高亮、删除线等直接格式
  - 章标题和参考文献标题另页起排，标题与后文保持同页
  - `摘要：`、`关键词：`只加粗前缀；参考文献条目使用悬挂缩进
  - 从模板提取校名/论文类型和校徽，安全重建封面；未知元数据明确标为“（待填写）”
  - 重建摘要、目录、正文、参考文献分节；前置页使用罗马页码，正文从阿拉伯数字 1 重新编号
  - 目录使用真实 TOC 域，正文页眉使用 STYLEREF 随章标题变化，页脚使用 PAGE 域
- 不盲拷贝模板中的示例作者、英文题名或法律声明。声明页默认关闭；奇偶页不同页眉仍不在当前范围

## 环境要求

- Python 3.10+，依赖：`pip install -r requirements.txt`（python-docx、requests、PyMuPDF、streamlit；Windows 渲染另需 pywin32）
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

论文模板会自动选择严格清洗；也可显式覆盖：

```bash
python main.py --template 论文模板.docx --target 待排版.docx --out 输出目录/论文.docx \
    --cleanup-mode strict --refresh-fields
```

三种清洗策略：`controlled` 仅清理规则控制字段；`strict` 严格克隆模板样式；
`preserve_emphasis` 清理颜色、下划线等噪声，但保留作者的粗体/斜体强调。

`--refresh-fields` 会在 Windows 上调用本机 Microsoft Word，把目录、动态页眉和页码
刷新后落盘；不加该参数时，文档也会设置为下次在 Word 打开时自动更新域。

### 无模型降级（全部跳过 LLM）

```bash
python main.py --spec-json examples/spec_std.json --rolemap-json examples/rolemap_std.json \
    --target examples/messy.docx --out 输出目录/排版后.docx
```

### 视觉自检（可选，需要视觉模型）

加 `--verify`：排版后渲染成图，由视觉模型对照规范质检，发现偏差定向修复并重排一次（只一轮，不做开放循环）。

### 演示界面

```bash
streamlit run app.py   # app.py 在项目主仓库，本 skill 目录只含流水线
```

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
- 目录仍显示空白/旧页码 → Windows 加 `--refresh-fields` 重跑，或在 Word 中全选后按 F9

## examples/ 目录

- `messy.docx`：格式全乱的示例公文（通知）
- `spec.txt`：公文排版规范文字示例
- `spec_std.json` / `rolemap_std.json`：上述示例的标准答案（验收基准 + 降级输入）
- `模板-党委会议题样表.docx`：模板模式的格式来源示例
