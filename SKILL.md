---
name: format-agent
description: 通用文档格式排版。当用户给一份"格式规范"（自然语言规范文字，或排好版的 Word 模板）加一份格式混乱的 docx，要求按规范重排时使用。AI 只产出两个 JSON（格式规则 FormatSpec、段落角色 RoleMap），所有文档修改由确定性代码完成，输出可继续编辑的命名样式文档、修订模式文档和修改对照报告。
---

# 格式排版 Agent

把格式混乱的 Word 文档（.docx）按给定规范自动重排。**理解归 AI，动手归代码，中间用 JSON ⋅交接。**

## 能力边界

- 支持两种格式来源：自然语言规范文字（.txt）、排版正确的 Word 模板（.docx）
- **任何"已经排好版的文档"都可以当模板**：党委会/股东会/董事会材料、合同、
  法律意见书、咨询报告、论文、新闻稿……Agent 会读它的字体/字号/行距/缩进/编号体例，
  套用到待处理文档上
- 支持无结构文档（靠模型语义判断段落角色）和有序号文档（一、（一）、1. 手工编号、
  第一章/第一条 条款编号，或 Word 自动编号，确定性识别）
- 页面级：页边距、行网格、页眉页脚（含页码域）
- 段落级：字体/字号/加粗/对齐/行距/首行缩进/编号体例
- 表格级：首行表头（加粗/居中）、单元格字体字号对齐、边框线
- 模板未规定的角色自动与正文保持一致
- 正文文字内容永不改动
- 当前版本不重排：封面页设计、目录、多栏排版；合并单元格的复杂表格结构保持原样

## 环境要求

- Python 3.10+，依赖：`pip install -r requirements.txt`（python-docx、requests、PyMuPDF；Windows 渲染另需 pywin32）
- **在 Agent 环境里无需配置任何模型**：理解工作由宿主 Agent 完成（见"用法一"）
- 独立 API 模式（无 Agent 的环境）才需要模型配置：
  环境变量 / `.env` 文件：`LLM_BASE_URL`、`LLM_API_KEY`、`LLM_MODEL`（可选 `LLM_VISION_MODEL`、`LLM_TIMEOUT`、`LLM_TEMPERATURE`）

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

本 skill 有两种用法。**在 Agent 环境里优先用"内置智能模式"——零配置，
不需要任何 API key，理解工作由你（宿主 Agent）自己完成。**

### 用法一：Agent 内置智能模式（默认推荐，零配置）

你是宿主 Agent 时的执行流程：

1. **抽取目标文档结构**：

   ```bash
   python main.py --extract-only --target 待排版.docx --out 输出目录/排版后.docx
   ```

   得到 `输出目录/排版后_paragraphs.json`，每段含
   `idx/text/size_pt/bold/alignment/style_name/in_table` 及编号元数据。

2. **你自己产出 FormatSpec**（格式规则 JSON）：
   - 用户给的是**规范文字**：读懂它，按下面的 schema 写出 JSON，保存为
     `输出目录/排版后_formatspec.json`；
   - 用户给的是**模板 docx**：先对模板跑一次 `--extract-only`，读懂各段角色
     （标题/正文/落款……），写入 `模板_rolemap.json`（格式
     `{"0": "title", "1": "body", ...}`），稍后用 `--template 模板.docx
     --template-rolemap-json 模板_rolemap.json` 让代码确定性读取模板格式。

   FormatSpec schema：

   ```json
   {
     "page": {"margin": {"top_mm": 37, "bottom_mm": 35, "left_mm": 28, "right_mm": 26},
              "line_grid": {"line_pt": 28}},
     "roles": {
       "body": {"font_eastasia": "仿宋_GB2312", "font_ascii": "Times New Roman",
                "size_pt": 16, "bold": false, "alignment": "justify",
                "first_line_indent_chars": 2, "line_spacing": {"type": "exact", "pt": 28}}
     }
   }
   ```

   角色枚举：`title/subtitle/heading_1/heading_2/heading_3/body/signature/date/
   attachment_label/attachment/other`。`body` 必填；对齐取值 `left/center/right/justify`；
   `size_pt` 8~72，页边距 5~50mm。规范里没提的角色不要写，没提的字段不要编。

3. **你自己产出 RoleMap**（段落角色 JSON）：读第 1 步的段落清单，
   为每个 `in_table=false` 的段落从枚举里选角色，保存为
   `输出目录/排版后_rolemap.json`，格式 `{"0": "title", "1": "body", ...}`，
   必须覆盖所有非表格段落。判断依据：文字内容、位置顺序、当前格式提示；
   落款在末尾署名感强；日期含"年/月/日"；标题在最前且独立成行。

4. **执行排版**（确定性代码接管）：

   ```bash
   # 规范文字来源（FormatSpec 由你在第 2 步产出）
   python main.py --spec-json 输出目录/排版后_formatspec.json \
       --rolemap-json 输出目录/排版后_rolemap.json \
       --target 待排版.docx --out 输出目录/排版后.docx

   # 模板来源
   python main.py --template 模板.docx --template-rolemap-json 模板_rolemap.json \
       --rolemap-json 输出目录/排版后_rolemap.json \
       --target 待排版.docx --out 输出目录/排版后.docx
   ```

   你产出的 JSON 会被代码再次校验（角色合法、段落全覆盖、数值边界），
   校验失败会报错并说明原因，修正后重跑即可。

### 用法二：独立 API 模式（可选）

在没有 Agent 的环境（服务器、定时任务、本地裸跑）才需要配置模型：
复制 `.env.example` 为 `.env`，填 `LLM_BASE_URL / LLM_API_KEY / LLM_MODEL`
（建议多模态模型，见上节），然后：

```bash
# 规范文字作为格式来源（LLM 自动抽取规则）
python main.py --spec 规范文字.txt --target 待排版.docx --out 输出目录/排版后.docx

# Word 模板作为格式来源
python main.py --template 模板.docx --target 待排版.docx --out 输出目录/排版后.docx
```

模板可以是任何已按规范排好版的文档（党委会议题材料、合同范本、法律意见书、
咨询报告、论文范文……），无需为模板做任何改造或标注。

### 视觉自检（可选，仅用法二 + 多模态模型）

加 `--verify`：排版后渲染成图，由视觉模型对照规范质检，发现偏差定向修复并重排一次（只一轮，不做开放循环）。
用法一中若你（宿主 Agent）具备图像理解能力，也可以自己渲染后检查：输出目录 `*_verify_render/` 下是逐页 PNG。

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
