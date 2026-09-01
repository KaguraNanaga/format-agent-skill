---
name: format-agent
description: 按自然语言规范、内置 Style Pack 或 Word 模板，重排以段落、标题和常规表格为主的中文/英文 Word 文稿；覆盖公文、论文、Chicago/Turabian、基础技术手册和美国法律 brief/TOA。可将 .doc/.wps/.odt/.rtf 有损转换为临时 DOCX；拒绝 PDF 版面猜测。不要用于简历、表单、宣传册、Newsletter、海报或内容控件应用；这些文档只可预检和保留。
---

# 格式排版 Agent

把格式混乱的 Word 文档（.docx）按给定规范自动重排。**理解归 AI，动手归代码，中间用 JSON 交接。**

## 能力边界

- 支持两种格式来源：自然语言规范文字（.txt）、排版正确的 Word 模板（.docx）
- 支持无结构文档（靠模型语义判断段落角色）和有序号文档（一、（一）、1. 手工编号或 Word 自动编号，确定性识别）
- 页面级：纸张、横竖方向、页边距、行网格、分节覆盖，以及首页/奇数页/偶数页独立页眉页脚；段落级：字体/字号/粗斜体/下划线/颜色/对齐/行距/首行或悬挂缩进/编号体例/分页约束
- 执行前同时扫描目标稿和参考模板的正文、各页眉页脚、脚注、尾注、批注、文本框、内容控件、修订、域、嵌入对象、宏、签名、文档保护和全部分节
- 表格支持对齐、宽度、列宽、固定/自动布局、单元格边距与垂直对齐、重复表头、跨页断行和显式点名的横向分节；不按正文 RoleMap 猜测表格语义
- 脚注/尾注可在 `notes` 明确给出规则后调整字体、字号、对齐、行距与缩进；题注可显式转换为 SEQ 域，已有题注标题可插入真实图目录/表目录域
- 全文盘点 CITATION、BIBLIOGRAPHY、REF、PAGEREF、NOTEREF、SEQ、TOC、PAGE、STYLEREF 等域及书签，报告失效交叉引用目标；只保留和刷新域，不改写引文或参考文献内容
- 正文、脚注、尾注、批注和文本框文字必须通过完整性校验
- 模板未规定的角色自动与正文保持一致
- 多节模板按节索引迁移纸张、方向、页边距、栏数、页码格式和独立页眉页脚；目标稿本身已有多节而规范只给全局页边距时默认保留各节特例，避免覆盖横向宽表等布局
- 自动选择这些专用 Profile：
  - `official_cn`：发文字号、主送机关、公文标题层级、结束语、落款、日期和抄送
  - `english_general`：报告、备忘录、信函等一般英文文稿
  - `english_academic`：识别 Abstract、Keywords、Chapter/Part/Section、编号标题、Figure/Table/Equation caption、块引用、代码块、作者与单位、Author Note、References/Bibliography/Works Cited
  - `english_legal`：识别 ARTICLE、Section、定义条款、多级字母/罗马数字条款、签署块、Exhibit/Schedule/Annex 和 Table of Authorities 标题，并默认保留粗斜体、小型大写等强调
  - `english_technical`：识别步骤、命令/代码、WARNING/CAUTION/NOTE/TIP，并绑定现有图形与相邻题注的分页关系
  - `english_legal_brief`：识别 court caption、案号、brief 标题、TOC/TOA、律师信息与 certificate；TA/TOA 只按显式配置创建
- 可显式选择公文、APA/MLA/IEEE、Chicago/Turabian、技术手册和法律 brief Style Pack；正式模板与主管机关/学校/法院规则始终优先。按任务阅读 [references/style-packs.md](references/style-packs.md) 与 [references/advanced-modes.md](references/advanced-modes.md)
- 论文模板出现“摘要 + 关键词 + 参考文献”等高置信信号时启用 thesis profile：
  - `strict` 清掉源文档的异常斜体、下划线、颜色、高亮、删除线等直接格式
  - 章标题和参考文献标题另页起排，标题与后文保持同页
  - `摘要：`、`关键词：`只加粗前缀；参考文献条目使用悬挂缩进
  - 从模板提取校名/论文类型和校徽，安全重建封面；未知元数据明确标为“（待填写）”
  - 重建摘要、目录、正文、参考文献分节；前置页使用罗马页码，正文从阿拉伯数字 1 重新编号
  - 目录使用真实 TOC 域，正文页眉使用 STYLEREF 随章标题变化，页脚使用 PAGE 域
- 不盲拷贝模板中的示例作者、英文题名或法律声明。声明页默认关闭

### 安全边界

- 原生处理标准 `.docx`；`.doc/.wps` 通过本机 Word/WPS、`.odt/.rtf` 优先通过 LibreOffice 转成临时 DOCX，再重新预检。转换有损且只输出 `.docx`；PDF 和 `.docm` 不进入该链路
- 正文与各级嵌套表格按真实 OOXML 阅读顺序抽取；页眉页脚、脚注尾注、批注、文本框、内容控件和修订文字作为只读 Story 纳入预检与完整性审计，不参与正文角色排版
- 简历、宣传册、Newsletter、海报、复杂申请表和可填写合同模板通常依赖文本框、浮动对象或内容控件；本技能不得把它们当作普通段落文稿自动重排
- 脚注/尾注只有在 FormatSpec 明确提供 `notes.footnote` / `notes.endnote` 时才调整；不会增删注释、改变注号或重写注释文字
- 引用域、索引域和书签只做盘点/保留/检查/刷新。法律模式可在用户提供精确 `citation_marks` 时写 TA 域，并可在已有 TOA 标题后插入 TOA 域；不会自行发现、生成或校正引证，也不会按任何体例改写引文与参考文献
- 自动题注只处理单段、可确定识别的简单整数编号；章节式编号、跨 Run 编号和不确定题注保留原样并写诊断
- 源稿含未接受修订、`altChunk`、编辑保护、数字签名或宏时默认阻断，不生成终稿
- 已有多节文档若还要求重建论文结构或分栏，默认阻断；只有用户理解分节可能被重构并明确授权时，才可加 `--allow-risky-structure`
- `table.landscape_table_indices` 会新增分节符，必须先预检并显式允许结构调整；索引从 0 开始，且只操作点名的顶层表格
- 输出路径不得与源稿相同。所有 DOCX 和报告先写入同目录候选文件，文本完整性通过后才原子替换最终路径；失败时旧终稿保持不变

## 环境要求

- Python 3.10+，依赖：`pip install -r requirements.txt`（python-docx、requests、PyMuPDF；Windows 渲染另需 pywin32）
- 模型配置（二选一）：
  - 环境变量 / `.env` 文件：`LLM_BASE_URL`、`LLM_API_KEY`、`LLM_MODEL`（可选 `LLM_VISION_MODEL`、`LLM_TIMEOUT`、`LLM_TEMPERATURE`）
  - 或完全不用模型：直接提供 `spec_std.json` 与 `rolemap_std.json` 走确定性降级链路

## 模型选择

本技能不绑定特定模型，使用 OpenAI 兼容端点：

- 排版主流程（规则抽取、角色标注）只需文本能力
- 仅在用户要求视觉自检或命令包含 `--verify` 时，才需要支持图像输入的多模态模型和文档渲染环境
- 未启用 `--verify` 时，不要因模型不支持图像而暂停主流程

## 使用方法

在本 skill 目录下运行：

### 先做能力预检（推荐用于陌生文档）

```bash
python main.py --preflight-only --target 待排版.docx --out 输出目录/排版后.docx
```

如同时使用 Word 模板，把 `--template 模板.docx` 加到同一命令，预检会分别标记 `target` 和 `template` 风险。命令生成 `排版后_preflight.json`。返回码 `0` 表示没有硬阻断；`2` 表示必须先处理修订、保护、签名等问题。

### 规范文字 → 排版

```bash
python main.py --spec 规范文字.txt --target 待排版.docx --out 输出目录/排版后.docx
```

### Word 模板 → 排版

```bash
python main.py --template 模板.docx --target 待排版.docx --out 输出目录/排版后.docx
```

### 内置 Style Pack → 排版

```bash
python main.py --style-pack apa7-student --target paper.docx --out output/paper_apa.docx
python main.py --style-pack mla9 --running-head Rivera --target paper.docx --out output/paper_mla.docx
python main.py --style-pack ieee-journal --target article.docx --out output/article_ieee.docx
python main.py --style-pack official-cn-gbt9704 --target notice.docx --out output/notice.docx
python main.py --style-pack chicago18-notes-bibliography --target paper.docx --out output/paper_chicago.docx
python main.py --style-pack turabian9-student --target thesis.docx --out output/thesis_turabian.docx
python main.py --style-pack technical-manual --target manual.odt --out output/manual.docx
python main.py --style-pack us-legal-brief --target brief.doc --out output/brief.docx
```

Style Pack 是可审计的通用排版基线，不代替机关、学校、课程、法院或期刊的正式模板，也不校正内容。差异和来源见 [references/style-packs.md](references/style-packs.md)。技术、法律和输入转换配置见 [references/advanced-modes.md](references/advanced-modes.md)。

论文模板会自动选择严格清洗；也可显式覆盖：

```bash
python main.py --template 论文模板.docx --target 待排版.docx --out 输出目录/论文.docx \
    --cleanup-mode strict --refresh-fields
```

三种清洗策略：`controlled` 仅清理规则控制字段；`strict` 严格克隆模板样式；
`preserve_emphasis` 清理颜色、下划线等噪声，但保留作者的粗体/斜体强调。
英文 academic/legal/technical/brief 默认使用 `preserve_emphasis`。

`--refresh-fields` 会在 Windows 上调用本机 Microsoft Word，把目录、动态页眉和页码
刷新后落盘；不加该参数时，文档也会设置为下次在 Word 打开时自动更新域。

### 无模型降级（全部跳过 LLM）

```bash
python main.py --spec-json examples/spec_std.json --rolemap-json examples/rolemap_std.json \
    --target examples/messy.docx --out 输出目录/排版后.docx
```

### 视觉自检（可选，需要视觉模型）

加 `--verify`：排版后渲染成图，由视觉模型对照规范质检，发现偏差定向修复并重排一次（只一轮，不做开放循环）。

## 输出产物

每次运行产出（同目录同名前缀）：

- `排版后.docx` —— 命名样式写入的干净稿
- `排版后_tracked.docx` —— **段落/字符属性修订稿**：可审阅 `pPrChange/rPrChange`；页面、分节、页眉页脚、表格、封面、目录和结构移动不属于可逐条接受/拒绝的修订
- `排版后_report.docx` / `.md` —— 修改对照报告（页面设置、规则、段落明细）
- `排版后_formatspec.json` / `_rolemap.json` / `_stylemap.json` —— 中间产物（给评审看的"理解证据"）
- `排版后_preflight.json` —— Story、分节、复杂对象及阻断/警告清单

上述正式产物只在排版、语义内容完整性校验全部成功后提交。允许且必须审计的变化仅包括手工编号转自动编号、摘要标签结构化和明确声明的目录/封面字段；其他文字差异会使任务失败。文本完整性失败返回码为 `3`，并仅写出 `_integrity_failure.json` 诊断，不覆盖已有终稿。

## 排错速查

- `缺少 LLM 配置` → 检查 .env 或环境变量
- `.doc/.wps/.odt/.rtf` 转换失败 → 安装 Word/WPS 或 LibreOffice；不要把 PDF 重命名成 Word 文件
- 角色标注反复失败 → 模型 JSON 输出不稳定，改用 `--rolemap-json` 手工标注
- 能力预检返回 2 → 先查看 `_preflight.json`；未接受修订、保护、签名、宏必须在 Word 中处理后重试
- 多节结构重建被阻断 → 优先取消模板的结构重建/分栏要求；确需重构且已备份时才加 `--allow-risky-structure`
- 文本完整性返回 3 → 查看 `_integrity_failure.json`，旧终稿未被覆盖；不要把候选临时文件当成交付稿
- 视觉复核 400 temperature → 该模型只许 temperature= 1，在 .env 设 `LLM_TEMPERATURE= 1`
- 视觉复核报图片错误 → 端点不支持 data:base64 内联图片，或当前模型无图像输入能力；换多模态模型
- 渲染失败 → Windows 需装 Word（走 COM），macOS/Linux 需 LibreOffice
- 目录仍显示空白/旧页码 → Windows 加 `--refresh-fields` 重跑，或在 Word 中全选后按 F9

## examples/ 目录

- `messy.docx`：格式全乱的示例公文（通知）
- `spec.txt`：公文排版规范文字示例
- `spec_std.json` / `rolemap_std.json`：上述示例的标准答案（验收基准 + 降级输入）
- `模板-党委会议题样表.docx`：模板模式的格式来源示例
