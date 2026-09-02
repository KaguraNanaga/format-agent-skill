# format-agent-skill

**中文 | [English](README_EN.md)**

把格式规范和待排版 Word 文稿交给 Agent，自动完成结构识别、样式迁移、页面整理和质量检查，输出可继续编辑、可审计的 DOCX。

格式来源可以是：

- 一段自然语言规范；
- 一份排版正确的 Word 模板；
- APA、MLA、IEEE、Chicago、Turabian、公文、技术手册等内置 Style Pack；
- 人工准备的 `FormatSpec` / `RoleMap` JSON。

> 本仓库就是 Skill 本体：`SKILL.md` 位于根目录，clone 或下载整个仓库即可安装。

## 它能做什么

- 处理中英文标题、正文、编号、字体字号、行距缩进、分页和命名样式。
- 迁移纸张、方向、页边距、分栏、页码以及首页/奇偶页页眉页脚。
- 识别手工编号与 Word 自动编号，保留模板未覆盖的正文内容。
- 调整常规表格的宽度、列宽、单元格边距、重复表头、跨页断行和垂直对齐。
- 扫描正文、嵌套表格、页眉页脚、脚注尾注、批注、文本框、域、书签、内容控件、修订和嵌入对象。
- 生成主稿、属性修订稿、修改对照报告、预检报告和结构化中间结果。
- 可选视觉自检：把排版结果渲染成页面图，由视觉模型再检查一轮。
- 支持 `.doc/.wps/.odt/.rtf` 输入转换，统一输出 `.docx`。

## 适合的文稿

| 场景 | 已有能力 |
|---|---|
| 讲话稿、理论学习材料、汇报材料、普通报告 | 标题层级、正文、编号、页面、页眉页脚和常规表格 |
| 通知、请示、报告、函、会议纪要 | 可读取机关模板，也提供 `official-cn-gbt9704` 公文排版基线 |
| 中文论文、毕业论文 | 摘要、关键词、标题层级、参考文献格式、封面字段、目录、分节、动态页眉和页码 |
| 英文 memo、essay、report、academic paper | Abstract、Keywords、Chapter/Section、block quote、running head、figure/table caption、References/Works Cited |
| APA 7、MLA 9、IEEE、Chicago 18、Turabian 9 | 内置可审计 Style Pack，也可以直接读取学校、课程、期刊的 Word 模板 |
| 技术手册、产品说明 | 代码与命令、步骤、WARNING/CAUTION/NOTE/TIP、图形与相邻题注分页绑定 |
| 美国法律 brief | court caption、案号、标题、TOC/TOA、律师信息、certificate，并可按显式配置创建 TA/TOA 域 |
| 合同、法律意见书、招标和长篇报告 | Article/Section 层级、普通表格、多节页面和附件标题的保守格式迁移 |

正式机关、学校、法院、客户或期刊提供的模板优先于通用 Style Pack。

## 示例

| 场景 | 改前 | 改后 |
|---|---|---|
| 党委会议题议案 | ![前](docs/images/case4-党委议案-改前.png) | ![后](docs/images/case4-党委议案-改后.png) |
| 课程论文 | ![前](docs/images/case3-论文-改前.png) | ![后](docs/images/case3-论文-改后.png) |
| 英文 board resolution | ![before](docs/images/en-board-before.png) | ![after](docs/images/en-board-after.png) |

## 安装

对支持 Skill 的 Agent 说：

```text
帮我安装这个 skill：https://github.com/KaguraNanaga/format-agent-skill
```

也可以手动克隆：

```bash
gh repo clone KaguraNanaga/format-agent-skill
pip install -r format-agent-skill/requirements.txt
```

宿主需要能够读取本地文件并运行 Python。旧格式转换、域刷新和页面渲染会按任务需要调用本机 Word、WPS 或 LibreOffice。

## 快速使用

先对陌生文档做一次预检：

```bash
python main.py --preflight-only --target source.docx --out output/result.docx
```

然后选择一种格式来源：

```bash
# 自然语言规范
python main.py --spec rules.txt --target source.docx --out output/result.docx

# Word 模板
python main.py --template template.docx --target source.docx --out output/result.docx

# 内置 Style Pack
python main.py --style-pack apa7-student --target paper.docx --out output/paper.docx

# 完全跳过模型
python main.py --spec-json examples/spec_std.json \
  --rolemap-json examples/rolemap_std.json \
  --target examples/messy.docx --out output/result.docx
```

常用 Style Pack：

```bash
python main.py --style-pack mla9 --running-head Rivera --target paper.docx --out output/mla.docx
python main.py --style-pack ieee-journal --target article.docx --out output/ieee.docx
python main.py --style-pack chicago18-notes-bibliography --target paper.docx --out output/chicago.docx
python main.py --style-pack turabian9-student --target thesis.docx --out output/turabian.docx
python main.py --style-pack official-cn-gbt9704 --target notice.docx --out output/notice.docx
python main.py --style-pack technical-manual --target manual.docx --out output/manual.docx
python main.py --style-pack us-legal-brief --target brief.docx --out output/brief.docx
```

需要视觉复核时加 `--verify`；需要在 Windows Microsoft Word 中刷新 TOC、PAGE、REF、SEQ、TA/TOA 等域时加 `--refresh-fields`。更多命令、退出码与排错说明见 [references/cli-and-output.md](references/cli-and-output.md)。

## 工作原理

模型负责理解，代码负责执行。排版主链路使用两个 JSON 契约：

```text
格式来源（规范文字 / Style Pack / Word 模板）
        │
        ▼
   FormatSpec（格式规则 JSON）◄── Schema 校验
        │
目标文稿 ── 全 Story / 结构抽取 ──► 段落清单
        │                              │
        │      确定性识别 + 语义兜底   ▼
        │                         RoleMap（角色 JSON）
        ▼                              │
  确定性执行器（命名样式 / OOXML）◄───┘
        │
        ▼
主稿 + 属性修订稿 + 对照报告 + 中间 JSON
```

- 模型不直接编辑 DOCX，实际修改由确定性代码完成。
- 文档先经过全 Story 和复杂结构预检，再进入排版。
- 主要段落格式写入 Word 命名样式，方便继续编辑。
- 正式产物先写候选文件，文字完整性校验通过后再原子提交。
- 手工编号转换、摘要/关键词标签结构化等必要变化会进入审计白名单和修改报告。

## 输出产物

| 产物 | 说明 |
|---|---|
| `result.docx` | 可继续编辑的排版主稿 |
| `result_tracked.docx` | 记录段落和字符属性变化的修订稿 |
| `result_report.docx` / `.md` | 页面、结构和段落修改对照报告 |
| `result_formatspec.json` / `_rolemap.json` / `_stylemap.json` | Agent 对规范和结构的理解结果 |
| `result_preflight.json` | Story、分节、对象、警告和阻断项 |
| `result_issues.json` | 视觉复核发现问题时生成 |

## 输入与运行环境

- `.docx` 原生处理。
- `.doc/.wps` 可通过本机 Word/WPS 转换。
- `.odt/.rtf` 优先通过 LibreOffice，也可尝试兼容的 Word/WPS 导入器。
- 正式输出统一为 `.docx`。
- Python 3.10+；建议使用独立虚拟环境安装 `requirements.txt`。
- Word/WPS 转换和渲染在独立限时进程中运行，单个 Office 服务无响应不会拖住整个任务。
- `--spec-json` + `--rolemap-json` 可以完全不调用模型；只有 `--verify` 需要支持图像输入的模型。

## 使用说明

当前版本聚焦“段落 + 常规表格”型 Word 文稿。简历、宣传册、海报、复杂表单等高度依赖文本框和绝对定位的版面，会以保留和预检为主；PDF/扫描件不走 Word 格式迁移链路。

引文、参考文献、法律引证和索引内容会保留，但不会自动改写或校正。复杂浮动对象、内容控件和嵌入对象也以安全保留为主。涉及大规模分节重建、双栏或横向宽表时，系统会先预检并要求显式授权。

旧格式转换依赖本机 Office 组件，可能出现字体或分页差异；ODT 推荐安装 LibreOffice。更细的论文、公文、技术手册、法律 brief 和旧格式说明分别见：

- [论文模式](references/thesis-mode.md)
- [Style Pack](references/style-packs.md)
- [公文、技术、法律和旧格式](references/advanced-modes.md)

## 真实环境验证

2026-09-02 在 Windows + Microsoft Word + WPS 环境完成端到端验证：

- 原生 DOCX、真实二进制 DOC、真实 RTF 和 WPS 路由完成转换、排版和文字完整性检查。
- 5 份文档共 10 页逐页检查，无缺字、裁切、重叠、表格错位或分页漂移。
- 回归测试：`42 passed`。

## 相关项目

- 黑客松完整版（GUI、案例和路演页）：[format-agent](https://github.com/KaguraNanaga/format-agent)（已归档）
