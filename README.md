# format-agent-skill

**中文 | [English](README_EN.md)**

一个用于中文/英文 Word 文稿的格式排版 Agent Skill。它接收自然语言规范、内置 Style Pack 或 Word 模板，识别文档结构后由确定性代码修改 OOXML，输出可继续编辑、可审计的 DOCX。

> 本仓库就是 Skill 本体：`SKILL.md` 位于根目录，clone 或下载整个仓库即可安装。

它适合以段落和常规表格为主的文稿，不是通用桌面出版、OCR 或格式合规认证工具。机关、学校、法院、客户和期刊的正式模板及书面规范始终优先于内置基线。

## 安装

可以直接对支持 Skill 的 Agent 说：

```text
帮我安装这个 skill：https://github.com/KaguraNanaga/format-agent-skill
```

也可以克隆仓库，再把整个目录放入宿主的技能目录：

```bash
gh repo clone KaguraNanaga/format-agent-skill
pip install -r format-agent-skill/requirements.txt
```

宿主需要能够读取本地文件、运行 Python，并在安装依赖或调用 Office/WPS 前取得相应权限。不同 Agent 产品对 Skill 协议、命令执行和桌面应用调用的支持不同，安装成功不等于所有高级功能都可用。

## 快速使用

先对陌生文档做能力预检：

```bash
python main.py --preflight-only --target source.docx --out output/result.docx
```

选择一种格式来源：

```bash
# 自然语言规范
python main.py --spec rules.txt --target source.docx --out output/result.docx

# Word 模板
python main.py --template template.docx --target source.docx --out output/result.docx

# 内置排版基线
python main.py --style-pack apa7-student --target paper.docx --out output/paper.docx

# 完全跳过模型
python main.py --spec-json examples/spec_std.json \
  --rolemap-json examples/rolemap_std.json \
  --target examples/messy.docx --out output/result.docx
```

只有需要视觉复核时才加 `--verify`；它还需要可用的 Word/WPS/LibreOffice 渲染环境和支持图像输入的模型。需要刷新 TOC、PAGE、REF、SEQ、TA/TOA 等安全域时，可在 Windows + Microsoft Word 环境使用 `--refresh-fields`。

更多命令、退出码和产物说明见 [references/cli-and-output.md](references/cli-and-output.md)。

## 工作原理

排版主链路只接收模型产出的两个结构化结果，模型不直接修改 DOCX；可选视觉复核只返回问题清单，实际修复仍由代码执行。

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

- 模型输出受 JSON Schema 约束，降低模型直接破坏文档的风险，但不代表结构识别永远正确。
- 主要段落格式写入 Word 命名样式；表格、域、标签和部分结构操作仍会使用直接 OOXML 修改。
- 正式产物先写候选文件，预检和文字完整性校验通过后再原子替换。
- 内容原则是“保持语义内容不变”。手工编号转换、摘要/关键词标签结构化或显式创建目录/TOA 标题等变化必须进入审计白名单和报告；其他文字差异会拒绝主稿提交。

## 输出产物

| 产物 | 说明 |
|---|---|
| `result.docx` | 命名样式写入的主稿 |
| `result_tracked.docx` | 仅记录段落/字符属性 `pPrChange/rPrChange` 的修订稿 |
| `result_report.docx` / `.md` | 页面规则、结构操作和段落修改报告 |
| `result_formatspec.json` / `_rolemap.json` / `_stylemap.json` | 结构化中间产物 |
| `result_preflight.json` | Story、域、分节、对象和安全风险 |
| `result_issues.json` | 仅在视觉复核发现问题时生成 |

修订稿不能逐项接受或拒绝页面、分节、页眉页脚、表格、目录、封面和结构移动；这些变化必须结合对照报告审阅。

## 文体覆盖

“支持”表示已有确定性执行路径；“条件支持”表示需要正式模板、显式配置、外部软件或人工复核；“暂不支持”表示只能预检/保留，不能声称完成了该类版式迁移。

| 文体 | 状态 | 当前能力与边界 |
|---|---|---|
| 中文讲话稿、理论学习材料、汇报材料、普通报告 | 支持 | 标题、正文、编号、页面、页眉页脚和常规表格；复杂浮动图文仅保留 |
| 通知、请示、报告、函、会议纪要 | 条件支持 | 可用机关模板或 `official-cn-gbt9704` 基线；不生成完整红头、红线、签发人、印章、密级、紧急程度和版记，不能单凭 Style Pack 宣称完整符合 GB/T 9704 |
| 中文普通论文、毕业论文 | 条件支持 | 支持摘要/关键词、标题层级、封面字段、目录、分节和动态页眉页码；复杂封面、声明、英文扉页、书脊、章节式题注和复杂公式编号需人工复核，详见 [论文模式](references/thesis-mode.md) |
| 英文 memo、essay、普通报告和学术论文 | 支持/条件支持 | 支持英文标题角色、Letter/A4、running head、block quote 和参考文献悬挂缩进；不改写或核验引文、DOI、书目信息和语言质量 |
| APA 7、MLA 9、Chicago 18、Turabian 9 | 条件支持 | 提供可审计的版面基线，不是完整合规检查器；不生成缺失元数据，不核对引文与参考文献 |
| IEEE 期刊稿 | 条件支持（高风险） | 提供保守双栏基线；不完整处理复杂跨栏对象、作者单位块和参考文献重编号，优先使用目标期刊模板 |
| 基础技术手册、产品说明 | 条件支持 | 处理代码/命令、步骤和 WARNING/CAUTION/NOTE/TIP；不移动浮动图、不重建 DTP 版面、不识别截图内容或自动补题注 |
| 美国法律 brief、TA/TOA | 条件支持 | 保留现有 TA/TOA，可按用户逐字提供的引证标记插入域；不发现引证、不验证 Bluebook/local rules、不生成行号 |
| 合同、法律意见书 | 部分支持 | 可处理普通条款层级；签署页、复杂附件、内容控件、填写逻辑和各法域合规仍需专用模板与人工审查 |
| 招标、可研、年度/财务/审计报告、图书和长篇手册 | 部分支持 | 可保守迁移段落、普通表格和既有分节；大量宽表、跨栏对象、章首页、索引和复杂附件不能自动完整制作 |
| 阿拉伯语、希伯来语等 RTL/复杂文字 | 部分支持，未专项验收 | 可从模板迁移 `font_cs`、语言和 RTL/Bidi 属性；没有专用文体 Profile 或 Style Pack |
| 简历/CV、宣传册、Newsletter、海报 | 暂不支持 | 依赖文本框、图标、分栏和绝对定位；只能预检并尽量保留 |
| 复杂表单、申请表、可填写合同模板 | 暂不支持 | 不创建或修改内容控件、复选框、域联动、保护和填写逻辑 |

内置 Style Pack 的具体基线见 [references/style-packs.md](references/style-packs.md)。它们不替代正式模板，也不校正内容。

## Word 功能边界

| 功能 | 状态与限制 |
|---|---|
| 段落、分页、编号 | 支持字体字号、行距缩进、分页约束、手工编号和 Word 自动编号；未识别角色回退到正文 |
| 页面和多节结构 | 可迁移纸张、方向、边距、栏数、页码和首页/奇偶页眉页脚；重建已有复杂分节或创建横向表格节必须显式使用 `--allow-risky-structure` |
| 表格 | 支持常规宽度、列宽、布局、单元格边距、重复表头、跨页断行和垂直对齐；复杂嵌套表、浮动表、计算公式及大型财务表需人工核对 |
| 脚注/尾注 | 扫描并保留；只有规范明确提供 `notes` 规则时调整格式，不增删或改写注释 |
| 题注/图表目录 | 仅对已识别且含简单整数编号的题注创建 SEQ；章节式编号、跨 Run 题注、复杂公式编号和复杂交叉引用不自动转换 |
| 引文/参考文献/索引/TA/TOA | 保留既有域并只刷新安全白名单；不校正引文内容或来源元数据；TOA 仅依据精确配置创建 |
| 文本框、形状、批注、内容控件、修订 Story | 纳入预检和文字完整性校验，但不参与正文角色重排，不重新定位或重建布局 |
| 公式、嵌入对象、浮动图片、SVG | 仅盘点和保留，不编辑对象内容或锚点；无题注图片只报告风险 |
| 直接格式清理 | `controlled` 最保守，`preserve_emphasis` 保留语义强调，`strict` 用于强模板迁移；`strict` 可能移除有意义的字符样式、外文变量或强调 |
| 未接受修订、编辑保护、数字签名、宏、`altChunk` | 默认阻断，避免生成不可安全审计的终稿 |
| 视觉自检 | 可选且最多一轮；渲染器或视觉模型失败时会降级，但不能声称视觉检查通过 |

## 输入格式与环境

正式输出始终是 `.docx`。旧格式会先转换为临时 DOCX，再重新执行全部预检和完整性校验；转换可能改变分页、字体、域和浮动对象。

| 输入 | 状态 | 要求与验收边界 |
|---|---|---|
| `.docx` | 原生支持，已验收 | 标准 OOXML；仍可能被保护、签名、宏或未接受修订阻断 |
| `.doc` | 条件支持，已验收 | Windows + pywin32，并安装 Word 或 WPS；已用真实 OLE/CFBF 二进制 DOC 验收 |
| `.rtf` | 条件支持，已验收 | 优先 LibreOffice，也可由兼容 Word/WPS 回退；属于有损导入 |
| `.wps` | 条件支持，部分验收 | 已验证 OOXML-backed `.wps` 路由；旧版专有二进制 WPS 尚无真实样本，不能宣称覆盖 |
| `.odt` | 条件支持，当前验收环境未通过 | 推荐安装 LibreOffice；没有可用导入筛选器时会明确失败 |
| `.pdf`、扫描件 | 不支持 | 拒绝 OCR 和 PDF 版面还原 |
| `.docm` | 不支持 | 不执行或迁移宏；应在可信环境中另存为无宏 DOCX |
| 加密、需密码或损坏文档 | 不支持 | 不绕过密码或修复损坏包 |

运行环境：

- Python 3.10+；建议使用独立虚拟环境安装 `requirements.txt`。
- `.doc/.wps` 导入和 Windows 渲染需要 pywin32 及 Word/WPS；`.odt/.rtf` 推荐 LibreOffice。
- Word/WPS 候选在独立进程中运行，默认 45 秒超时后回退，并且只清理本次新启动的 Office 进程。可用 `FORMAT_AGENT_COM_TIMEOUT_SECONDS` 在 5–300 秒内调整。
- `--refresh-fields` 仍只支持 Windows + Microsoft Word；Word 不可用时安全降级为下次打开更新域。
- 自然语言规则和语义角色识别需要可用模型；`--spec-json` + `--rolemap-json` 可完全跳过模型。

旧格式、技术手册和法律模式的详细说明见 [references/advanced-modes.md](references/advanced-modes.md)。

## 真实环境验收

2026-09-02 在 Windows + Microsoft Word + WPS 环境进行了端到端验收：

- 原生 DOCX、真实二进制 DOC、真实 RTF 和 OOXML-backed `.wps` 路由通过。
- 5 份可渲染文档共 10 页逐页检查，无缺字、裁切、重叠、表格错位或分页漂移。
- 回归测试：`42 passed`。
- ODT 因该环境没有 LibreOffice且 Word/WPS 导入不可用而未通过。
- 旧版专有二进制 `.wps` 仍缺少真实样本。

因此当前结论是“有条件通过”，不是对所有 Word 文体和所有 Office 环境的无条件兼容声明。

## 示例

| 场景 | 改前 | 改后 |
|---|---|---|
| 党委会议题议案 | ![前](docs/images/case4-党委议案-改前.png) | ![后](docs/images/case4-党委议案-改后.png) |
| 课程论文 | ![前](docs/images/case3-论文-改前.png) | ![后](docs/images/case3-论文-改后.png) |
| 英文 board resolution | ![before](docs/images/en-board-before.png) | ![after](docs/images/en-board-after.png) |

## 相关项目

- 黑客松完整版（GUI、案例和路演页）：[format-agent](https://github.com/KaguraNanaga/format-agent)（已归档）
