---
name: format-agent
description: 按规范文字、内置 Style Pack 或 Word 模板重排以段落和常规表格为主的中文/英文 Word 文稿。适用于公文、论文、普通报告、基础技术手册和法律 brief；可有损导入 .doc/.wps/.odt/.rtf。不要用于简历、表单、宣传册、Newsletter、海报、PDF 版面还原或内容控件应用。
---

# 格式排版 Agent

把格式混乱的 Word 文稿重排成可继续编辑、可审计的 DOCX。模型只生成 `FormatSpec` 和 `RoleMap`；确定性代码负责修改 OOXML。

## 工作流

1. 对陌生目标稿和参考模板先运行 `--preflight-only`。未接受修订、保护、数字签名、宏和 `altChunk` 是默认阻断项。
2. 在规范文字、Word 模板、`--spec-json` 和内置 Style Pack 中选择一种格式来源。正式机关、学校、法院或期刊规则优先于通用 Style Pack。
3. 运行排版；非零退出码表示未完成。核对 `_preflight.json`、对照报告和完整性结果后再交付。
4. 只有用户要求视觉复核时才加 `--verify`；未启用时，不要因模型缺少图像能力而暂停。

常用命令、退出码、产物和排错见 [references/cli-and-output.md](references/cli-and-output.md)。

## 核心能力

- 页面、分节覆盖、首页/奇偶页眉页脚、命名段落样式、编号、分页约束和普通表格几何。
- 按真实 OOXML 顺序抽取正文及嵌套表格；扫描全部 Story、域、书签、嵌入对象和分节。
- 正文与受保护 Story 的可见文字必须通过完整性校验；模板未规定的角色回退到正文样式。
- 多节目标稿默认保留既有页面特例；结构重建、分栏或横向表格节需要显式 `--allow-risky-structure`。
- 脚注/尾注仅在 `notes` 有规则时格式化；简单题注可显式转换为 SEQ 域并插入图表目录域。
- 域刷新采用本地域白名单，跳过 `INCLUDETEXT/LINK/DDE/RD` 等外部域；外部关系只盘点和保留。

## 文体路由

- 中文公文、技术手册、美国法律 brief/TA/TOA 或旧格式输入：阅读 [references/advanced-modes.md](references/advanced-modes.md)。
- APA、MLA、IEEE、Chicago、Turabian 及全部内置包：阅读 [references/style-packs.md](references/style-packs.md)。
- 中文毕业论文封面、前置页、分节、目录和动态页眉：阅读 [references/thesis-mode.md](references/thesis-mode.md)。

Profile 包括 `official_cn`、`thesis`、`english_general`、`english_academic`、`english_legal`、`english_technical` 和 `english_legal_brief`。英文 academic/legal/technical/brief 默认使用 `preserve_emphasis`。

## 安全边界

- 原生处理 `.docx`；旧格式转换有损且正式输出只允许 `.docx`。拒绝 PDF、`.docm` 和伪装扩展名。
- 简历、宣传册、Newsletter、海报、复杂表单和可填写合同只能预检和尽量保留，不能声称完成自动版式迁移。
- 批注、文本框、内容控件和修订 Story 不参与正文角色重排；公式与嵌入对象只保留。
- 引文、参考文献、索引和法律引证内容不由排版器改写；TA/TOA 只根据用户的精确配置创建。
- 未接受修订、`altChunk`、编辑保护、数字签名或宏默认阻断。
- 输出不得覆盖源稿；完整性失败不替换主终稿。
- 修订稿只记录 `pPrChange/rPrChange`；页面、分节、页眉页脚、表格和结构调整必须结合报告审阅。

## 环境与模型

- Python 3.10+；建议在独立虚拟环境中运行 `pip install -r requirements.txt`，避免宿主环境的 requests/chardet 等依赖冲突。
- `.doc/.wps` 导入和 Windows 渲染需要 pywin32 及 Word/WPS；`.odt/.rtf` 推荐 LibreOffice。Word/WPS 候选在独立限时进程中运行，默认 45 秒后回退；`--refresh-fields` 仍只支持 Microsoft Word。
- 需要模型时配置 `LLM_BASE_URL`、`LLM_API_KEY`、`LLM_MODEL`；视觉模式可另设 `LLM_VISION_MODEL`。
- `--spec-json` 与 `--rolemap-json` 可完全跳过模型。
