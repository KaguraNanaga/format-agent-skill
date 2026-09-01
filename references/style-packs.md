# 内置 Style Pack

Style Pack 只负责可确定执行的版面和段落格式。它们保留现有域及可见文字，不改作者、年份、题名、DOI、引文顺序或参考文献标点。

学校、课程、会议和期刊的明确要求优先于这里的通用基线。能取得官方 `.docx` 模板时，应优先使用 `--template`，而不是 Style Pack。

## 可用包

| 名称 | 当前基线 | 不负责的事项 |
|---|---|---|
| `apa7-student` | Letter、四边 1 英寸、Times New Roman 12 pt、正文双倍行距、0.5 英寸首行缩进、右上页码、参考文献悬挂缩进 | 不生成学生封面元数据，不验证文内引文与参考文献对应关系，不改写 APA 引文 |
| `mla9` | Letter、四边 1 英寸、Times New Roman 12 pt、全文双倍行距、0.5 英寸首行缩进、Works Cited 悬挂缩进；可用 `--running-head` 提供姓氏 | 不生成姓名/教师/课程/日期块，不核验作品引用，不改写 MLA 引文 |
| `ieee-journal` | 保守的 Letter 双栏期刊基线，以及标题、正文、标题层级、图表题注和参考文献的常见字号 | IEEE 各刊物版式并不完全相同；不替代目标刊物官方模板，不重编号方括号引文，不整理参考文献元数据 |
| `chicago18-notes-bibliography` | Letter、1 英寸边距、正文双倍行距、脚注 10 pt 单倍、块引文与书目条目单倍 | 不改写 Chicago 18 引证，不决定应使用脚注还是尾注 |
| `chicago18-author-date` | 与 Chicago 基线相同，参考文献条目采用 author-date 常用悬挂缩进 | 不核对作者—年份对应关系 |
| `turabian9-student` | Chicago notes-bibliography 的学生论文保守基线 | 不生成院系标题页和声明；学校指南优先 |
| `official-cn-gbt9704` | A4、公文常见版心、字号字体和奇偶页外侧页码 | 不生成完整红头、印章和版记；机关模板优先 |
| `technical-manual` | 代码/命令、四类提示框、步骤和图题邻接分页绑定 | 不移动浮动图或重建 DTP 版面 |
| `us-legal-brief` | Letter、常见 brief 段落角色并保留 TA/TOA；可显式插入域 | 不判断 Bluebook/法院 local rules，不自动发现引证或生成行号 |

APA 7 允许多种易读字体；本包为保持结果确定性选择 Times New Roman 12 pt。IEEE 包刻意称为 `journal` baseline，而不是通用合规转换器。

## 命令

```bash
python main.py --style-pack apa7-student --target paper.docx --out output/paper_apa.docx
python main.py --style-pack mla9 --running-head Rivera --target paper.docx --out output/paper_mla.docx
python main.py --style-pack ieee-journal --target article.docx --out output/article_ieee.docx
python main.py --style-pack chicago18-notes-bibliography --target paper.docx --out output/paper_chicago.docx
python main.py --style-pack turabian9-student --target thesis.docx --out output/thesis_turabian.docx
python main.py --style-pack official-cn-gbt9704 --target notice.docx --out output/notice.docx
python main.py --style-pack technical-manual --target manual.docx --out output/manual.docx
python main.py --style-pack us-legal-brief --target brief.docx --out output/brief.docx
```

先执行 `--preflight-only`。IEEE 包会新增双栏节；源文档已有复杂分节时默认阻断，确认备份和结构风险后才使用 `--allow-risky-structure`。

## 学术域和图表

- `academic.preserve_fields=true`：保留已有域并设置 Word 打开时更新域。
- `academic.caption_numbering=true`：只将已经标为 `figure_caption` / `table_caption` 且含简单整数编号的题注转换为 `SEQ Figure` / `SEQ Table`；可见文字不变。
- `academic.lists.figures=true` / `tables=true`：只有 RoleMap 已识别 `list_of_figures_heading` / `list_of_tables_heading` 时才在其后插入图表目录域，不凭空创建标题。
- `table.landscape_table_indices=[0]`：把第 1 个顶层表格包入横向节。它是结构性操作，必须预检并显式授权。

插入或保留的域需要 Microsoft Word 刷新才能得到最终页码和编号。Windows 可加 `--refresh-fields`；否则在 Word 中全选后按 F9。

## 规范来源

- APA：[Concise Guide to APA Style, Seventh Edition](https://www.apa.org/pubs/books/concise-guide-apa-style-7th-edition-spiral)；[APA Student Paper Setup Guide](https://apastyle.apa.org/instructional-aids/student-paper-setup-guide.pdf)
- MLA：[Formatting a Research Paper](https://style.mla.org/formatting-papers/)；[Styling Headings and Subheadings](https://style.mla.org/styling-headings-and-subheadings/)
- IEEE：[IEEE Article Templates](https://journals.ieeeauthorcenter.ieee.org/create-your-ieee-journal-article/authoring-tools-and-templates/tools-for-ieee-authors/ieee-article-templates/)；[IEEE Reference Guide](https://journals.ieeeauthorcenter.ieee.org/wp-content/uploads/sites/7/IEEE_Reference_Guide.pdf)
- Chicago：[Chicago Manual of Style](https://www.chicagomanualofstyle.org/home.html)
- Turabian：[Turabian Citation Quick Guide](https://www.chicagomanualofstyle.org/turabian/citation-guide.html)
- 中文公文：[GB/T 9704-2012](https://openstd.samr.gov.cn/bzgk/gb/newGbInfo?hcno=F3CC9BEF482524C895FDA7A08BB4A70E)
