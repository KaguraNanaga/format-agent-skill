# CLI、产物和退出码

需要运行命令、解释失败或核对交付物时阅读本页。所有命令都在 Skill 目录执行。

## 常用命令

```bash
# 先预检；加入 --template 时会分别报告 target/template 风险
python main.py --preflight-only --target source.docx --out output/result.docx

# 自然语言规范或 Word 模板
python main.py --spec rules.txt --target source.docx --out output/result.docx
python main.py --template template.docx --target source.docx --out output/result.docx

# 内置 Style Pack
python main.py --style-pack apa7-student --target paper.docx --out output/paper.docx

# 完全跳过模型
python main.py --spec-json examples/spec_std.json \
  --rolemap-json examples/rolemap_std.json \
  --target examples/messy.docx --out output/result.docx
```

详细 Style Pack 名称见 [style-packs.md](style-packs.md)。论文、技术/法律和旧格式参数分别见 [thesis-mode.md](thesis-mode.md) 与 [advanced-modes.md](advanced-modes.md)。

只有用户要求视觉复核时才加 `--verify`。它最多做一轮定向修复；视觉模型或渲染失败会降级成警告，已通过结构/文字验收的排版结果仍可提交，但不能声称视觉检查通过。

## 输出产物

- `result.docx`：命名样式写入的主稿。
- `result_tracked.docx`：仅包含段落/字符属性 `pPrChange/rPrChange` 的修订稿。
- `result_report.docx` / `.md`：页面规则和段落修改报告。
- `result_formatspec.json` / `_rolemap.json` / `_stylemap.json`：结构化中间产物。
- `result_preflight.json`：Story、域、分节、对象和安全风险。
- `result_issues.json`：仅在视觉复核返回问题时生成。

手工编号转自动编号、明确创建的目录/封面/TOA 标题等可见变化必须进入审计白名单。其他正文或受保护 Story 文字差异会拒绝主稿提交并生成 `_integrity_failure.json`。

## 退出码

| 码 | 含义 |
|---|---|
| `0` | 结构与文字验收完成；仍需阅读预检警告和报告 |
| `2` | 能力预检阻断或 DOCX 结构错误 |
| `3` | 文本完整性失败，主终稿未替换 |
| `4` | 输出路径与源稿/其他产物冲突 |
| `5` | 旧格式输入转换失败 |
| `6` | JSON、RoleMap、FormatSpec 或参数契约错误 |
| `7` | 文件读写错误 |

## 域与渲染

`--refresh-fields` 仅在 Windows + Microsoft Word 下可用。它只刷新 PAGE、TOC、REF、SEQ、TA/TOA 等白名单域；`INCLUDETEXT`、`LINK`、`DDE`、`RD` 和未知域会跳过。刷新在独立限时进程中执行；Word 被占用、不可用或超时时会降级为文档下次打开时更新，不阻断已通过文字校验的终稿。交付前也可在 Word 中全选并按 F9。

`--verify` 在 Windows 按 Microsoft Word → WPS → LibreOffice 回退渲染，在 macOS/Linux 使用 LibreOffice；Office 候选采用与输入转换相同的独立进程和超时边界。PDF 转 PNG 按 PyMuPDF → pypdfium2 → pdf2image 回退。视觉复核还需要支持图像输入的模型。

## 示例文件

- `examples/messy.docx`、`spec.txt`：中文公文示例。
- `examples/spec_std.json`、`rolemap_std.json`：无模型基准输入。
- `examples/模板-党委会议题样表.docx`：Word 模板来源示例。
