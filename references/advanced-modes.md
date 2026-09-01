# 公文、技术手册、法律 brief 与输入转换

仅在任务涉及对应文体或旧格式输入时阅读本页。这里的 Profile/Style Pack 是可审计基线，不替代机关、法院、客户或厂商给出的正式 Word 模板。

## 中文公文

`official-cn-gbt9704` 提供 A4、常见版心边距、正文/标题字号字体、公文四级标题、发文字号、主送机关、结束语、落款日期、抄送和奇偶页外侧 `— 页码 —` 基线：

```bash
python main.py --style-pack official-cn-gbt9704 \
  --target notice.docx --out output/notice.docx
```

此实现参考用户的 [document-format-skills](https://github.com/KaguraNanaga/document-format-skills) 所采用的 Word/WPS 兼容思路，但没有复制完整的版记/红头生成工作流。发文机关标志、红色分隔线、签发人版式、印章、版记、密级和紧急程度仍应使用机关专用模板或该专门 Skill。

## 技术手册

`technical-manual` 确定性处理：

- `WARNING`、`CAUTION/IMPORTANT`、`NOTE`、`TIP` 的底纹与左边框；
- `Step n`、shell/PowerShell 命令、HTTP method + path 和常见代码字体；
- 已有图形段与相邻 `figure_caption` 的 `keep_with_next/keep_together`；
- 无题注图形只报告 `UNBOUND_FIGURE`，不移动或重新锚定对象。

它不处理 DTP 式文本框、多栏浮动图、标注线、SVG 编辑、截图内容识别或自动生成缺失题注。

## 美国法律 brief 与 Table of Authorities

`us-legal-brief` 默认只保留并盘点现有 TA/TOA。新增域必须显式提供配置：

CLI 使用独立的引证标记数组：

```json
[
  {
    "text": "Example v. Sample",
    "long": "Example v. Sample, 123 F.3d 456 (2024)",
    "short": "Example",
    "category": 1
  }
]
```

```bash
python main.py --style-pack us-legal-brief \
  --legal-citations-json citation-marks.json --insert-toa \
  --target brief.docx --out output/brief.docx
```

若文稿没有 TOA 标题，只有用户明确要求时才把 `--insert-toa` 改为 `--create-toa-heading`。FormatSpec API 也支持下列 `legal` 对象：

```json
{
  "profile": "english_legal_brief",
  "legal": {
    "preserve_toa": true,
    "insert_toa": true,
    "create_heading": false,
    "citation_marks": [
      {
        "text": "Example v. Sample",
        "long": "Example v. Sample, 123 F.3d 456 (2024)",
        "short": "Example",
        "category": 1
      }
    ]
  },
  "roles": {
    "body": {
      "font_eastasia": "Times New Roman",
      "font_ascii": "Times New Roman",
      "size_pt": 12,
      "alignment": "justify"
    }
  }
}
```

- `text` 必须逐字存在于正文；找不到就记录诊断，不猜测替代引证。
- `insert_toa=true` 默认要求文档已经有被识别的 `table_of_authorities_heading`。
- `create_heading=true` 才允许在文末新增 `TABLE OF AUTHORITIES`；该文字进入完整性白名单与修改日志。
- 再次运行不会重复插入相同段落中的 TA 或已有 TOA 域。
- Word 打开时会请求更新域；Windows 可用 `--refresh-fields` 刷新本地安全域。`INCLUDETEXT/LINK/DDE/RD` 和未知外部域会被跳过。

本模块不验证 Bluebook、引证真实性、法院管辖规则、封面颜色、字数限制或电子提交要求；也不生成行号。具体法院的 local rules 永远优先。

## `.doc/.wps/.odt/.rtf` 输入

这些格式只作为输入，正式输出始终是 `.docx`：

| 输入 | 转换顺序 | 要求 |
|---|---|---|
| `.doc` | Microsoft Word → WPS | Windows + pywin32，并安装至少一个应用 |
| `.wps` | WPS → Microsoft Word | Windows + pywin32；WPS 优先 |
| `.odt` / `.rtf` | LibreOffice → Word/WPS | 推荐安装 LibreOffice；RTF 的 COM 兼容较广，ODT 是否可回退取决于本机 Office 导入筛选器 |
| `.pdf` | 拒绝 | PDF/OCR 版面还原不是安全的 Word 格式迁移 |

转换文件位于临时目录；系统把转换后的 DOCX 当成新输入，重新执行全部 Story 预检和文本完整性校验，结束后清理临时文件。`input_conversion` 会记录源路径、转换器、有损标记和警告。

Word/WPS COM 转换会禁用自动宏和打开时链接更新；仍只应处理用户授权的本地文件。每个 Office 候选在独立工作进程中运行，默认 45 秒超时后只清理本次新启动的 Office 进程并回退下一候选，不结束运行前已有的 Office 会话。可用 `FORMAT_AGENT_COM_TIMEOUT_SECONDS` 在 5~300 秒间调整。LibreOffice 使用独立临时转换目录。

常见损失包括分页变化、缺失字体替代、浮动对象锚点变化、域降级及 WPS/Word 私有功能丢失。即使校验通过，交付前仍应视觉核对关键页面。
