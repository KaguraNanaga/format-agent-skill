# format-agent-skill

通用文档格式排版 Agent Skill。**理解归 AI，动手归代码，中间用 JSON 交接。**

给它一份格式规范（自然语言规范文字，或排好版的 Word 模板）加一份格式混乱的 docx，
Agent 自主完成理解规范、识别文档结构、逐段改写格式，输出可直接交付的 Word 文档。

> 本仓库即为 Skill 本体：`SKILL.md` 在根目录，clone/下载整个仓库即可安装。

## 安装（对你的 Agent 说一句话）

> 帮我安装这个 skill：https://github.com/KaguraNanaga/format-agent-skill

Agent 会读取 `SKILL.md`，自行完成安全审计、依赖安装与冒烟测试。

- 支持 Skill 机制的 Agent 环境均可使用：Kimi Code、腾讯 WorkBuddy、字节 Trae、阿里 Qoder、百度 Comate 等
- 私有仓库克隆卡住时，用已登录的 gh CLI：`gh repo clone KaguraNanaga/format-agent-skill`
- 手动安装：把本仓库整个目录复制到 Agent 的技能目录（如 `~/.workbuddy/skills/format-agent/`）

## 使用

安装后，对 Agent 直接描述任务即可，例如：

> 按《规范文字.txt》的要求，把"待排版.docx"重排一下。

> 参照这份党委会议题模板，把这份汇报材料排成正式格式。

也可以用命令行直接运行（本仓库根目录）：

```bash
pip install -r requirements.txt

# 规范文字作为格式来源
python main.py --spec examples/spec.txt --target examples/messy.docx --out out/排版后.docx

# Word 模板作为格式来源
python main.py --template "examples/模板-党委会议题样表.docx" --target examples/messy.docx --out out/排版后.docx

# 排版后再做一轮视觉自检（需多模态模型）
python main.py --spec examples/spec.txt --target examples/messy.docx --out out/排版后.docx --verify
```

## 模型配置

复制 `.env.example` 为 `.env` 并填写：

```bash
LLM_BASE_URL=https://你的端点/v1
LLM_API_KEY=你的key
LLM_MODEL=你的模型
```

**建议多模态模型**（支持图像输入：GPT-4o、Kimi K3、Qwen-VL、GLM-4V 等）。
排版主流程只需文本能力；视觉自检环节要把渲染图交给模型质检，纯文本模型无法使用。

不配模型也能跑：直接提供 FormatSpec / RoleMap JSON，走确定性降级链路
（`python main.py --spec-json ... --rolemap-json ...`）。

## 输出产物

- `排版后.docx` —— 命名样式写入的干净稿（Word 里可继续编辑）
- `排版后_tracked.docx` —— 修订模式：Word 审阅视图逐条接受/拒绝每处格式改动
- `排版后_report.docx` / `.md` —— 段落级修改对照报告
- `_formatspec.json` / `_rolemap.json` —— Agent 理解规范的中间产物

## 原理一句话

LLM 只产出两个 JSON——格式规则（FormatSpec）和段落角色（RoleMap），
JSON Schema 校验失败会回喂自我修正；所有文档修改由确定性代码完成。
模型负责理解，代码负责动手，排版零幻觉。

## 相关

- 黑客松完整版（GUI 演示界面、四象限案例、路演页）：
  [format-agent-01](https://github.com/KaguraNanaga/format-agent-01)（已归档）
