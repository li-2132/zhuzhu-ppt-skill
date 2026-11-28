# zhuzhu-ppt-skill

面向计算机 / 科研组会与技术调研汇报的**极简白底学术 PPT 制作 Skill**（适用于 Antigravity / Claude Code / Codex 等支持 `SKILL.md` 的 AI Agent）。

## 核心特点

- **纯白极简学术风**：`16:9` 纯白底、大留白、极细分割线、克制三色板（深蓝 + 青绿 + 局部暖橙），彻底抛弃花哨商务模板和小字密排卡片。
- **讲懂优先，主动深挖一手论文**：不机械照搬提纲目录，而是按组会逻辑重组叙事，并自动拉取引用论文 PDF（如 arXiv），深挖机制、实测数据与局限。
- **多模态辅助**：
  - `scripts/crop_pdf_figure.py`：从论文 PDF 索引图注并以 **300 DPI** 精准截取关键原图（自动去白边）；
  - 生图模型提示词规范：生成**纯白底、零文字**的极简概念示意图，再用 `python-pptx` 叠加清晰原生中文标注；
  - `scripts/pptx_helpers.py`：原生可编辑学术三线表、能力对比矩阵（`● / ◐ / ○`）与演讲者备注。
- **严格去 AI 味（文案 + 版式双约束）**：内置禁用词与改写规范（去除感叹号、公众号夸饰词、反问句小标题），同时坚持**去盒化纯白排版**——禁止滥用“彩色竖线 + 淡色背景框”堆砌卡片，靠留白、字重层级与极细分割线建立层次；口播细节全部写入演讲者备注栏。
- **逐页视觉与结构自检**：`scripts/verify_pptx.py` 忠实渲染每页 PNG，同时自动检测文本框纵向溢出、OOXML ZIP 完整性、残留 AI 腔/省略号及视觉 AI 卡片滥用（`VISUAL_AI`）。

## 目录结构

```text
zhuzhu-ppt-skill/
├── SKILL.md                          # Agent 入口指南与 6 步标准工作流
├── references/
│   ├── design-system.md              # 白底网格、配色、字号梯度与生图规范
│   └── de-ai-style-guide.md          # 去 AI 味规则、禁用词清单与改写示例
└── scripts/
    ├── crop_pdf_figure.py            # PDF 图注索引、整页预览与 300 DPI 裁图
    ├── pptx_helpers.py               # 纯白学术风 python-pptx 封装库
    └── verify_pptx.py                # 逐页 PNG 渲染 + 溢出 / OOXML / 文风自检
```

## 安装方法

### 1. 安装到项目工作区（推荐）

```bash
mkdir -p .agents/skills
git clone https://github.com/li-2132/zhuzhu-ppt-skill.git .agents/skills/zhuzhu-ppt-skill
```

### 2. 或安装为全局 Skill

```bash
mkdir -p ~/.gemini/config/skills
git clone https://github.com/li-2132/zhuzhu-ppt-skill.git ~/.gemini/config/skills/zhuzhu-ppt-skill
```

### 3. Python 依赖

```bash
pip install python-pptx pymupdf pillow playwright
playwright install chromium
```

## 使用示例

在支持 Skill 的 Agent 中直接说：

> 请用 `zhuzhu-ppt-skill` 读取当前目录下的 `xxx调研.md`，帮我做一份 8–10 页的组会 PPT。


## 效果展示
以AGY为例

效果图：
<img width="2542" height="660" alt="image" src="https://github.com/user-attachments/assets/36849268-b4cf-40ad-b142-d4c0b333c117" />

<img width="1920" height="1131" alt="image" src="https://github.com/user-attachments/assets/c8d177c8-aec4-49a9-bc1a-638e737a3a48" />

<img width="1920" height="1118" alt="image" src="https://github.com/user-attachments/assets/7ee29a58-790a-430e-a6c8-bb6a1701680b" />

## 🙏 鸣谢

感谢 [LINUX DO](https://linux.do/) 社区的交流与支持，以及 [python-pptx](https://python-pptx.readthedocs.io/)、[PyMuPDF](https://pymupdf.readthedocs.io/)、[Playwright](https://playwright.dev/) 等开源项目。

[![LINUX DO](https://img.shields.io/badge/LINUX%20DO-Community-1677FF?style=flat-square)](https://linux.do/)

## License

MIT License
