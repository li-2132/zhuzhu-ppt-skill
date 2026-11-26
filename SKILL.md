---
name: zhuzhu-ppt-skill
description: >-
  面向计算机/科研组会与技术调研汇报的极简白底 PPT 制作技能。强调“讲懂且有深度”而非照搬提纲：主动拉取引用论文原 PDF 并以 300 DPI 截取关键原图、调用生图模型制作纯白底无文字概念示意图、用 python-pptx 构建原生可编辑学术表格与克制白底排版，并严格执行去 AI 味写作规范与逐页视觉自检。当用户需要制作组会汇报、文献调研、论文精读、开题/阶段进展 PPT，且偏好简洁白底、少字去 AI 味、带论文原图或高质量示意图时使用。
---

# zhuzhu-ppt-skill

用于把调研文档、论文笔记或研究进展做成**极简白底、有技术深度、去 AI 味、像真实组会现场**的原生 `.pptx` 演示文稿。

## 核心原则

1. **极简纯白底，拒绝花哨商务模板**：纯白 `#FFFFFF` 画布、16:9 宽屏、大留白、极细灰线；禁止使用带商务图标集群、小旗子路线图、水印装饰或 `10.5pt` 密字卡片的现成商务模板。
2. **讲懂优先，主动深挖一层**：不要机械照搬用户 md 的章节顺序；按“问题出发点 → 核心前提/架构澄清 → 全景定位 → 关键论文机制深挖 → 竞品对比 → 行动计划与讨论”重组叙事。原文引用的关键论文有公开 PDF（如 arXiv）时，**主动下载原文深读机制与局限**，补足笔记里没展开的硬核细节。
3. **以图/表代长文（多模态辅助）**：
   - **论文原图**：从核心论文 PDF 以 300 DPI 截取拓扑图、实测曲线、系统架构或执行轨迹图，附简短出处，并可用原生虚线框 + 批注框直接指出论文局限。
   - **生图模型示意图**：封面主视觉或抽象冲突场景可调用 `image-generator` 生成纯白底极简示意图——**必须要求零文字**，再由 `python-pptx` 叠加清晰原生中文标签。
   - **原生学术表格**：全景矩阵与竞品对比用原生三线表 + `● / ◐ / ○` 符号，保持完全可编辑。
   - **精确文字图（如时间轴）**：用 HTML/SVG + Playwright 高清渲染，避免生图文字乱码。
4. **严格去 AI 味、少字留白**：页面只留短句、关键数字和结论，详细背景与口播词全部放进 PowerPoint 演讲者备注（Speaker Notes）。完整文案红线见 [references/de-ai-style-guide.md](./references/de-ai-style-guide.md)。
5. **页数克制**：组会汇报通常控制在 **8–12 页**，不硬凑独立分章扉页或凑满模板页数。

---

## 标准工作流

### 1. 读懂素材并重构 8–12 页叙事

先完整阅读用户素材（如 `.md` 调研文档），梳理出：
- 听众最关心的核心问题与一个能立刻立住动机的硬核事实/数字；
- 是否存在需要先澄清的前提或路线转向；
- 最接近的 1–2 篇竞品和 2–3 篇可借鉴机制的论文（列出可拉取的 arXiv/PDF 链接）；
- 下一步实验与组会讨论议题。

### 2. 深挖一手论文 & 准备视觉素材

- **下载与索引 PDF**：用 `curl -sL` 拉取核心论文 PDF 到临时/scratch 目录，用 `PyMuPDF (fitz)` 扫描摘要、核心机制段与全部 Figure 标题。
- **300 DPI 精准裁图**：运行 [scripts/crop_pdf_figure.py](./scripts/crop_pdf_figure.py) 导出页图定位坐标后，按比例裁剪并自动去除多余白边：
  ```bash
  # 先导出整页预览确认坐标
  python3 scripts/crop_pdf_figure.py preview paper.pdf --pages 1,3,5 --out-dir ./page_renders
  # 按归一化坐标 (x0,y0,x1,y1) 裁剪 300 DPI 高清图
  python3 scripts/crop_pdf_figure.py crop paper.pdf --page 3 --rect 0.51,0.08,0.87,0.255 --out fig.png
  ```
- **生图模型示意图（按需 1–2 张）**：提示词规范与配色约束见 [references/design-system.md](./references/design-system.md)（务必指定 `#FFFFFF` 纯白背景与 `ABSOLUTELY NO TEXT`）。
- **逐图肉眼核对**：用读图工具检查每张裁剪图和生成图，确认无裁切残留、无乱码文字、边缘干净。

### 3. 用 `python-pptx` 构建原生白底 PPTX

- 排版网格、配色、字号梯度、三线表样式规范见 [references/design-system.md](./references/design-system.md)。
- 直接复用 [scripts/pptx_helpers.py](./scripts/pptx_helpers.py) 中的 `add_header_footer`、`add_picture_fit`、`style_cell`、`add_textbox`、`add_callout`、`add_notes` 等基础组件编写构建脚本。
- **所有文字、表格、批注框均为原生可编辑对象**，每页务必填充对应演讲者备注（Speaker Notes）。

### 4. 去 AI 味文案自审

写入前对照 [references/de-ai-style-guide.md](./references/de-ai-style-guide.md) 逐页过一遍：
- 删掉所有感叹号 `！`、公众号夸饰词（“杀手 / 死穴 / 悬崖 / 暗藏 / 怎么打 / 积木 / 定锚”等）和反问句小标题；
- 砍掉空洞修饰语和重复总结，确保每页简洁、平实、直奔事实。

### 5. 逐页渲染与视觉/结构验证（必做）

运行 [scripts/verify_pptx.py](./scripts/verify_pptx.py) 校验 OOXML ZIP 完整性（零重复项、零断链）、检测文本框纵向溢出，并把每一页渲染为 PNG：

```bash
python3 scripts/verify_pptx.py output.pptx ./slide_previews
```

然后用读图工具**逐页查看渲染图**，重点检查并修正：
- 文本框溢出警告必须为 `0`；
- 消除 1–2 字的孤字换行（widow）；
- 核对叠在生图上的原生中文标签是否居中对齐；
- 核对论文图上的虚线强调框是否精准落在目标区域。

### 6. 交付与防缓存提醒

- 迭代修改用户已打开过的 PPT 时，优先另存为带清晰版本后缀的新文件名（避免 Windows PowerPoint / WPS 未刷新旧缓存导致用户以为没保存）。
- 临时下载的 PDF、裁剪图和渲染预览统一放临时/scratch 目录，保持用户工作目录干净，只留下最终 `.pptx`。

---

## 目录索引

- [references/design-system.md](./references/design-system.md) — 纯白学术版式网格、色板、字号梯度与生图提示词模板
- [references/de-ai-style-guide.md](./references/de-ai-style-guide.md) — 去 AI 味写作规则、禁用词清单、改写对照与备注分工
- [scripts/crop_pdf_figure.py](./scripts/crop_pdf_figure.py) — arXiv/PDF 论文整页预览与 300 DPI 高清裁图工具
- [scripts/pptx_helpers.py](./scripts/pptx_helpers.py) — 纯白学术风 `python-pptx` 可复用绘制库
- [scripts/verify_pptx.py](./scripts/verify_pptx.py) — `.pptx` 忠实渲染 PNG + 文本溢出检测 + OOXML 完整性校验工具
