---
name: ppt-template-generator
description: |
  Generate, format, or audit PowerPoint presentations following Career Compass / Dior LVMH brand template.
  Triggers: "create a PPT", "generate a presentation", "PPT template", "make a slide deck",
  "Career Compass style PPT", "format a PPT", "check PPT formatting", "audit PPT compliance",
  "unify PPT format", "品牌PPT", "生成PPT", "格式统一".
  Provides Career Compass color system, 10 branded layout generators, format compliance checker,
  and universal format unifier.
agent_created: true
---

# PPT Template Generator — Career Compass / Dior LVMH

## Overview

Generate CC-branded PPTs, audit for compliance, or unify existing PPTs to CC palette + Futura Bk BT.
Built from `05. Career Compass Presentation Template.pptx` (47 slides, 16:9, 13.33" × 7.50").

## Color System

| Name | HEX | RGB | Usage |
|------|-----|-----|-------|
| PRIMARY | `#9D3D1D` | (157,61,29) | Background fills, accent bars, decorative elements |
| SECONDARY | `#B95B42` | (185,91,66) | Subtle decorative fills |
| ACCENT | `#D6755A` | (214,117,90) | Decorative elements |
| LIGHT | `#DDBCB0` | (221,188,176) | Card backgrounds |
| BASE | `#F1E2DD` | (241,226,221) | Large-area light fills |
| PAGE | `#F7F3ED` | (247,243,237) | Page base (lightest bg) |

Dark slides (Cover, Closing): bg=`#9D3D1D`, text=white. Light slides: bg=`#F7F3ED`, all text = black (`#000000`).

## Typography

**Font**: Futura Bk BT → Arial → Microsoft YaHei. Sizes: Cover 54pt, Section 32-36pt, Body 14-16pt, Caption 11-12pt. Min 9pt. Margins: 0.5".

## ⚠️ 用户硬性要求（必须遵守）

1. **直角边框** – 所有形状/卡片使用直角 Rectangle，不得使用圆角 Rounded Rectangle。
2. **黑色字体** – 所有文字（标题、正文、标签、副标题等）使用纯黑 `#000000`。CC_PRIMARY 棕色只用于背景色块、装饰分割线、色条等非文字元素。深色背景页（Cover/Closing）的白字保持不变。

## Layout Functions (10)

Import from `scripts/generate_layouts`:
1. `cover_slide(pptx, title, subtitle="")` — Dark bg centered 54pt title
2. `three_col_points(pptx, heading, points)` — 3 cards with circle badge
3. `pillars_showcase(pptx, title, pillars)` — 2×2 grid with accent bar
4. `timeline(pptx, title, steps)` — Horizontal connected nodes
5. `data_showcase(pptx, title, metrics)` — Big number cards
6. `detailed_content(pptx, title, body, info_box="")` — Two-column content
7. `discussion_slide(pptx, title, questions)` — Numbered Q items
8. `team_grid(pptx, title, members)` — Avatar + name cards
9. `quote_slide(pptx, quote, author="")` — Decorative quote mark
10. `closing_slide(pptx, text="THANK YOU", subtitle="")` — Dark bg closing

All accept optional `number=N` for slide page number.

## Workflows

### Generate new PPT
```python
from scripts.generate_layouts import new_presentation, cover_slide, data_showcase, closing_slide
pptx = new_presentation()
cover_slide(pptx, "CAREER COMPASS", "2026 Q3 Review")
# ... add more slides
pptx.save("output.pptx")
```

### Audit existing PPT
`python scripts/format_check.py <file.pptx>`

### Unify format of existing PPT
`python scripts/apply_cc_format.py <file.pptx>`
`python scripts/apply_cc_format.py <file.pptx> --dry-run`  (preview)
`python scripts/apply_cc_format.py <file.pptx> --no-fonts`  (colors only)

## Resource Files

| File | Purpose |
|------|---------|
| `scripts/generate_layouts.py` | 10 layout generators |
| `scripts/format_check.py` | Brand compliance auditor |
| `scripts/apply_cc_format.py` | Universal format unifier |
| `references/colors-fonts.md` | Color system + font specs |
| `references/checklist.md` | Manual format checklist |
| `references/python-api.md` | Python import & pattern reference |
| `references/template-info.md` | Source template metadata |
