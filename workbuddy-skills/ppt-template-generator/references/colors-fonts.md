# Career Compass Color System & Typography

## Color Palette

| Code | HEX | RGB | Usage |
|------|-----|-----|-------|
| Color 01 | `#9D3D1D` | (157, 61, 29) | Primary — titles, emphasis, dark backgrounds |
| Color 02 | `#B95B42` | (185, 91, 66) | Secondary — subtitles, medium emphasis |
| Color 03 | `#D6755A` | (214, 117, 90) | Accent — body emphasis |
| Color 04 | `#DDBCB0` | (221, 188, 176) | Light — card backgrounds, auxiliary areas |
| Color 05 | `#F1E2DD` | (241, 226, 221) | Base — large-area light fills |
| Color 06 | `#F7F3ED` | (247, 243, 237) | Page base — lightest background |
| White | `#FFFFFF` | (255, 255, 255) | Text on dark backgrounds |
| Black | `#000000` | (0, 0, 0) | Text on light backgrounds |

## Python Usage

```python
from pptx.dml.color import RGBColor

# Define colors
CC_PRIMARY   = RGBColor(0x9D, 0x3D, 0x1D)  # #9D3D1D - warm dark red-brown
CC_SECONDARY = RGBColor(0xB9, 0x5B, 0x42)  # #B95B42
CC_ACCENT    = RGBColor(0xD6, 0x75, 0x5A)  # #D6755A
CC_LIGHT     = RGBColor(0xDD, 0xBC, 0xB0)  # #DDBCB0
CC_BASE      = RGBColor(0xF1, 0xE2, 0xDD)  # #F1E2DD
CC_PAGE      = RGBColor(0xF7, 0xF3, 0xED)  # #F7F3ED
CC_WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
CC_BLACK     = RGBColor(0x00, 0x00, 0x00)
```

## Typography

### Font Family
- **Primary Font**: Futura Bk BT (installed on Dior workstations)
- **Fallback**: Arial, Helvetica, sans-serif
- **Chinese Fallback**: Microsoft YaHei, SimHei

### Font Sizes
| Element | Size | Weight | Example |
|---------|------|--------|---------|
| Cover Title | 44–60pt | Bold / SemiBold | CAREER COMPASS |
| Section Title | 36–44pt | Bold | 2026 Q3 REVIEW |
| Subtitle | 24–32pt | SemiBold | Overview & Strategy |
| Body Heading | 24pt | SemiBold | Key Metrics |
| Body Text | 18–20pt | Regular | Description text |
| Caption | 14–16pt | Light | Source / footnote |
| Slide Number | 12pt | Regular | 01, 02... |

### Paragraph Rules
- Line spacing: 1.2–1.5
- Heading–body gap: 1.5x heading size
- Bullet indent: 0.5"
- Max body width: 9"

## Slide Dimensions
- **Standard**: 13.33" × 7.50" (16:9 widescreen)
- **Margin**: 0.5" all sides
- **Content area**: 12.33" × 6.50"

## Format Checklist (Before Export)

- [ ] All titles use Color 01 (#9D3D1D) on light background
- [ ] Dark cover slides use white text on #9D3D1D background
- [ ] Fonts: Futura Bk BT, fallback to Arial/Helvetica
- [ ] No mixed font families within same text block
- [ ] Font sizes match the hierarchy above
- [ ] Colors use the exact HEX values (no approximations)
- [ ] Slide numbers: bottom-right, 12pt, white on dark / black on light
- [ ] No text smaller than 12pt
- [ ] Images: high-res (min 300 DPI for print)
