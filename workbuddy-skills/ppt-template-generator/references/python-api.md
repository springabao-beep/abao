# Career Compass Python API Reference

## Imports

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
```

## Color Constants

```python
CC_PRIMARY   = RGBColor(0x9D, 0x3D, 0x1D)  # #9D3D1D — titles, emphasis, dark bg
CC_SECONDARY = RGBColor(0xB9, 0x5B, 0x42)  # #B95B42 — subtitles
CC_ACCENT    = RGBColor(0xD6, 0x75, 0x5A)  # #D6755A — body emphasis
CC_LIGHT     = RGBColor(0xDD, 0xBC, 0xB0)  # #DDBCB0 — card backgrounds
CC_BASE      = RGBColor(0xF1, 0xE2, 0xDD)  # #F1E2DD — large-area light fills
CC_PAGE      = RGBColor(0xF7, 0xF3, 0xED)  # #F7F3ED — page base
CC_WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
CC_BLACK     = RGBColor(0x00, 0x00, 0x00)
```

## Common Patterns

### Set slide background
```python
slide.background.fill.solid()
slide.background.fill.fore_color.rgb = CC_PAGE
```

### Create text box
```python
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN

txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(0.7))
tf = txBox.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.LEFT
run = p.add_run()
run.text = "Hello"
run.font.size = Pt(32)
run.font.bold = True
run.font.color.rgb = CC_PRIMARY
run.font.name = "Futura Bk BT"
```

### Add filled rectangle
```python
shape = slide.shapes.add_shape(1, Inches(0.5), Inches(1.2), Inches(3.8), Inches(5.0))
shape.fill.solid()
shape.fill.fore_color.rgb = CC_BASE
shape.line.fill.background()
```
