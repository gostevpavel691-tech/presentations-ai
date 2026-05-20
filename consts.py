# ==================== ШАБЛОНЫ ДЛЯ ПРЕМИУМ ====================
TEMPLATE1 = """
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import os

prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)

# Титульный слайд
slide = prs.slides.add_slide(prs.slide_layouts[6])
title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2), Inches(9), Inches(1.5))
title_frame = title_box.text_frame
title_frame.text = "НАЗВАНИЕ ПРЕЗЕНТАЦИИ"
title_frame.paragraphs[0].font.size = Pt(44)
title_frame.paragraphs[0].font.bold = True
title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

sub_box = slide.shapes.add_textbox(Inches(0.5), Inches(4), Inches(9), Inches(1))
sub_frame = sub_box.text_frame
sub_frame.text = "Подзаголовок"
sub_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

# Контентный слайд (текст слева, картинка справа)
slide = prs.slides.add_slide(prs.slide_layouts[6])
title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
title_frame = title_box.text_frame
title_frame.text = "ЗАГОЛОВОК СЛАЙДА"
title_frame.paragraphs[0].font.size = Pt(32)
title_frame.paragraphs[0].font.bold = True

text_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(4.5), Inches(5.5))
text_frame = text_box.text_frame
text_frame.word_wrap = True
text_frame.text = "• Первый пункт"
p = text_frame.add_paragraph()
p.text = "• Второй пункт"
p = text_frame.add_paragraph()
p.text = "• Третий пункт"

if os.path.exists("img_0.jpg"):
    slide.shapes.add_picture("img_0.jpg", Inches(5.5), Inches(1.5), width=Inches(4), height=Inches(5))

prs.save("presentation.pptx")
"""

TEMPLATE2 = """
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import os

prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)

# Титульный слайд
slide = prs.slides.add_slide(prs.slide_layouts[6])
title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2), Inches(9), Inches(1.5))
title_frame = title_box.text_frame
title_frame.text = "НАЗВАНИЕ ПРЕЗЕНТАЦИИ"
title_frame.paragraphs[0].font.size = Pt(44)
title_frame.paragraphs[0].font.bold = True
title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

sub_box = slide.shapes.add_textbox(Inches(0.5), Inches(4), Inches(9), Inches(1))
sub_frame = sub_box.text_frame
sub_frame.text = "Подзаголовок"
sub_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

# Контентный слайд (текст сверху, картинка снизу)
slide = prs.slides.add_slide(prs.slide_layouts[6])
title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
title_frame = title_box.text_frame
title_frame.text = "ЗАГОЛОВОК СЛАЙДА"
title_frame.paragraphs[0].font.size = Pt(32)
title_frame.paragraphs[0].font.bold = True

text_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.3), Inches(9), Inches(2.5))
text_frame = text_box.text_frame
text_frame.word_wrap = True
text_frame.text = "• Первый пункт"
p = text_frame.add_paragraph()
p.text = "• Второй пункт"
p = text_frame.add_paragraph()
p.text = "• Третий пункт"

if os.path.exists("img_0.jpg"):
    slide.shapes.add_picture("img_0.jpg", Inches(2.5), Inches(4.2), width=Inches(5), height=Inches(3.5))

prs.save("presentation.pptx")
"""