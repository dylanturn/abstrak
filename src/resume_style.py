from docx.shared import Pt
from docx.shared import RGBColor
from docx.shared import Inches

margins = {
  "top": Inches(0.5),
  "bottom": Inches(0.5),
  "left": Inches(0.5),
  "right": Inches(0.5)
}

document_fonts = {

  "default": {
    "name": "Calibri",
    "size": Pt(9),
    "color": RGBColor(0x76, 0x71, 0x71) # R-118  G-113  B-113
  },

  "header_title_1": {
    "size": Pt(36)
  },
  "header_title_2": {
    "size": Pt(36),
    # R-46  G-116  B-181
    "color": RGBColor(0x2E, 0x74, 0xB5)
  },
  "header_subtitle": {
    "size": Pt(10),
    # R-46  G-116  B-181
    #"color": RGBColor(0x2E, 0x74, 0xB5)
  },
  "footer": {
    "size": Pt(10)
  }
}