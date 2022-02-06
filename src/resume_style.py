from docx.shared import Pt
from docx.shared import RGBColor
from docx.shared import Inches

margins = {
  "top": Inches(0.5),
  "bottom": Inches(0.5),
  "left": Inches(0.5),
  "right": Inches(0.5)
}

# base_color = RGBColor(0x76, 0x71, 0x71) # R-118  G-113  B-113
base_color = RGBColor(0x3B, 0x38, 0x38) # R-59  G-56  B-56
accent_color = RGBColor(0x2E, 0x74, 0xB5) # R-46  G-116  B-181

document_font = "Segoe UI Symbol"
heading_font = "Calibri Light"

document_fonts = {

  "default": {
    "name": "Consolas",
    "size": Pt(9),
    "color": base_color
  },

  "header_title_1": {
    "size": Pt(36),
    "color": base_color
  },
  "header_title_2": {
    "size": Pt(36),
    "color": accent_color
  },
  "header_subtitle": {
    "size": Pt(10),
    "color": base_color
  },



  "section_title": {
    "name": heading_font,
    "size": Pt(11),
    "color": accent_color
  },
  "sub_section_title": {
    "name": heading_font,
    "size": Pt(10),
    "color": accent_color
  },
  "section_content": {
    "name": document_font,
    "size": Pt(9),
    "color": base_color
  },
  "list_bullet": {
    "name": document_font,
    "size": Pt(9),
    "color": accent_color
  },
  
  "list": {
    "name": document_font,
    "size": Pt(9),
    "color": base_color
  },

  "footer": {
    "size": Pt(10),
    "color": base_color
  }
}

