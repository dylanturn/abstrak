from docx.enum.section import WD_SECTION
from resume_style import document_fonts, margins
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import yaml



def load_resume_data(data_path):
  with open(data_path, 'r') as stream:
      return yaml.safe_load(stream)

def get_font_style_attribute(style_name, attribute_name):
  style = document_fonts.get(style_name, document_fonts.get("default"))
  if style:
    return style.get(attribute_name, document_fonts.get("default").get(attribute_name))
  raise Exception(f"Failed to get attribute {attribute_name} for either the default or {style_name} styles")


def format_font(style, font):
  font.name = get_font_style_attribute(style, "name")
  font.size = get_font_style_attribute(style, "size")
  font.color.rgb = get_font_style_attribute(style, "color")
  return font


def insert_standard_section(document, start_type):
  section = document.add_section(WD_SECTION.CONTINUOUS)
  section.start_type = start_type
  section.top_margin = margins.get("top")
  section.right_margin = margins.get("right")
  section.bottom_margin = margins.get("bottom")
  section.left_margin = margins.get("left")
  return section


def insert_columns_section(document, start_type, column_count):
  section = document.add_section(WD_SECTION.NEW_COLUMN)
  section.start_type = start_type
  section.top_margin = margins.get("top")
  section.right_margin = margins.get("right")
  section.bottom_margin = margins.get("bottom")
  section.left_margin = margins.get("left")
  sectPr = section._sectPr
  cols = sectPr.xpath('./w:cols')[0]
  cols.set(qn('w:num'),str(column_count))
  return section


def insert_horizontal_rule(paragraph):
    p = paragraph._p  # p is the <w:p> XML element
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    pPr.insert_element_before(pBdr,
        'w:shd', 'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap',
        'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN',
        'w:bidi', 'w:adjustRightInd', 'w:snapToGrid', 'w:spacing', 'w:ind',
        'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc',
        'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap',
        'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr',
        'w:pPrChange'
    )
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'dotted')
    bottom.set(qn('w:sz'), '24')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'auto')
    pBdr.append(bottom)