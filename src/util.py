from pydoc import doc
from docx.enum.section import WD_SECTION
from resume_style import document_fonts, margins
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import yaml
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Inches


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


def configure_styles(document):

  num_xml = document.part.numbering_part.numbering_definitions._numbering
  num_1 = num_xml.num_having_numId(1)
  abstract_id = num_1.abstractNumId.val
  element = document.part.numbering_part.element.xpath(f"//w:abstractNum[@w:abstractNumId={abstract_id}]/w:lvl/w:rPr")[0]
  color = OxmlElement('w:color')
  color.set(qn('w:val'), '2E74B5')
  element.insert(0,color)


  normal_style = document.styles['Normal']
  format_font("section_content",normal_style.font)
  normal_style.font.bold = False
  rFonts = normal_style.element.rPr.rFonts
  rFonts.set(qn("w:asciiTheme"), "Calibri Light")


  style = document.styles['Heading 1']
  format_font("section_title",style.font)
  style.font.bold = False
  rFonts = style.element.rPr.rFonts
  rFonts.set(qn("w:asciiTheme"), "Calibri Light")


  style = document.styles['Heading 2']
  format_font("sub_section_title",style.font)
  style.font.bold = False
  rFonts = style.element.rPr.rFonts
  rFonts.set(qn("w:asciiTheme"), "Calibri Light")