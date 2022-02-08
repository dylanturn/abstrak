from pydoc import doc
import docx
from docx.enum.section import WD_SECTION
from resume_style import margins
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import yaml
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Inches
from docx.shared import RGBColor
from docx.shared import Pt
from docx.enum.dml import MSO_THEME_COLOR_INDEX

def load_resume_data(data_path):
  with open(data_path, 'r') as stream:
      return yaml.safe_load(stream)

#def get_font_style_attribute(style_name, attribute_name):
#  style = document_fonts.get(style_name, document_fonts.get("default"))
#  if style:
#    return style.get(attribute_name, document_fonts.get("default").get(attribute_name))
#  raise Exception(f"Failed to get attribute {attribute_name} for either the default or {style_name} styles")


# def format_font(style, font):
#   font.name = get_font_style_attribute(style, "name")
#   font.size = get_font_style_attribute(style, "size")
#   font.color.rgb = get_font_style_attribute(style, "color")
#   return font


def insert_standard_section(document):
  section = document.add_section(WD_SECTION.CONTINUOUS)
  section.start_type = WD_SECTION.CONTINUOUS
  cols = section._sectPr.xpath('./w:cols')[0]
  cols.set(qn('w:num'),str(1))
  return section


def insert_columns_section(document, column_count):
  section = document.add_section(WD_SECTION.NEW_COLUMN)
  section.start_type = WD_SECTION.CONTINUOUS
  cols = section._sectPr.xpath('./w:cols')[0]
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
    bottom.set(qn('w:color'), '808080')
    pBdr.append(bottom)


def create_element(name):
    return OxmlElement(name)

def create_attribute(element, name, value):
    element.set(qn(name), value)

def add_page_number(run, type="PAGE", style=None):
    fldChar1 = create_element('w:fldChar')
    create_attribute(fldChar1, 'w:fldCharType', 'begin')

    instrText = create_element('w:instrText')
    create_attribute(instrText, 'xml:space', 'preserve')
    instrText.text = type

    fldChar2 = create_element('w:fldChar')
    create_attribute(fldChar2, 'w:fldCharType', 'end')

    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)

    run.style = style



def add_hyperlink(paragraph, text, url):
    # This gets access to the document.xml.rels file and gets a new relation id value
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    # Create the w:hyperlink tag and add needed values
    hyperlink = docx.oxml.shared.OxmlElement('w:hyperlink')
    hyperlink.set(docx.oxml.shared.qn('r:id'), r_id, )

    # Create a w:r element and a new w:rPr element
    new_run = docx.oxml.shared.OxmlElement('w:r')
    rPr = docx.oxml.shared.OxmlElement('w:rPr')

    # Join all the xml elements together add add the required text to the w:r element
    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)

    # Create a new Run object and add the hyperlink into it
    r = paragraph.add_run ()
    r._r.append (hyperlink)

    # A workaround for the lack of a hyperlink style (doesn't go purple after using the link)
    # Delete this if using a template that has the hyperlink style in it
    r.font.color.theme_color = MSO_THEME_COLOR_INDEX.HYPERLINK
    r.font.underline = True

    return hyperlink