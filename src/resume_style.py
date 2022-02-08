from docx.shared import Pt
from docx.shared import RGBColor
from docx.shared import Inches
import yaml
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement

margins = {
  "top": Inches(0.5),
  "bottom": Inches(0.5),
  "left": Inches(0.5),
  "right": Inches(0.5)
}

class AbstrakStyle:

  def _get_attribute(self, style, attribute_name):
    attribute = style.get(attribute_name)
    if attribute is not None:
      return attribute
    
    base_style = style.get("base_style")
    if base_style is not None:
      attribute = self._get_attribute(self.loaded_styles["styles"].get(base_style), attribute_name) 
      if attribute is not None:
        return attribute
    
    raise Exception(f"Failed to get attribute {attribute_name} for style {style['style_name']}")


  def __init__(self, document, style_path="resume-style.yml"):
    with open(style_path, 'r') as stream:
      self.loaded_styles = yaml.safe_load(stream)

    self.configure_bullets(document, self.loaded_styles["bullet_color_hex"])

    for style in self.loaded_styles["styles"]:
      self.load_style(document, self.loaded_styles["styles"][style])

    for style in self.loaded_styles["styles"]:
      self.set_style_defaults(document)

  def set_style_defaults(self, document):
    for style_name in self.loaded_styles["styles"]:
      style = self.loaded_styles["styles"][style_name]
      style_base = style.get("base_style")
      if style_base:
        style_base = self.loaded_styles["styles"][style_base]
        doc_style = document.styles[style["style_name"]]
        doc_base_style = document.styles[style_base["style_name"]]
        doc_style.base_style = doc_base_style

  def load_style(self, document, style):
    # Try get the style, if it doesn't exist we'll make a new one
    try: doc_style = document.styles[style["style_name"]]
    except: 
      if self._get_attribute(style, "style_type") == "character":
        doc_style = document.styles.add_style(style["style_name"], WD_STYLE_TYPE.CHARACTER)
      else:
        doc_style = document.styles.add_style(style["style_name"], WD_STYLE_TYPE.PARAGRAPH)

    # font_name: Calibri Light
    doc_style.font.name = self._get_attribute(style, "font_name")
    rFonts = doc_style.element.rPr.rFonts
    if rFonts:
      rFonts.set(qn("w:asciiTheme"), self._get_attribute(style, "font_name"))

    # color: ["0x3B", "0x38", "0x38"]
    font_color = self._get_attribute(style, "color")
    doc_style.font.color.rgb = RGBColor(font_color[0], font_color[1], font_color[2])
    # size: 10
    doc_style.font.size = Pt(self._get_attribute(style, "size"))
    # bold: false
    doc_style.font.bold = self._get_attribute(style, "bold")
    # italic: false
    doc_style.font.italic = self._get_attribute(style, "italic")
    # underline: false
    doc_style.font.underline = self._get_attribute(style, "underline")
    # all_caps: false
    doc_style.font.all_caps = self._get_attribute(style, "all_caps")
    # small_caps: false
    doc_style.font.small_caps = self._get_attribute(style, "small_caps")
    # subscript: false
    doc_style.font.subscript = self._get_attribute(style, "subscript")
    # superscript: false
    doc_style.font.superscript = self._get_attribute(style, "superscript")
    # no_proof: false
    doc_style.font.no_proof = self._get_attribute(style, "no_proof")


  def configure_bullets(self, document, color_hex):
    num_xml = document.part.numbering_part.numbering_definitions._numbering
    num_1 = num_xml.num_having_numId(1)
    abstract_id = num_1.abstractNumId.val
    element = document.part.numbering_part.element.xpath(f"//w:abstractNum[@w:abstractNumId={abstract_id}]/w:lvl/w:rPr")[0]
    color = OxmlElement('w:color')
    color.set(qn('w:val'), color_hex) # 2E74B5
    element.insert(0,color)
