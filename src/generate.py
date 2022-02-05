import sys
import util
import jmespath
from docx import Document
from docx.enum.section import WD_SECTION
from docx.shared import Pt
from resume_style import margins


def build_header(document, header_data):
  default_section = document.sections[0]
  header = default_section.header.paragraphs[0]

  default_section.top_margin = margins.get("top")
  default_section.right_margin = margins.get("right")
  default_section.bottom_margin = margins.get("bottom")
  default_section.left_margin = margins.get("left")

  first,last = header_data["title"].split()
  city = header_data["location"]["city"]
  state = header_data["location"]["state"]
  zip = header_data["location"]["zip"]

  contacts = []
  for contact in header_data["contacts"]:
    contacts.append(f"{contact['name'].upper()} {contact['value']}")

  util.format_font("header_title_1", header.add_run(f"{first} ").font)
  util.format_font("header_title_2", header.add_run(last).font)
  util.format_font("header_subtitle", header.add_run(f"\n{city} • {state} • {zip}").font)
  util.format_font("header_subtitle", header.add_run(f"\n{' • '.join(contacts)}").font)


def build_footer(document, footer_data):
  default_section = document.sections[0]
  footer = default_section.footer.paragraphs[0]

  contacts = []
  for contact in footer_data["contacts"]:
    contacts.append(f"{contact['name'].upper()} {contact['value']}")

  util.format_font("header_subtitle", footer.add_run(f"\n{' • '.join(contacts)}").font)


def build_summary(document, summary_data):
  document.add_heading('SUMMARY')
  summary = document.add_paragraph(summary_data)


def build_highlights(document, highlights_data):
  util.insert_columns_section(document, WD_SECTION.CONTINUOUS, 2)
  document.add_heading(highlights_data["skillset"]["title"].upper())
  for skill in highlights_data["skillset"]["skills"]:
    document.add_paragraph(
      skill,
      style='List Bullet'
    )

  document.add_heading(highlights_data["personal_projects"]["title"].upper())
  for project in highlights_data["personal_projects"]["projects"]:
    document.add_paragraph(
      f"{project['name']} - {project['url']}\n{project['description']}",
      style='List Bullet'
    )


def build_roles(document, role_data):
  util.insert_columns_section(document, WD_SECTION.CONTINUOUS, 1)

  # Try build the volunteer roles
  document.add_heading("VOLUNTEER ROLES")
  volunteer_role_expression = jmespath.compile("[?type=='volunteer']")
  volunteer_roles = volunteer_role_expression.search(role_data)
  for role in volunteer_roles:
    build_role_positions(document, role)

  # Try build the professional roles
  document.add_heading("PROFESSIONAL ROLES")
  professional_role_expression = jmespath.compile("[?type=='professional']")
  professional_roles = professional_role_expression.search(role_data)
  for role in professional_roles:
    build_role_positions(document, role)


def build_role_positions(document, role_data):
  org_name = role_data["organization"]["name"]
  org_city = role_data["organization"]["location"]["city"]
  org_state = role_data["organization"]["location"]["state"]
  
  if role_data["organization"]["location"]["type"] == "remote": remote = True
  else: remote = False

  for position in role_data["positions"]:
    pos_start = f"{position['dates']['start']['month']}/{position['dates']['start']['year']}"
    pos_end = f"{position['dates']['end']['month']}/{position['dates']['end']['year']}"
    pos_title = position["title"]

    position_paragraph = document.add_paragraph()
    position_paragraph.add_run(org_name).bold = True
    position_paragraph.add_run(f", {org_city}, {org_state} • {pos_start} - {pos_end}\n")
    position_paragraph.add_run(pos_title).bold = True
    position_paragraph.paragraph_format.space_after = Pt(1)
    position_desc = document.add_paragraph(
      position["accomplishments"]+position["duties"],
      style='List Bullet'
    )
    position_desc.paragraph_format.keep_together = True
    position_desc.paragraph_format.space_after = Pt(5)

if __name__ == "__main__":
  # Make sure a data file has been specified.
  # TODO: Maybe use something like Click for input args
  if(len(sys.argv) == 1):
    print("Please specify a data file")
    sys.exit(2)

  # Get the resume data
  resume_data = util.load_resume_data(sys.argv[1])
  
  # Create the document object
  document = Document()

  # Try build the header
  header_data = resume_data.get("header")
  if header_data is None:
    raise Exception("No header found!")
  build_header(document, header_data)

  # Try build the summary
  summary_data = resume_data.get("summary")
  if summary_data is None:
    raise Exception("No summary found!")
  build_summary(document, summary_data)
  
  # Try build the highlights
  highlights_data = resume_data.get("highlights")
  if highlights_data is None:
    raise Exception("No highlights found!")
  build_highlights(document, highlights_data)

  # Try build the roles
  roles_data = resume_data.get("roles")
  if roles_data is None:
    raise Exception("No roles found!")
  build_roles(document, roles_data)
  
  # Try build the footer
  footer_data = resume_data.get("header")
  if footer_data is None:
    raise Exception("No footer found!")
  build_footer(document, footer_data)

  # Save the resume
  # TODO: Make the name configurable
  document.save("resume.docx")