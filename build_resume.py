#!/usr/bin/env python3
"""
Generate an ATS-friendly, Word-editable resume for Rahul Chaudhary.

Design goals
------------
* Based on the CCET (Chandigarh College of Engineering & Technology) template:
  - Header block (name / title / contact) with a photo placeholder
  - Education rendered as a clean table
  - Full-width shaded section header bars
* Sky-blue accents (#0EA5E9) instead of the template's grey bars
* Single, selectable-text, standard-heading layout => high ATS compatibility
* Comfortable whitespace, easy to edit in Microsoft Word

Run:  python build_resume.py
Out:  Rahul_Chaudhary_Resume.docx
"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ---------------------------------------------------------------- palette / fonts
ACCENT       = RGBColor(0x0E, 0xA5, 0xE9)   # sky blue
ACCENT_DARK  = RGBColor(0x03, 0x69, 0xA1)   # deep sky blue
INK          = RGBColor(0x1F, 0x29, 0x37)   # near-black text
MUTED        = RGBColor(0x4B, 0x55, 0x63)   # grey body text
WHITE        = RGBColor(0xFF, 0xFF, 0xFF)

ACCENT_HEX      = "0EA5E9"
ACCENT_DARK_HEX = "0369A1"
SOFT_HEX        = "E0F2FE"                    # very light sky blue (table header / photo box)

BODY_FONT = "Calibri"     # clean, ubiquitous, ATS-safe
NAME_FONT = "Calibri"


# ---------------------------------------------------------------- low-level helpers
def set_cell_bg(cell, hex_color):
    """Shade a table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def set_cell_border(cell, **kwargs):
    """kwargs: top/bottom/left/right -> dict(sz, val, color)."""
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        if edge in kwargs:
            spec = kwargs[edge]
            el = OxmlElement(f"w:{edge}")
            el.set(qn("w:val"), spec.get("val", "single"))
            el.set(qn("w:sz"), str(spec.get("sz", 4)))
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), spec.get("color", "auto"))
            borders.append(el)
    tcPr.append(borders)


def set_cell_margins(cell, top=40, bottom=40, left=80, right=80):
    """Cell inner margins in twips (1/20 pt)."""
    tcPr = cell._tc.get_or_add_tcPr()
    m = OxmlElement("w:tcMar")
    for side, val in (("top", top), ("bottom", bottom), ("start", left), ("end", right)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        m.append(el)
    tcPr.append(m)


def vertical_center(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    va = OxmlElement("w:vAlign")
    va.set(qn("w:val"), "center")
    tcPr.append(va)


def no_table_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        borders.append(el)
    tblPr.append(borders)


def add_run(p, text, size=10.5, bold=False, italic=False, color=INK, font=BODY_FONT):
    r = p.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size)
    r.font.color.rgb = color
    r.font.name = font
    # ensure east-asian / complex scripts also use the font
    rpr = r._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), font)
    rfonts.set(qn("w:hAnsi"), font)
    return r


def space(p, before=0, after=0, line=1.0):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line


def section_header(doc, title):
    """Full-width sky-blue shaded bar with white uppercase title (template style)."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    no_table_borders(tbl)
    cell = tbl.rows[0].cells[0]
    set_cell_bg(cell, ACCENT_HEX)
    set_cell_margins(cell, top=30, bottom=30, left=110, right=110)
    p = cell.paragraphs[0]
    space(p, before=0, after=0, line=1.0)
    add_run(p, title.upper(), size=11, bold=True, color=WHITE)
    # tiny spacer after the bar
    sp = doc.add_paragraph()
    space(sp, before=0, after=2, line=1.0)
    sp.runs and None
    return tbl


def bullet(doc, segments, space_after=2.5, indent=0.28):
    """segments = list of (text, kwargs) tuples rendered on one bulleted line."""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Inches(indent)
    pf.first_line_indent = Inches(-0.16)
    space(p, before=0, after=space_after, line=1.06)
    add_run(p, "\u2022  ", size=10.5, bold=True, color=ACCENT)
    for text, kw in segments:
        add_run(p, text, **kw)
    return p


def simple_bullet(doc, text, **kw):
    return bullet(doc, [(text, {"size": 10.5, "color": MUTED, **kw})])


# ---------------------------------------------------------------- document setup
doc = Document()

# page margins (A4-ish, comfortable)
sec = doc.sections[0]
sec.top_margin = Inches(0.5)
sec.bottom_margin = Inches(0.5)
sec.left_margin = Inches(0.6)
sec.right_margin = Inches(0.6)

# base style
normal = doc.styles["Normal"]
normal.font.name = BODY_FONT
normal.font.size = Pt(10.5)
normal.font.color.rgb = INK
normal.paragraph_format.space_after = Pt(0)
normal.paragraph_format.line_spacing = 1.06


# ---------------------------------------------------------------- HEADER
# Two columns: left = photo placeholder, right = name / title / contact
head = doc.add_table(rows=1, cols=2)
no_table_borders(head)
head.autofit = False
head.columns[0].width = Inches(1.25)
head.columns[1].width = Inches(6.05)

photo_cell, info_cell = head.rows[0].cells
photo_cell.width = Inches(1.25)
info_cell.width = Inches(6.05)

# --- photo placeholder box
set_cell_bg(photo_cell, SOFT_HEX)
set_cell_border(
    photo_cell,
    top={"sz": 8, "val": "dashed", "color": ACCENT_HEX},
    bottom={"sz": 8, "val": "dashed", "color": ACCENT_HEX},
    left={"sz": 8, "val": "dashed", "color": ACCENT_HEX},
    right={"sz": 8, "val": "dashed", "color": ACCENT_HEX},
)
vertical_center(photo_cell)
set_cell_margins(photo_cell, top=180, bottom=180, left=60, right=60)
pp = photo_cell.paragraphs[0]
pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
space(pp, before=0, after=0, line=1.05)
add_run(pp, "PHOTO", size=8.5, bold=True, color=ACCENT_DARK)
pp2 = photo_cell.add_paragraph()
pp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
space(pp2, before=0, after=0, line=1.05)
add_run(pp2, "Insert photo", size=7.5, color=ACCENT_DARK)

# --- name / title / contact
vertical_center(info_cell)
set_cell_margins(info_cell, top=20, bottom=20, left=180, right=40)

n = info_cell.paragraphs[0]
space(n, before=0, after=1, line=1.0)
add_run(n, "RAHUL CHAUDHARY", size=24, bold=True, color=ACCENT_DARK, font=NAME_FONT)

t = info_cell.add_paragraph()
space(t, before=0, after=5, line=1.0)
add_run(t, "Generative AI Engineer  |  Full Stack Developer", size=12, bold=True, color=INK)

c1 = info_cell.add_paragraph()
space(c1, before=0, after=1, line=1.12)
add_run(c1, "Chandigarh, India", size=9.5, color=MUTED)
add_run(c1, "   \u2022   ", size=9.5, color=ACCENT)
add_run(c1, "+91 78190 22307", size=9.5, color=MUTED)
add_run(c1, "   \u2022   ", size=9.5, color=ACCENT)
add_run(c1, "rahulch19905@gmail.com", size=9.5, color=MUTED)

c2 = info_cell.add_paragraph()
space(c2, before=0, after=0, line=1.12)
add_run(c2, "linkedin.com/in/rahul-chaudhary-9a7b82310", size=9.5, color=MUTED)
add_run(c2, "   \u2022   ", size=9.5, color=ACCENT)
add_run(c2, "github.com/RRNauhwar", size=9.5, color=MUTED)

# accent divider under header
div = doc.add_paragraph()
space(div, before=4, after=6, line=1.0)
pPr = div._p.get_or_add_pPr()
pbdr = OxmlElement("w:pBdr")
bottom = OxmlElement("w:bottom")
bottom.set(qn("w:val"), "single")
bottom.set(qn("w:sz"), "18")
bottom.set(qn("w:space"), "1")
bottom.set(qn("w:color"), ACCENT_HEX)
pbdr.append(bottom)
pPr.append(pbdr)


# ---------------------------------------------------------------- PROFESSIONAL SUMMARY
section_header(doc, "Professional Summary")
s = doc.add_paragraph()
s.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
space(s, before=0, after=6, line=1.12)
add_run(
    s,
    "Generative AI Engineer and Full Stack Developer who builds complete AI-powered products, "
    "from backend architecture to deployment. Experienced in developing LLM applications, RAG pipelines, "
    "and agentic AI workflows using the OpenAI and Gemini APIs, alongside production-ready React and "
    "Node.js platforms with secure authentication and scalable REST APIs. Skilled in backend engineering, "
    "workflow automation, and applied computer vision on embedded devices, with a focus on designing "
    "reliable systems that solve real-world problems.",
    size=10.5, color=MUTED,
)


# ---------------------------------------------------------------- EDUCATION (table)
section_header(doc, "Education")

edu = doc.add_table(rows=1, cols=4)
edu.alignment = WD_TABLE_ALIGNMENT.CENTER
edu.autofit = False
widths = [Inches(2.9), Inches(2.7), Inches(0.95), Inches(0.75)]
headers = ["Qualification", "Institute", "Year", "Score"]
for i, (w, h) in enumerate(zip(widths, headers)):
    cell = edu.rows[0].cells[i]
    cell.width = w
    set_cell_bg(cell, ACCENT_DARK_HEX)
    set_cell_margins(cell, top=40, bottom=40, left=90, right=90)
    p = cell.paragraphs[0]
    space(p, before=0, after=0, line=1.0)
    add_run(p, h, size=9.5, bold=True, color=WHITE)

edu_rows = [
    ("B.E., Electronics & Communication Engineering",
     "Chandigarh College of Engineering & Technology", "2023 \u2013 2027*", "7.37"),
    ("Senior Secondary (Class XII), CBSE",
     "Kanha Makhan Public School", "2021", "82%"),
    ("Secondary (Class X), CBSE",
     "SRBS International School", "2019", "91.8%"),
]
for r_i, row in enumerate(edu_rows):
    cells = edu.add_row().cells
    fill = "FFFFFF" if r_i % 2 == 0 else SOFT_HEX
    for i, (val, w) in enumerate(zip(row, widths)):
        cell = cells[i]
        cell.width = w
        set_cell_bg(cell, fill)
        set_cell_margins(cell, top=40, bottom=40, left=90, right=90)
        p = cell.paragraphs[0]
        space(p, before=0, after=0, line=1.04)
        bold = (i == 0)
        col = INK if i == 0 else MUTED
        if i == 3:
            add_run(p, val, size=9.5, bold=True, color=ACCENT_DARK)
        else:
            add_run(p, val, size=9.5, bold=bold, color=col)

note = doc.add_paragraph()
space(note, before=2, after=4, line=1.0)
add_run(note, "*Expected graduation", size=8, italic=True, color=MUTED)


# ---------------------------------------------------------------- TECHNICAL SKILLS
section_header(doc, "Technical Skills")

skills = [
    ("Programming Languages", "Python, JavaScript, Java, C++"),
    ("Web Development", "React, Node.js, Express.js, REST APIs, HTML, CSS"),
    ("Databases", "MongoDB, PostgreSQL, Supabase"),
    ("Generative AI", "OpenAI API, Gemini API, Prompt Engineering, RAG, Agentic AI, Workflow Automation"),
    ("Tools & Platforms", "Git, GitHub, Docker, Postman, n8n"),
    ("Embedded & Vision", "Raspberry Pi, Computer Vision, YOLO, Object Detection"),
]
for label, items in skills:
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Inches(0.1)
    space(p, before=0, after=3, line=1.1)
    add_run(p, f"{label}:  ", size=10, bold=True, color=ACCENT_DARK)
    add_run(p, items, size=10, color=MUTED)


# ---------------------------------------------------------------- EXPERIENCE
section_header(doc, "Experience")

eh = doc.add_paragraph()
space(eh, before=0, after=1, line=1.05)
add_run(eh, "AI Automation Intern", size=11, bold=True, color=INK)
add_run(eh, "  |  D2 Automation", size=10.5, bold=True, color=ACCENT_DARK)
# right-aligned date via tab
tabs = eh.paragraph_format
add_run(eh, "\t", size=10.5)
# add a right tab stop
tab_stops = eh.paragraph_format.tab_stops
tab_stops.add_tab_stop(Inches(7.2), alignment=3)  # 3 = right
add_run(eh, "June 2025 \u2013 July 2025", size=9.5, italic=True, color=MUTED)

exp_bullets = [
    "Designed and deployed AI automation workflows in n8n to orchestrate multi-step business processes and eliminate repetitive manual tasks.",
    "Built a Telegram chatbot integrated with the OpenAI and Gemini APIs to handle user queries and deliver automated conversational responses.",
    "Developed Python scripts and API integrations connecting third-party services and messaging platforms for event-driven task execution.",
    "Engineered reusable automation pipelines, improving reliability and reducing turnaround time across internal operations.",
]
for b in exp_bullets:
    simple_bullet(doc, b)


# ---------------------------------------------------------------- PROJECTS
section_header(doc, "Projects")


def project(name, subtitle, stack, bullets):
    ph = doc.add_paragraph()
    space(ph, before=3, after=1, line=1.05)
    add_run(ph, name, size=10.8, bold=True, color=INK)
    if subtitle:
        add_run(ph, f"  \u2013  {subtitle}", size=10, color=ACCENT_DARK)
    sp = doc.add_paragraph()
    sp.paragraph_format.left_indent = Inches(0.1)
    space(sp, before=0, after=2, line=1.05)
    add_run(sp, "Tech Stack: ", size=9, bold=True, color=INK)
    add_run(sp, stack, size=9, italic=True, color=MUTED)
    for b in bullets:
        simple_bullet(doc, b)


project(
    "NyayaSim", "AI-Powered Virtual Courtroom",
    "React, Node.js, Express.js, Supabase, PostgreSQL, OpenAI API, Gemini API",
    [
        "Developed an AI-powered virtual courtroom platform leveraging Generative and Agentic AI to simulate realistic legal proceedings.",
        "Engineered secure authentication, case management, and digital evidence modules backed by Supabase and PostgreSQL.",
        "Designed scalable REST APIs following Clean Architecture and SOLID principles for maintainable backend services.",
        "Integrated OpenAI and Gemini models to simulate judges, witnesses, and legal arguments through coordinated AI agents.",
    ],
)

project(
    "Smart Glasses for the Visually Impaired", "Embedded AI & Computer Vision",
    "Raspberry Pi, Python, YOLO, Computer Vision, Text-to-Speech",
    [
        "Built a real-time object detection system using YOLO on Raspberry Pi to identify surroundings and support safe navigation.",
        "Implemented an agentic decision layer that prioritizes hazards by proximity and context to deliver relevant guidance.",
        "Converted visual detections into spoken audio cues, enabling hands-free assistive feedback for visually impaired users.",
        "Optimized the computer vision pipeline for low latency and power efficiency on a resource-constrained edge device.",
    ],
)

project(
    "DealHunt", "AI-Powered Deal Discovery Platform",
    "React, Node.js, Express.js, MongoDB, Firebase",
    [
        "Built a full-stack platform for discovering and sharing local and online deals within an engaged user community.",
        "Implemented secure authentication, user profiles, and deal management features across the stack.",
        "Integrated AI-powered personalized deal recommendations based on user preferences and behavior.",
        "Designed responsive interfaces and scalable backend REST APIs for a seamless user experience.",
    ],
)

project(
    "AI Telegram News Assistant", "Automated AI News Pipeline",
    "Python, n8n, Telegram Bot API, OpenAI API, Gemini API",
    [
        "Developed an Agentic AI Telegram bot that automates personalized news delivery from multiple sources.",
        "Built AI workflows to fetch, summarize, and categorize news using LLM-driven processing.",
        "Automated scheduling and content delivery through event-driven n8n workflows.",
        "Leveraged LLMs to generate concise, accurate, and reader-friendly news summaries.",
    ],
)


# ---------------------------------------------------------------- CERTIFICATIONS
section_header(doc, "Certifications")
certs = [
    ("Full Stack Generative & Agentic AI with Python", "Hitesh Choudhary \u2013 Udemy",
     "Building LLM applications, RAG pipelines, and AI agents with Python."),
    ("100 Days of Code: The Complete Python Pro Bootcamp", "Dr. Angela Yu \u2013 Udemy",
     "100 hands-on Python projects spanning automation, APIs, and web apps."),
    ("Complete Web Development Course", "Hitesh Choudhary \u2013 Udemy",
     "End-to-end full stack web development, from front end to back end."),
]
for title, provider, desc in certs:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.28)
    p.paragraph_format.first_line_indent = Inches(-0.16)
    space(p, before=0, after=3, line=1.08)
    add_run(p, "\u2022  ", size=10.5, bold=True, color=ACCENT)
    add_run(p, f"{title} ", size=10, bold=True, color=INK)
    add_run(p, f"({provider}) \u2013 ", size=9.5, color=ACCENT_DARK)
    add_run(p, desc, size=9.5, color=MUTED)


# ---------------------------------------------------------------- ACHIEVEMENTS
section_header(doc, "Achievements")
achievements = [
    "Solved 170+ Data Structures & Algorithms problems on GeeksforGeeks.",
    "Participated in the Hackfinity 2025 Hackathon and built an AI-powered mental wellness application.",
    "Appeared for the GATE examination while pursuing B.E. in Electronics & Communication Engineering.",
]
for a in achievements:
    simple_bullet(doc, a)


# ---------------------------------------------------------------- LEADERSHIP & EXTRACURRICULAR
section_header(doc, "Leadership & Activities")
p = doc.add_paragraph()
p.paragraph_format.left_indent = Inches(0.1)
space(p, before=0, after=2, line=1.1)
add_run(p, "Member \u2013 Robotics Club, CCET", size=10, color=MUTED)
add_run(p, "      \u2022      ", size=10, color=ACCENT)
add_run(p, "Member \u2013 ACM Student Chapter", size=10, color=MUTED)

p2 = doc.add_paragraph()
p2.paragraph_format.left_indent = Inches(0.1)
space(p2, before=0, after=2, line=1.1)
add_run(p2, "Interests: ", size=10, bold=True, color=ACCENT_DARK)
add_run(p2, "Chess  \u2022  Volleyball", size=10, color=MUTED)


# ---------------------------------------------------------------- save
out = "Rahul_Chaudhary_Resume.docx"
doc.save(out)
print(f"Saved {out}")
