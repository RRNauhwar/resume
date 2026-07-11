#!/usr/bin/env python3
"""
Generate an ATS-friendly, Word-editable ONE-PAGE resume for Rahul Chaudhary
that strictly follows the CCET (Chandigarh College of Engineering & Technology)
template layout.

Template fidelity
-----------------
* Header: CCET logo (left) | Name / Branch / College (center) | Contact (right)
* Thick horizontal rule under the header
* Education rendered as a bordered table (Qualification | Institute | Year | Score)
* Full-width shaded section bars for every section
* Serif typography (Times New Roman), matching the classic template look
* Sky-blue accent on the section bars

Layout is tuned to match index.html, whose one-page fit was verified with
WeasyPrint (content height ~9.76in vs 10.2in usable => ~0.44in slack).

Run:  python build_resume.py   ->   Rahul_Chaudhary_Resume.docx
"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ---------------------------------------------------------------- palette / fonts
ACCENT_DARK  = RGBColor(0x03, 0x69, 0xA1)
INK          = RGBColor(0x1A, 0x1A, 0x1A)
MUTED        = RGBColor(0x33, 0x33, 0x33)
WHITE        = RGBColor(0xFF, 0xFF, 0xFF)
ACCENT_HEX      = "0EA5E9"
ACCENT_DARK_HEX = "0369A1"
FONT = "Times New Roman"
LOGO = "ccet_logo.png"

BODY = 9.0    # base body size (pt)
LINEH = 10.3  # fixed line height for body text (pt), kept in sync with index.html


# ---------------------------------------------------------------- helpers
def set_cell_bg(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def set_cell_borders(cell, **kwargs):
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


def set_cell_margins(cell, top=16, bottom=16, left=80, right=80):
    tcPr = cell._tc.get_or_add_tcPr()
    m = OxmlElement("w:tcMar")
    for side, val in (("top", top), ("bottom", bottom), ("start", left), ("end", right)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(val)); el.set(qn("w:type"), "dxa")
        m.append(el)
    tcPr.append(m)


def no_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "none"); el.set(qn("w:sz"), "0")
        borders.append(el)
    tblPr.append(borders)


def add_run(p, text, size=BODY, bold=False, italic=False, color=INK, font=FONT):
    r = p.add_run(text)
    r.bold = bold; r.italic = italic
    r.font.size = Pt(size); r.font.color.rgb = color; r.font.name = font
    rpr = r._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts"); rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), font); rfonts.set(qn("w:hAnsi"), font)
    return r


def space(p, before=0, after=0, line=1.0, exact=None):
    pf = p.paragraph_format
    pf.space_before = Pt(before); pf.space_after = Pt(after)
    if exact is not None:
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(exact)
    else:
        pf.line_spacing = line


def section_header(doc, title):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    no_table_borders(tbl)
    cell = tbl.rows[0].cells[0]
    set_cell_bg(cell, ACCENT_HEX)
    set_cell_margins(cell, top=8, bottom=8, left=100, right=100)
    p = cell.paragraphs[0]
    space(p, before=0, after=0, line=1.0)
    add_run(p, title.upper(), size=10, bold=True, color=WHITE)
    sp = doc.add_paragraph()
    space(sp, before=0, after=1, exact=3)
    return tbl


def bullet(doc, text, size=BODY, after=1.0, indent=0.22):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Inches(indent); pf.first_line_indent = Inches(-0.14)
    space(p, before=0, after=after, exact=LINEH)
    add_run(p, "\u2022  ", size=BODY, bold=True, color=ACCENT_DARK)
    add_run(p, text, size=size, color=MUTED)
    return p


# ---------------------------------------------------------------- document setup
doc = Document()
sec = doc.sections[0]
sec.top_margin = Inches(0.4); sec.bottom_margin = Inches(0.4)
sec.left_margin = Inches(0.5); sec.right_margin = Inches(0.5)

normal = doc.styles["Normal"]
normal.font.name = FONT; normal.font.size = Pt(BODY); normal.font.color.rgb = INK
normal.paragraph_format.space_after = Pt(0); normal.paragraph_format.line_spacing = 1.0


# ---------------------------------------------------------------- HEADER
head = doc.add_table(rows=1, cols=3)
no_table_borders(head)
head.autofit = False; head.allow_autofit = False
widths = [Inches(3.90), Inches(2.45), Inches(1.15)]
for i, w in enumerate(widths):
    head.columns[i].width = w
# order: identity (left) | contact (middle) | logo (right)
id_cell, contact_cell, logo_cell = head.rows[0].cells
for cell, w in zip((id_cell, contact_cell, logo_cell), widths):
    cell.width = w; cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

set_cell_margins(id_cell, top=0, bottom=0, left=0, right=50)
n = id_cell.paragraphs[0]; space(n, before=0, after=0, line=1.0)
add_run(n, "RAHUL CHAUDHARY", size=19, bold=True, color=INK)
br = id_cell.add_paragraph(); space(br, before=1, after=0, line=1.0)
add_run(br, "Generative AI Engineer  |  Full Stack Developer", size=10, bold=True, color=ACCENT_DARK)
b2 = id_cell.add_paragraph(); space(b2, before=1, after=0, line=1.0)
add_run(b2, "B.E. Electronics & Communication Engineering", size=8.5, color=MUTED)
b3 = id_cell.add_paragraph(); space(b3, before=0, after=0, line=1.0)
add_run(b3, "Chandigarh College of Engineering & Technology", size=8.5, color=MUTED)

set_cell_margins(contact_cell, top=0, bottom=0, left=40, right=40)
contacts = [
    ("Phone: ", "+91 78190 22307"),
    ("Email: ", "rahulch19905@gmail.com"),
    ("LinkedIn: ", "in/rahul-chaudhary-9a7b82310"),
    ("GitHub: ", "github.com/RRNauhwar"),
    ("Location: ", "Chandigarh, India"),
]
for idx, (label, val) in enumerate(contacts):
    cp = contact_cell.paragraphs[0] if idx == 0 else contact_cell.add_paragraph()
    space(cp, before=0, after=0, line=1.08)
    add_run(cp, label, size=8, bold=True, color=ACCENT_DARK)
    add_run(cp, val, size=8, color=MUTED)

# logo on the right of the contact information
set_cell_margins(logo_cell, top=0, bottom=0, left=40, right=0)
lp = logo_cell.paragraphs[0]; lp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
space(lp, before=0, after=0, line=1.0)
lp.add_run().add_picture(LOGO, width=Inches(1.0))

rule = doc.add_paragraph(); space(rule, before=3, after=3, line=1.0)
pPr = rule._p.get_or_add_pPr()
pbdr = OxmlElement("w:pBdr"); bottom = OxmlElement("w:bottom")
bottom.set(qn("w:val"), "single"); bottom.set(qn("w:sz"), "20")
bottom.set(qn("w:space"), "1"); bottom.set(qn("w:color"), ACCENT_DARK_HEX)
pbdr.append(bottom); pPr.append(pbdr)


# ---------------------------------------------------------------- EDUCATION
section_header(doc, "Education")
edu = doc.add_table(rows=1, cols=4)
edu.alignment = WD_TABLE_ALIGNMENT.CENTER
edu.autofit = False; edu.allow_autofit = False
ew = [Inches(3.05), Inches(2.85), Inches(0.85), Inches(0.75)]
for i, (w, h) in enumerate(zip(ew, ["Qualification", "Institute", "Year", "Score"])):
    c = edu.rows[0].cells[i]; c.width = w
    set_cell_bg(c, ACCENT_DARK_HEX); set_cell_margins(c, top=10, bottom=10, left=90, right=90)
    p = c.paragraphs[0]; space(p, before=0, after=0, line=1.0)
    add_run(p, h, size=8.5, bold=True, color=WHITE)
edu_rows = [
    ("B.E., Electronics & Communication Engineering",
     "Chandigarh College of Engineering & Technology", "2023 \u2013 2027*", "7.37"),
    ("Senior Secondary (Class XII), CBSE", "Kanha Makhan Public School", "2021", "82%"),
    ("Secondary (Class X), CBSE", "SRBS International School", "2019", "91.8%"),
]
for row in edu_rows:
    cells = edu.add_row().cells
    for i, (val, w) in enumerate(zip(row, ew)):
        c = cells[i]; c.width = w
        set_cell_margins(c, top=10, bottom=10, left=90, right=90)
        set_cell_borders(c, bottom={"sz": 4, "val": "single", "color": "D0D7DE"})
        p = c.paragraphs[0]; space(p, before=0, after=0, line=1.0)
        if i == 3:
            add_run(p, val, size=8.5, bold=True, color=ACCENT_DARK)
        else:
            add_run(p, val, size=8.5, bold=(i == 0), color=(INK if i == 0 else MUTED))
gp = doc.add_paragraph(); space(gp, before=1, after=1, exact=8)
add_run(gp, "*Expected graduation", size=7.5, italic=True, color=MUTED)


# ---------------------------------------------------------------- PROFESSIONAL SUMMARY
section_header(doc, "Professional Summary")
s = doc.add_paragraph(); s.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
space(s, before=0, after=2, exact=LINEH)
add_run(
    s,
    "Generative AI Engineer and Full Stack Developer specialising in the design and development of "
    "intelligent, scalable software solutions. Experienced in building AI-driven applications, LLM-powered "
    "systems, RAG architectures, and agentic workflows, combined with robust full-stack platforms and "
    "backend services. Skilled in developing secure, high-performance applications with a focus on "
    "automation, system reliability, and solving complex real-world problems through emerging technologies.",
    size=BODY, color=MUTED,
)


# ---------------------------------------------------------------- TECHNICAL SKILLS
section_header(doc, "Technical Skills")
skills = [
    ("Languages", "Python, JavaScript, Java, C++"),
    ("Web & Databases", "React, Node.js, Express.js, REST APIs, MongoDB, PostgreSQL, Supabase"),
    ("Generative AI", "OpenAI API, Gemini API, Prompt Engineering, RAG, Agentic AI, Workflow Automation"),
    ("Tools & Embedded", "Git, GitHub, Docker, Postman, n8n, Raspberry Pi, Computer Vision, YOLO"),
]
for label, items in skills:
    p = doc.add_paragraph(); p.paragraph_format.left_indent = Inches(0.05)
    space(p, before=0, after=1, exact=LINEH)
    add_run(p, f"{label}:  ", size=BODY, bold=True, color=ACCENT_DARK)
    add_run(p, items, size=BODY, color=MUTED)


# ---------------------------------------------------------------- EXPERIENCE
section_header(doc, "Experience")
eh = doc.add_paragraph(); space(eh, before=0, after=1, exact=BODY + 3)
eh.paragraph_format.tab_stops.add_tab_stop(Inches(7.5), alignment=2)
add_run(eh, "AI Automation Intern", size=9.5, bold=True, color=INK)
add_run(eh, "  |  D2 Automation", size=9.5, bold=True, color=ACCENT_DARK)
add_run(eh, "\t", size=BODY)
add_run(eh, "June 2025 \u2013 July 2025", size=8.5, italic=True, color=MUTED)
for b in [
    "Built AI automation workflows in n8n and integrated the OpenAI and Gemini APIs to orchestrate multi-step, LLM-driven business processes.",
    "Developed a Telegram chatbot and Python automation scripts to handle user queries and eliminate repetitive manual tasks through API integrations.",
]:
    bullet(doc, b)


# ---------------------------------------------------------------- PROJECTS
section_header(doc, "Projects")


def project(name, subtitle, stack, bullets):
    ph = doc.add_paragraph(); space(ph, before=2, after=0, exact=BODY + 3)
    add_run(ph, name, size=9.3, bold=True, color=INK)
    add_run(ph, f"  \u2013  {subtitle}", size=BODY, italic=True, color=ACCENT_DARK)
    sp = doc.add_paragraph(); sp.paragraph_format.left_indent = Inches(0.05)
    space(sp, before=0, after=1, exact=9)
    add_run(sp, "Tech Stack: ", size=8, bold=True, color=INK)
    add_run(sp, stack, size=8, italic=True, color=MUTED)
    for b in bullets:
        bullet(doc, b)


project(
    "NyayaSim", "AI-Powered Virtual Courtroom",
    "React, Node.js, Express.js, Supabase, PostgreSQL, OpenAI API, Gemini API",
    [
        "Built an AI-powered virtual courtroom platform that enables law students to practice real-world litigation through interactive trial simulations.",
        "Developed intelligent courtroom participants, including AI judges, advocates, and witnesses, along with case management, evidence handling, and legal document drafting workflows.",
        "Designed a scalable full-stack architecture to bridge the gap between classroom legal education and practical courtroom advocacy using modern AI technologies.",
    ],
)
project(
    "Smart Glasses for the Visually Impaired", "Embedded AI & Computer Vision",
    "Python, Raspberry Pi, YOLO, Computer Vision, Text-to-Speech",
    [
        "Developed AI-powered smart glasses designed to assist visually impaired individuals by providing real-time object detection and environmental awareness.",
        "Integrated embedded hardware including Raspberry Pi, camera modules, sensors, and audio feedback systems with computer vision models for intelligent assistance.",
        "Built an edge AI pipeline capable of processing visual information and delivering real-time voice-based insights to improve user independence and accessibility.",
    ],
)
project(
    "DealHunt", "AI-Powered Deal Discovery Platform",
    "React, Node.js, Express.js, MongoDB, Firebase",
    [
        "Built an AI-powered e-commerce platform that aggregates products from multiple online retailers, enabling users to compare prices and discover the best deals.",
        "Implemented intelligent search, personalized product recommendations, and real-time deal discovery to enhance the shopping experience.",
    ],
)
project(
    "AI Telegram News Assistant", "Automated AI News Pipeline",
    "Python, n8n, Telegram Bot API, OpenAI API",
    [
        "Engineered an Agentic AI Telegram bot that fetches, summarizes, and categorizes news from multiple sources using LLM-driven processing.",
        "Automated scheduling and personalized delivery through event-driven n8n workflows.",
    ],
)


# ---------------------------------------------------------------- CERTIFICATIONS
section_header(doc, "Certifications")
certs = [
    ("100 Days of Code: The Complete Python Pro Bootcamp", "Dr. Angela Yu, Udemy"),
    ("Full Stack Generative & Agentic AI with Python", "Hitesh Choudhary, Udemy"),
    ("Complete Web Development Course", "Hitesh Choudhary, Udemy"),
    ("DEMA: Drone Engineering, Mechanics & Applications Bootcamp",
     "NIT Jalandhar & CCET (MeitY, Govt. of India) \u2013 UAV & embedded systems; scored 27/30"),
]
for title, provider in certs:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.22); p.paragraph_format.first_line_indent = Inches(-0.14)
    space(p, before=0, after=1, exact=LINEH)
    add_run(p, "\u2022  ", size=BODY, bold=True, color=ACCENT_DARK)
    add_run(p, f"{title} ", size=BODY, bold=True, color=INK)
    add_run(p, f"\u2013 {provider}", size=8.5, color=MUTED)


# ---------------------------------------------------------------- ACHIEVEMENTS
section_header(doc, "Achievements")
for a in [
    "Qualified the GATE Computer Science examination while pursuing a B.E. in Electronics & Communication Engineering.",
    "Solved 170+ Data Structures & Algorithms problems on GeeksforGeeks, ranking 16th on the institute leaderboard.",
]:
    bullet(doc, a)


# ---------------------------------------------------------------- INTERESTS & ADDITIONAL
section_header(doc, "Technical Interests & Additional Information")
for label, val in [
    ("Technical Interests:  ", "Generative AI, AI Automation, Agentic AI, Backend System Design"),
    ("Extra-Curricular:  ", "Chess, Space & Astronomy"),
    ("Languages:  ", "English, Hindi, German (Beginner)"),
]:
    p = doc.add_paragraph(); p.paragraph_format.left_indent = Inches(0.05)
    space(p, before=0, after=1, exact=LINEH)
    add_run(p, label, size=BODY, bold=True, color=ACCENT_DARK)
    add_run(p, val, size=BODY, color=MUTED)


# ---------------------------------------------------------------- save
out = "Rahul_Chaudhary_Resume.docx"
doc.save(out)
print(f"Saved {out}")
