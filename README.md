# Resume — Rahul Chaudhary

ATS-friendly resume for **Rahul Chaudhary** — Generative AI Engineer | Full Stack Developer.
Targeted at Generative AI Engineer, Full Stack Developer, and Backend Developer roles (internships / entry-level).

## Files

- **`Rahul_Chaudhary_Resume.docx`** — the primary, **Microsoft Word–editable** resume. Clean, modern,
  single-column layout with **sky-blue accents (`#0EA5E9`)**, a centered header, an education table, and
  shaded section bars. No photo, no rating bars, no ATS-breaking graphics. Just open it in Word to edit.
- **`build_resume.py`** — the Python (`python-docx`) generator that produces the `.docx`. Edit content here
  and re-run to regenerate a pixel-consistent document.
- **`index.html`** — a self-contained web/PDF version (HTML + inline CSS, no dependencies) with the same content.

## Regenerate the Word document

```bash
pip install python-docx
python3 build_resume.py     # -> Rahul_Chaudhary_Resume.docx
```

## Export the HTML version to PDF

1. Open `index.html` in Chrome or Edge.
2. Press `Ctrl/Cmd + P`, set **Destination** to *Save as PDF*.
3. Set **Margins** to *None* and enable **Background graphics** so the sky-blue accents render.

## ATS notes

Content uses standard section headings and real, selectable text (not images), with natural keyword
coverage for Generative AI, AI Engineering, Agentic AI, RAG, full stack (React / Node.js / Express),
backend, REST APIs, authentication, MongoDB / PostgreSQL / Supabase, workflow automation (n8n),
computer vision, YOLO, Raspberry Pi, and Docker. No keyword stuffing, no fabricated metrics.
