# Resume — Rahul Chaudhary

ATS-friendly resume for **Rahul Chaudhary** — Generative AI Engineer | Full Stack Developer.
Targeted at Generative AI Engineer, Full Stack Developer, and Backend Developer roles (internships / entry-level).

## Files

- **`Rahul_Chaudhary_Resume.docx`** — the primary, **Microsoft Word–editable** resume. Based on the
  CCET template layout with **sky-blue accents (`#0EA5E9`)**, a **photo placeholder**, an education table,
  and shaded section bars. Just open it in Word to edit text or drop in a photo.
- **`build_resume.py`** — the Python (`python-docx`) generator that produces the `.docx`. Edit content here
  and re-run to regenerate a pixel-consistent document.
- **`index.html`** — a self-contained web/PDF version (HTML + inline CSS, no dependencies) with the same content.

## Regenerate the Word document

```bash
pip install python-docx
python3 build_resume.py     # -> Rahul_Chaudhary_Resume.docx
```

## Add a profile photo (in Word)

1. Open `Rahul_Chaudhary_Resume.docx` in Microsoft Word.
2. Click inside the dashed **PHOTO** box in the top-left header cell.
3. Delete the placeholder text and use **Insert → Pictures** to add your photo.

> Note: photos are optional and can slightly reduce ATS parsing accuracy. Keep it if a photo is
> expected for your region/roles; otherwise the layout looks clean without it.

## Export the HTML version to PDF

1. Open `index.html` in Chrome or Edge.
2. Press `Ctrl/Cmd + P`, set **Destination** to *Save as PDF*.
3. Set **Margins** to *None* and enable **Background graphics** so the sky-blue accents render.

## ATS notes

Content uses standard section headings and real, selectable text (not images), with natural keyword
coverage for Generative AI, Agentic AI, RAG, full stack (React / Node.js / Express), backend, REST APIs,
authentication, MongoDB / PostgreSQL / Supabase, workflow automation (n8n), computer vision, and
embedded AI. No keyword stuffing, no fabricated metrics.
