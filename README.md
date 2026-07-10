# Resume — Rahul Chaudhary

One-page, ATS-friendly resume for **Rahul Chaudhary** — Generative AI Engineer | Full Stack Developer.
Targeted at Generative AI Engineer, AI Engineer, Full Stack Developer, and Backend Developer roles
(internships / entry-level). The layout **strictly follows the CCET template** (college logo header,
education table, full-width section bars) with a professional sky-blue accent.

## Files

- **`Rahul_Chaudhary_Resume.docx`** — the primary, **Microsoft Word–editable** resume. One page, serif
  (Times New Roman) typography, CCET logo in the header, education table, and sky-blue section bars.
  No photo, no rating bars, no ATS-breaking graphics.
- **`Rahul_Chaudhary_Resume_preview.pdf`** — a rendered PDF preview so you can see the exact layout
  without opening Word.
- **`index.html`** — a pixel-matched HTML twin of the resume (same fonts, sizes, and spacing). Used to
  verify the one-page fit and handy for exporting to PDF from a browser.
- **`build_resume.py`** — the Python (`python-docx`) generator for the `.docx`. Edit content here and
  re-run to regenerate.
- **`ccet_logo.svg` / `ccet_logo.png`** — the official CCET logo (source SVG + embeddable PNG).

## Regenerate the Word document

```bash
pip install python-docx
python3 build_resume.py     # -> Rahul_Chaudhary_Resume.docx
```

## One-page fit

The layout is tuned to a single US Letter page. The HTML twin was rendered with WeasyPrint and measured
at ~9.78 in of content vs 10.2 in usable height — about **0.42 in of whitespace to spare** — and the
`.docx` mirrors the same font sizes and fixed 10.3 pt line spacing. Please confirm pagination in your own
copy of Word, as rendering can vary slightly by version and default font substitution.

## Export the HTML version to PDF

1. Open `index.html` in Chrome or Edge (keep `ccet_logo.png` in the same folder).
2. Press `Ctrl/Cmd + P`, set **Destination** to *Save as PDF*.
3. Set **Margins** to *None* and enable **Background graphics** so the sky-blue accents render.

## ATS notes

Standard section headings and real, selectable text (not images) with natural keyword coverage for
Generative AI, AI Engineering, Agentic AI, RAG, full stack (React / Node.js / Express), backend, REST
APIs, authentication, MongoDB / PostgreSQL / Supabase, workflow automation (n8n), computer vision, YOLO,
Raspberry Pi, and Docker. No keyword stuffing, no fabricated metrics.
