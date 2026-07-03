# Resume — Rahul Chaudhary

An ATS-friendly, two-page resume for **Rahul Chaudhary** — Generative AI & Agentic AI Engineer | Full Stack Developer.

## File

- **`index.html`** — a self-contained resume (HTML + inline CSS, no dependencies). Features:
  - Modern professional design with a **sky-blue** accent (`#0ea5e9`)
  - **Photo placeholder** in the top-right corner (replace the `PHOTO PLACEHOLDER` box)
  - Two-column layout while keeping selectable, ATS-parseable text
  - Print styles tuned for **A4 PDF export**

## View it

Open `index.html` in any web browser.

## Export to PDF

1. Open `index.html` in Chrome or Edge.
2. Press `Ctrl/Cmd + P` (Print).
3. Set **Destination** to *Save as PDF*.
4. Set **Margins** to *None* and enable **Background graphics** (so the sky-blue accents render).
5. Save.

## Add a profile photo

In `index.html`, replace the `<div class="photo">PHOTO<br />PLACEHOLDER</div>` block with:

```html
<img class="photo" src="photo.jpg" alt="Rahul Chaudhary" />
```

(You may remove the dashed-border styling from `.photo` once a real image is added.)

## ATS notes

The content uses standard section headings, real text (not images), and keyword coverage for Generative AI, Agentic AI, LLM applications, RAG, vector databases, full stack, backend, computer vision, and embedded AI roles.
