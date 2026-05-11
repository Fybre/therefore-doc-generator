# Therefore Documentation Generator

Generates a Word document describing a Therefore implementation from its XML configuration export. Covers categories (with form renders), workflows (with flowchart diagrams), eForms, indexing profiles, keyword dictionaries, category relationships, queries, reports, and stamps.

## Quick start — Docker (recommended)

```bash
docker compose up --build
```

Open **http://localhost:8000**, drop in your `TheConfiguration-*.xml`, and download the result.

## Quick start — CLI

```bash
pip install -r requirements.txt
npm install -g @mermaid-js/mermaid-cli   # for workflow diagrams

python build_doc.py TheConfiguration-client.xml
# → Therefore_Documentation.docx

python build_doc.py TheConfiguration-client.xml -o ClientName.docx
```

### CLI options

| Flag | Description |
|------|-------------|
| `-o / --output` | Output path (default: `Therefore_Documentation.docx`) |
| `--render-dir` | Cache rendered PNGs here instead of a temp dir |
| `--no-images` | Skip category / eForm / workflow image renders |
| `--skip-eforms` | Exclude the eForms section |
| `--body-only` | Omit title page — output is content only, for inserting into a wrapper document |
| `--start-section N` | Start section numbering from N (0 = no numbers; default 1) |

## What's in the generated document

| Section | Content |
|---------|---------|
| Overview | Object counts summary |
| Categories | Form render image, field table (No, Name, Type, Size, Details) |
| Indexing Profiles | Target category, filter, init script, field assignment table |
| Workflows | Mermaid flowchart, task table with transitions and resolved field conditions |
| eForms | Form render image, field table with validation, calculated values, logic |
| Category Relationships | Cross-reference diagram (cat→cat), keyword dictionary usage, eForm submissions |
| Folder Structure | Recursive folder / object hierarchy |
| Keyword Dictionaries | Values table + cross-reference of every category that uses it |
| Queries | Name and ID listing |
| Report Definitions | Name, type (Category / Workflow / System / Custom), path |
| Stamps | Name and ID listing |

All field number references (e.g. `[-370]`) in workflow conditions and eForm scripts are resolved to their field names. Category, eForm, and workflow renders are indicative and may not exactly match the Therefore UI.

## Inserting into a wrapper document

If you have an existing document (title page, table of contents, project overview) and want to insert the Therefore documentation into a specific location:

**1. Add a placeholder** in your wrapper `.docx` — add a paragraph containing exactly:

```
{{THEREFORE_CONTENT}}
```

on its own line, where you want the content inserted.

**2. Generate body-only output:**

```bash
python build_doc.py TheConfiguration.xml -o body.docx --body-only
```

**3. Merge:**

```bash
python merge_docs.py "My Wrapper.docx" body.docx "Final Output.docx"
```

Custom placeholder:

```bash
python merge_docs.py wrapper.docx body.docx output.docx --placeholder "{{CLIENT_CONFIG}}"
```

The web UI handles this automatically — drop the wrapper `.docx` in the second zone.

## Prerequisites

| Requirement | Version | Notes |
|-------------|---------|-------|
| Python | 3.11+ | |
| python-docx | ≥ 1.1 | `pip install -r requirements.txt` |
| Pillow + numpy | any recent | category / eForm renders |
| Node.js | 18+ | for Mermaid CLI |
| `mmdc` | latest | `npm install -g @mermaid-js/mermaid-cli` |
| Chromium | any | required by mmdc — bundled automatically or use system install |

Docker handles all of the above automatically.

## Development

```bash
pip install -r requirements.txt
uvicorn web.app:app --reload --port 8000
```

## Project layout

```
build_doc.py              Main document builder
render_categories.py      PIL-based category form renderer
render_eform.py           PIL-based eForm renderer (Formio JSON → image)
render_workflow.py        Mermaid-based workflow renderer
merge_docs.py             Inserts body.docx into a wrapper at a placeholder
web/
  app.py                  FastAPI web application
  static/index.html       Single-page frontend
Dockerfile
docker-compose.yml
requirements.txt
puppeteer.json            Chromium config for mmdc in Docker
```
