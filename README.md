# Therefore Documentation Generator

Generates a Word document describing a Therefore implementation from its XML configuration export — either an uploaded `TheConfiguration-*.xml` or one exported directly from the server. Covers categories (with form renders), workflows (with flowchart diagrams), eForms, indexing profiles, keyword dictionaries, category relationships, queries, reports, and stamps.

## Quick start — Docker (recommended)

```bash
docker compose up --build
```

Open **http://localhost:8000**, enter your server connection details (the default **Export from server** mode) — or switch to **Upload XML export** and drop in a `TheConfiguration-*.xml` — and download the result.

## Quick start — CLI

```bash
pip install -r requirements.txt
npm install -g @mermaid-js/mermaid-cli   # for workflow diagrams

python build_doc.py TheConfiguration-client.xml
# → Therefore_Documentation.docx

python build_doc.py TheConfiguration-client.xml -o ClientName.docx
```

Or export the configuration directly from the server instead of supplying an XML file:

```bash
python build_doc.py --server https://tenant.thereforeonline.com --username jane.smith -o Tenant.docx
# password from $THEREFORE_PASSWORD, otherwise prompted; the exported XML is kept as TheConfiguration-<tenant>.xml
```

### CLI options

| Flag | Description |
|------|-------------|
| `-o / --output` | Output path (default: `Therefore_Documentation.docx`) |
| `--render-dir` | Cache rendered PNGs here instead of a temp dir |
| `--data-dir` | Directory of REST-pulled JSON files for additional enrichment |
| `--no-images` | Skip category / eForm / workflow image renders |
| `--skip-eforms` | Exclude the eForms section |
| `--body-only` | Omit title page — output is content only, for inserting into a wrapper document |
| `--start-section N` | Start section numbering from N (0 = no numbers; default 1) |
| `--theme` | Path to a YAML theme file (see `themes/` for examples) |
| `--format` | Image format for renders: `png` (default) or `svg` |
| `--sections` | Comma-separated list of section keys to include (default: all — see section keys below) |
| `--server` | Export directly from this server instead of reading an XML file |
| `--tenant` | Tenant name (auto-detected for `*.thereforeonline.com`) |
| `--username` / `--password` | Login for `--server` (password defaults to `$THEREFORE_PASSWORD`, else prompts) |
| `--save-xml` | Where to save the exported XML (default: `TheConfiguration-<tenant>.xml`) |

## Exporting directly from the server

With **Export from server** (web) or `--server` (CLI), the generator logs in to the server's `/TheXMLServer` endpoint — the interface Solution Designer uses — and runs the same export as Solution Designer's *Export configuration* with every object selected and role assignments included. Users, groups and memberships come from the export; the Server Configuration section (settings, version, licence, domains) is read over the same connection, so the REST API is not used in this mode. The exported XML can be downloaded from the web UI afterwards.

Limitations:

- This is Therefore's internal Solution Designer protocol, not a published API; verified on build 35.0.3.
- Password login only (ASCII passwords). SSO and MFA accounts are not supported.
- Therefore Online closes any request that runs longer than about five minutes. Very large tenants can exceed this while the server builds the export (Solution Designer hits the same limit). When this happens the generator reports it and stops — upload an XML export instead.
- The login uses a Solution Designer session, so the account needs Solution Designer access.

In upload mode, the optional Server Connection still uses the REST API to enrich the Server Configuration section, as before.

## What's in the generated document

| Section | Key | Content |
|---------|-----|---------|
| AI Summary | `ai_summary` | AI-generated executive summary (requires AI API — see below) |
| Server Configuration | `server_info` | API URL, tenant, region, service version, retention policies |
| Categories | `categories` | Form render image, field table (No, Name, Type, Size, Details) |
| Indexing Profiles | `indexing_profiles` | Target category, filter, init script, field assignment table |
| Workflow Processes | `workflows` | Mermaid flowchart, task table with transitions and resolved field conditions |
| eForms | `eforms` | Form render image, field table with validation, calculated values, logic |
| Users & Groups | `security` | All users, groups, and group membership (requires API credentials) |
| Category Relationships | `relationships` | Cross-reference diagram (cat→cat), keyword dictionary usage, eForm submissions |
| Folder Structure | `folder_structure` | Recursive folder / object hierarchy |
| Keyword Dictionaries | `keyword_dicts` | Values table + cross-reference of every category that uses it |
| Queries | `queries` | Name and ID listing |
| Report Definitions | `reports` | Name, type (Category / Workflow / System / Custom), path |
| Stamps | `stamps` | Name and ID listing |
| Script Inventory | `script_inventory` | All custom scripts from profiles, eForms, and workflows |

The **Key** column is the value to pass to `--sections`. Example: `--sections categories,workflows,eforms`.

All field number references (e.g. `[-370]`) in workflow conditions and eForm scripts are resolved to their field names. Category, eForm, and workflow renders are indicative and may not exactly match the Therefore UI.

## AI summary

The `ai_summary` section calls a local or remote OpenAI-compatible API to generate a concise executive summary of the implementation. Configure via environment variables:

| Variable | Default | Description |
|----------|---------|-------------|
| `AI_API_URL` | `http://localhost:1234/v1` | Base URL of the OpenAI-compatible endpoint |
| `AI_API_KEY` | `lm-studio` | API key (use any string for LM Studio) |
| `AI_MODEL` | _(first available model)_ | Model name to request |

If `AI_API_URL` is not reachable or the section is not in `--sections`, the summary is skipped silently.

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
fetch_server_info.py      Therefore REST API client — server info, users, groups
therefore_xmlserver.py    TheXMLServer client — direct configuration export and server info
ai_summary.py             AI summary generation (OpenAI-compatible)
render_categories.py      PIL-based category form renderer
render_eform.py           PIL-based eForm renderer (Formio JSON → image)
render_workflow.py        Mermaid-based workflow renderer
merge_docs.py             Inserts body.docx into a wrapper at a placeholder
themes.py                 Theme loader
themes/                   Built-in YAML theme files (default, dark, minimal, …)
web/
  app.py                  FastAPI web application
  static/index.html       Single-page frontend
Dockerfile
docker-compose.yml
requirements.txt
puppeteer.json            Chromium config for mmdc in Docker
```
