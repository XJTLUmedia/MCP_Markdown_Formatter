# AI Answer Copier — MCP-Powered Multi-Format Exporter

### *Turn "AI Noise" into "Classroom Ready" in 5 Seconds.*

[![npm version](https://img.shields.io/npm/v/@xjtlumedia/markdown-mcp-server.svg)](https://www.npmjs.com/package/@xjtlumedia/markdown-mcp-server)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![MCP Registry](https://img.shields.io/badge/MCP-Registry-purple)](https://registry.modelcontextprotocol.io)
[![Glama](https://glama.ai/mcp/servers/XJTLUmedia/AI_answer_copier/badge)](https://glama.ai/mcp/servers/XJTLUmedia/AI_answer_copier)
[![GitHub](https://img.shields.io/github/stars/XJTLUmedia/AI_answer_copier?style=social)](https://github.com/XJTLUmedia/AI_answer_copier)

**A Model Context Protocol (MCP) server that gives your AI assistant the power to convert Markdown into 23+ professional document formats**, analyze documents, repair broken Markdown, and convert across platforms — all without leaving your AI chat.

<img width="1018" height="480" alt="image" src="https://github.com/user-attachments/assets/da33b251-f9fe-4bd5-941f-cb92f42350b9" />

npm: https://www.npmjs.com/package/@xjtlumedia/markdown-mcp-server

---

## Why This Exists

You asked an AI to write 20 questions. It delivered. **But now you're stuck:**

| Pain Point | Manual Workflow | With AI Answer Copier |
|---|---|---|
| **Extracting Q&A** | 10–15 min copying each line | **2 seconds** (auto-detect) |
| **Formatting Math** | 20 min fixing broken symbols | **Instant** (KaTeX support) |
| **LMS Upload** | 15 min manual CSV entry | **1-click** export |
| **Total Prep Time** | **~45–60 minutes** | **< 1 minute** |

**Generating content takes seconds. Formatting it for the real world takes an hour. This MCP server eliminates that gap entirely.**

---

## Live Demo

Online web version (no install needed): [ai-answer-copier.vercel.app](https://ai-answer-copier.vercel.app)

<img width="1327" height="732" alt="live_demo1" src="https://github.com/user-attachments/assets/ed6b9b14-2778-45c7-a27d-c3cc8afd9cd9" />
<img width="1170" height="831" alt="live_demo2" src="https://github.com/user-attachments/assets/98c81708-e1e7-45e4-b27c-44450f51c017" />

Test MCP tool in inspector:

<img width="1436" height="491" alt="MCP server" src="https://github.com/user-attachments/assets/882ba600-6a33-41f7-9fdb-ff56e6fa346d" />

Use actual MCP instance:

<img width="925" height="567" alt="MCP1" src="https://github.com/user-attachments/assets/0edb36da-c048-4062-84a7-5521ad7b0f7f" />
<img width="950" height="608" alt="MCP2" src="https://github.com/user-attachments/assets/0dd0c695-1c5a-41ab-9acd-f071b683e3a2" />

---

## All 33 MCP Tools

### Document Conversion Tools (15)

| # | Tool | Output | Description |
|---|---|---|---|
| 1 | `harmonize_markdown` | `.md` | Standardize and normalize Markdown syntax (ATX headers, `-` list markers, fenced code blocks) without changing meaning |
| 2 | `convert_to_txt` | `.txt` | Strip all formatting to produce plain text |
| 3 | `convert_to_html` | `.html` | Full styled HTML document with GFM tables, KaTeX math, and embedded stylesheet |
| 4 | `generate_html` | HTML string | Generate self-contained HTML with all styles inlined (read-only, no file I/O) |
| 5 | `convert_to_pdf` | `.pdf` | Print-ready PDF via headless Chromium with full KaTeX math and syntax highlighting |
| 6 | `convert_to_docx` | `.docx` | Microsoft Word document with styled headings, lists, and code formatting |
| 7 | `convert_to_latex` | `.tex` | LaTeX source code with `\section`, list environments, and native math pass-through |
| 8 | `convert_to_rtf` | `.rtf` | Rich Text Format for legacy word processors and email clients |
| 9 | `convert_to_csv` | `.csv` | Extract GFM pipe-tables to comma-separated values |
| 10 | `convert_to_json` | `.json` | Structured JSON representation with sections, headings, lists, code blocks, and tables |
| 11 | `convert_to_xml` | `.xml` | Well-formed XML with `<?xml?>` declaration and structured elements |
| 12 | `convert_to_xlsx` | `.xlsx` | Excel spreadsheet — each Markdown table becomes a sheet |
| 13 | `convert_to_image` | `.png` | Full-page PNG screenshot via headless Chromium |
| 14 | `convert_to_md` | `.md` | Export Markdown to file, optionally harmonizing formatting first (`harmonize=true`) |
| 15 | `convert_to_email_html` | `.html` | Email-optimized HTML with all styles inlined, compatible with Outlook, Gmail, Apple Mail |

### Platform Converter Tools (10)

| # | Tool | Target Platform | Description |
|---|---|---|---|
| 16 | `convert_to_slack` | Slack | mrkdwn format — bold as `*`, links as `<url\|text>`, headers as bold text |
| 17 | `convert_to_discord` | Discord | Styled bold/underline headers, preserved code blocks |
| 18 | `convert_to_jira` | JIRA | Wiki markup — `h1.`/`h2.`, `{code}` blocks, `[text\|url]` links |
| 19 | `convert_to_confluence` | Confluence | Wiki markup with `{info}`, `{note}` panels and `{code:language=x}` syntax |
| 20 | `convert_to_asciidoc` | AsciiDoc | `=` headers, `----` code blocks, `url[text]` links, `image::` directives |
| 21 | `convert_to_rst` | reStructuredText | Underlined headers, `.. code-block::` directives, RST reference syntax |
| 22 | `convert_to_mediawiki` | MediaWiki | `==` headers, `'''bold'''`, `<syntaxhighlight>` tags, `{\ wikitable` |
| 23 | `convert_to_bbcode` | BBCode | `[b]`, `[i]`, `[code]`, `[url]`, `[img]` tags for phpBB/vBulletin forums |
| 24 | `convert_to_textile` | Textile | Markup for Redmine, Basecamp, and CMS platforms |
| 25 | `convert_to_orgmode` | Emacs Org Mode | `*` headers, `#+BEGIN_SRC`/`#+END_SRC` code, `[[url][text]]` links |

### Import Tools (1)

| # | Tool | Description |
|---|---|---|
| 26 | `html_to_markdown` | Convert HTML content (full document or fragment) back to Markdown |

### Repair & Lint Tools (2)

| # | Tool | Description |
|---|---|---|
| 27 | `repair_markdown` | Fix broken Markdown from LLM output or copy-paste — unclosed code fences, broken tables, stray emphasis markers, missing heading spaces, broken links |
| 28 | `lint_markdown` | Lint and report issues as JSON array with line number, severity, rule name, message, and fixable flag |

### Document Analysis Tools (5)

| # | Tool | Description |
|---|---|---|
| 29 | `extract_code_blocks` | Extract all code blocks with language, content, and start/end line numbers |
| 30 | `extract_links` | Extract all links and images with text, URL, line number, and type (inline, reference, image, autolink) |
| 31 | `generate_toc` | Generate Markdown Table of Contents from headings with configurable `max_depth` (1–6) |
| 32 | `analyze_document` | Comprehensive statistics — line/word/character/paragraph/sentence counts, heading/code/table/link/image counts, estimated reading time |
| 33 | `extract_structure` | Bird's-eye document architecture — statistics, heading outline, code block summary, link summary |

---

## 6 Built-in Prompts

Pre-configured prompt templates that orchestrate multiple tools:

| Prompt | Description |
|---|---|
| `convert-document` | Convert Markdown to any supported format (PDF, DOCX, HTML, LaTeX, CSV, JSON, XML, XLSX, RTF, PNG, TXT, MD) |
| `extract-tables` | Extract all tables from Markdown and export as CSV or XLSX |
| `format-for-sharing` | Harmonize formatting, then convert to both PDF and HTML for distribution |
| `analyze-and-repair` | Lint → Repair → Re-lint to confirm all issues resolved |
| `convert-for-platform` | Convert Markdown to any platform format (Slack, Discord, JIRA, Confluence, AsciiDoc, RST, MediaWiki, BBCode, Textile, Org Mode) |
| `document-overview` | Full overview — statistics, TOC, code blocks, and links in one report |

---

## 2 Resources

| Resource URI | Description |
|---|---|
| `markdown-formatter://supported-formats` | Complete JSON list of all 23+ output formats with tool names and types |
| `markdown-formatter://conversion-guide` | Guide for choosing the right output format based on your use case |

---

## Quick Start

### Install via npx (Recommended)

No installation needed — just configure your MCP client:

```json
{
  "mcpServers": {
    "ai-answer-copier": {
      "command": "npx",
      "args": ["-y", "@xjtlumedia/markdown-mcp-server"]
    }
  }
}
```

### Install Globally

```bash
npm install -g @xjtlumedia/markdown-mcp-server
```

Then configure:

```json
{
  "mcpServers": {
    "ai-answer-copier": {
      "command": "markdown-mcp-server"
    }
  }
}
```

---

## Configuration by AI Client

### 🔌 Claude Desktop

<img width="805" height="621" alt="image" src="https://github.com/user-attachments/assets/b012a055-b595-446a-9c12-934239f3d63b" />

Edit `%APPDATA%\Claude\claude_desktop_config.json` (Windows) or `~/Library/Application Support/Claude/claude_desktop_config.json` (macOS):

```json
{
  "mcpServers": {
    "ai-answer-copier": {
      "command": "npx",
      "args": ["-y", "@xjtlumedia/markdown-mcp-server"]
    }
  }
}
```

Restart Claude Desktop. You'll see a 🔌 icon — all 33 tools are now available.

### VS Code (GitHub Copilot)

Add to `.vscode/mcp.json` or VS Code settings:

```json
{
  "servers": {
    "ai-answer-copier": {
      "command": "npx",
      "args": ["-y", "@xjtlumedia/markdown-mcp-server"]
    }
  }
}
```

### Cursor / Windsurf / Any MCP Client

Use the same `npx` command pattern above in your client's MCP configuration.

### DIY version (local build)

```json
{
  "mcpServers": {
    "markdown-formatter": {
      "command": "node",
      "args": [
        "YOUR_PATH_TO_PROJECT/mcp-server/dist/mcp-server/src/index.js"
      ]
    }
  }
}
```

*Replace `YOUR_PATH_TO_PROJECT` with the absolute path to this folder.*

### HTTP Mode (Vercel)

A hosted HTTP endpoint is available for browser-based and remote integrations:

```
https://ai-answer-copier.vercel.app/api/mcp
```

Test with MCP Inspector:

```bash
npx @modelcontextprotocol/inspector https://ai-answer-copier.vercel.app/api/mcp
```

---

## Usage Examples

Once connected, just talk to your AI naturally:

> **"Generate 10 physics questions about Newton's Laws and export them as a Kahoot CSV."**
> → AI uses `convert_to_csv`

> **"Convert this markdown into a Word document and save it to my Desktop."**
> → AI calls `convert_to_docx` with `output_path`

> **"Take these lecture notes and produce a print-ready PDF."**
> → AI calls `convert_to_pdf` with Puppeteer rendering, KaTeX math, and syntax highlighting

> **"Turn this into a Moodle-compatible XML quiz bank."**
> → AI calls `convert_to_xml`

> **"Format this for Slack and send it to the team channel."**
> → AI calls `convert_to_slack`

> **"Check this Markdown for issues and fix them."**
> → AI calls `lint_markdown` then `repair_markdown`

> **"Give me an overview of this document — word count, structure, all links."**
> → AI chains `analyze_document` + `generate_toc` + `extract_links`

---

## Features

### 23+ Output Formats

Export to PDF, DOCX, HTML, LaTeX, CSV, JSON, XML, XLSX, RTF, PNG, TXT, Markdown — plus Slack, Discord, JIRA, Confluence, AsciiDoc, RST, MediaWiki, BBCode, Textile, Org Mode, and Email HTML.

### Math & Code as First-Class Citizens

Full support for KaTeX math expressions and syntax-highlighted code blocks. Your `$\sqrt{x^2 + y^2}$` and Python snippets survive every conversion perfectly.

### Smart Binary Handling

Binary formats (PDF, DOCX, XLSX, PNG) don't dump raw base64 into chat. The server returns actionable guidance so the AI knows to save to a file path.

### Document Analysis & Repair

Lint Markdown, auto-repair broken formatting, extract code blocks and links, generate Table of Contents, and get comprehensive document statistics — all without leaving the chat.

### GFM (GitHub Flavored Markdown)

Tables, task lists, strikethrough, autolinks — all parsed correctly via `remark-gfm`.

### Puppeteer-Powered PDF & PNG

PDF and image exports use headless Chromium for pixel-perfect rendering with full CSS styling, KaTeX math, syntax highlighting, and print-optimized layouts.

### Self-Contained

Zero runtime dependencies on external APIs. Everything runs locally. Your data never leaves your computer.

---

## Architecture

```
┌─────────────────────────┐
│  AI Client               │
│  (Claude, Copilot,      │
│   Cursor, Windsurf)     │
└──────────┬──────────────┘
           │ MCP Protocol (stdio / HTTP)
           ▼
┌─────────────────────────┐
│  AI Answer Copier        │
│  MCP Server (33 tools)   │
│                          │
│  ┌── Conversion ────────┐│
│  │ remark/rehype, docx, ││
│  │ puppeteer, xlsx,     ││
│  │ CSV/JSON/XML/RTF/    ││
│  │ LaTeX/TXT parsers    ││
│  └──────────────────────┘│
│  ┌── Platform ──────────┐│
│  │ Slack, Discord, JIRA,││
│  │ Confluence, AsciiDoc,││
│  │ RST, MediaWiki,      ││
│  │ BBCode, Textile, Org ││
│  └──────────────────────┘│
│  ┌── Analysis ──────────┐│
│  │ Code blocks, links,  ││
│  │ TOC, stats, structure││
│  └──────────────────────┘│
│  ┌── Repair ────────────┐│
│  │ Lint, auto-repair    ││
│  └──────────────────────┘│
│  ┌── Import ────────────┐│
│  │ HTML → Markdown      ││
│  └──────────────────────┘│
└─────────────────────────┘
```

---

## What's New in v2.1.0

This release hardens both the local (`npx`) and remote Vercel MCP transports with 8 logic fixes:

| # | Fix | Area |
|---|---|---|
| 1 | **Browser not closed on error** — `try/finally` around `browser.close()` so Chromium never leaks on PDF/PNG failures | Both transports |
| 2 | **`generate_html` missing KaTeX CSS** — restored viewport meta, KaTeX stylesheet (SRI hash), and full inline styles so math renders correctly | Vercel `/api/mcp` |
| 3 | **`convert_to_html` incomplete styles** — added SRI integrity hash, table borders, and blockquote styling to match local output | Vercel `/api/mcp` |
| 4 | **Session map memory leak** — `instances` Map now evicts stale sessions after 30 minutes (`SESSION_TTL_MS`); `McpInstance` tracks `lastUsed` | Vercel `/api/mcp` |
| 5 | **`keepAlive` interval not cleared** on normal SSE stream end — added inner `try/finally` that always calls `clearInterval(keepAlive)` | Vercel `/api/mcp` |
| 6 | **`_initialized` SDK-internal access** — replaced bare `@ts-ignore` with a `typeof` guard so the server fails gracefully after SDK upgrades | Vercel `/api/mcp` |
| 7 | **No input size validation** — added 1 GB hard limit (`MAX_INPUT_BYTES`) to all tool calls to prevent runaway CPU/memory on oversized payloads | Both transports |
| 8 | **DELETE method creates new session** — handler now returns `200` immediately on DELETE, preventing phantom session creation | Vercel `/api/mcp` |

---

## Known Limitations

- **PDF / PNG require local Chromium** — `convert_to_pdf` and `convert_to_image` launch a headless browser. This works fine with `npx` but is unavailable on Vercel's free tier (no bundled Chromium). Set the `PUPPETEER_EXECUTABLE_PATH` environment variable to point to your local Chrome if auto-detection fails.
- **Chromium version on Vercel** — the remote endpoint uses `@sparticuz/chromium-min` with a pinned tar at v131.0.1. If Vercel bumps the Lambda runtime you may need to update the `CHROMIUM_TAR_URL` env var.
- **1 MB input limit** — all tool calls reject inputs larger than 1 MB. Split large documents into sections before converting.
- **Session TTL (remote only)** — sessions on the Vercel endpoint expire after 30 minutes of inactivity. Long-running stateful workflows should issue periodic keep-alive calls.

---

## Development

```bash
# Clone the repository
git clone https://github.com/XJTLUmedia/AI_answer_copier.git

# Web app
npm install
npm run dev

# MCP server
cd mcp-server
npm install
npm run dev     # hot reload
npm run build   # production build
npm run inspector  # test with MCP Inspector
```

---

## Tech Stack

| Component | Technology |
|---|---|
| MCP SDK | `@modelcontextprotocol/sdk` |
| Markdown Parser | `unified` + `remark` + `rehype` |
| Math Rendering | `remark-math` + `rehype-katex` |
| Word Export | `docx` |
| Excel Export | `xlsx` |
| PDF/PNG Export | `puppeteer` |
| Schema Validation | `zod` |
| Runtime | Node.js (ESM) |

---

## Contributing

We welcome contributions! Whether it's a new export format, a bug fix, or documentation improvement — please consider making a pull request.

1. Fork the [repository](https://github.com/XJTLUmedia/AI_answer_copier)
2. Create a feature branch
3. Submit a pull request

---

## License

[MIT](https://opensource.org/licenses/MIT) © [XJTLUmedia](https://github.com/XJTLUmedia)

---

<p align="center">
  <b>Built by educators, for educators.</b><br/>
  <i>Reclaim 5 hours of your week. Let the AI handle the formatting.</i>
</p>
