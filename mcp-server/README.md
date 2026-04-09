# AI Answer Copier — Markdown MCP Server

### Turn AI output into any document format. Instantly.

[![npm version](https://img.shields.io/npm/v/@xjtlumedia/markdown-mcp-server.svg)](https://www.npmjs.com/package/@xjtlumedia/markdown-mcp-server)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![MCP Registry](https://img.shields.io/badge/MCP-Registry-purple)](https://registry.modelcontextprotocol.io)
[![Glama](https://glama.ai/mcp/servers/XJTLUmedia/AI_answer_copier/badge)](https://glama.ai/mcp/servers/XJTLUmedia/AI_answer_copier)
[![GitHub](https://img.shields.io/github/stars/XJTLUmedia/AI_answer_copier?style=social)](https://github.com/XJTLUmedia/AI_answer_copier)

**A Model Context Protocol (MCP) server that gives your AI assistant 34 powerful tools** — convert Markdown to 22 document formats (PDF, DOCX, HTML, LaTeX, CSV, JSON, XML, XLSX, RTF, PNG, Email HTML…), 10 platform-native formats (Slack, Discord, JIRA, Confluence, AsciiDoc, RST, MediaWiki, BBCode, Textile, Org Mode), **batch-convert multiple documents in one call**, plus HTML import, markdown repair & lint, and document analysis. Stop copy-pasting. Let the AI do the exporting.

---

## Why AI Answer Copier?

You asked an AI to generate 20 exam questions. It delivered — beautifully. But then reality hits:

| Pain Point | Manual Workflow | With AI Answer Copier |
|---|---|---|
| **Extracting Q&A** | 10–15 min copying each line | **2 seconds** (auto-detect) |
| **Formatting Math** | 20 min fixing broken symbols | **Instant** (KaTeX support) |
| **LMS Upload** | 15 min manual CSV entry | **1-click** export |
| **Total Prep Time** | **~45–60 minutes** | **< 1 minute** |

**The last mile of AI workflows is broken.** Generating content takes seconds; formatting it for the real world takes an hour. This MCP server eliminates that gap entirely.

---

## 34 MCP Tools at Your AI's Fingertips

### Document Conversion (15 tools)

| Tool | Output Format | Use Case |
|---|---|---|
| `harmonize_markdown` | Clean `.md` | Standardize messy AI output |
| `convert_to_txt` | Plain `.txt` | Strip all formatting |
| `convert_to_html` | `.html` | Web pages, email templates |
| `generate_html` | Full HTML doc | Self-contained pages with inline styles |
| `convert_to_email_html` | Email `.html` | Outlook/Gmail/Apple Mail compatible |
| `convert_to_pdf` | `.pdf` | Print-ready exams, handouts |
| `convert_to_docx` | `.docx` | Microsoft Word documents |
| `convert_to_latex` | `.tex` | Academic papers, journals |
| `convert_to_rtf` | `.rtf` | Rich text for legacy systems |
| `convert_to_csv` | `.csv` | Kahoot, Quizizz, Google Forms |
| `convert_to_json` | `.json` | APIs, Canvas LMS, custom apps |
| `convert_to_xml` | `.xml` | Moodle, Blackboard, SCORM |
| `convert_to_xlsx` | `.xlsx` | Excel spreadsheets |
| `convert_to_image` | `.png` | Social media, presentations |
| `convert_to_md` | `.md` | Documentation, GitHub |

### Platform Converters (10 tools)

| Tool | Target | Use Case |
|---|---|---|
| `convert_to_slack` | Slack | mrkdwn format for Slack messages |
| `convert_to_discord` | Discord | Styled markdown for Discord |
| `convert_to_jira` | JIRA | Wiki markup for Atlassian JIRA |
| `convert_to_confluence` | Confluence | Wiki markup with panels & macros |
| `convert_to_asciidoc` | AsciiDoc | Technical documentation format |
| `convert_to_rst` | reStructuredText | Sphinx / Read the Docs |
| `convert_to_mediawiki` | MediaWiki | Wikipedia-style markup |
| `convert_to_bbcode` | BBCode | phpBB / vBulletin forums |
| `convert_to_textile` | Textile | Redmine, Basecamp |
| `convert_to_orgmode` | Emacs Org Mode | Org-style headers & source blocks |

### Import, Repair & Analysis (8 tools)

| Tool | Description |
|---|---|
| `html_to_markdown` | Convert HTML back to Markdown |
| `repair_markdown` | Auto-fix broken Markdown syntax |
| `lint_markdown` | Check Markdown for style & syntax issues |
| `extract_code_blocks` | Pull out all fenced code blocks |
| `extract_links` | Extract all URLs and link text |
| `generate_toc` | Generate a table of contents |
| `analyze_document` | Word count, reading time, structure stats |
| `extract_structure` | Extract document heading hierarchy || `batch_convert` | Batch-convert multiple documents to 22 formats in one call — accepts items[] × formats[], error-isolated per item, optional output_dir |
Every tool accepts a `markdown` string input and an optional `output_path` to save directly to disk. Binary formats (PDF, DOCX, XLSX, PNG) intelligently guide the AI to save files rather than dumping raw base64.

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

Then configure your MCP client:

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

### Claude Desktop

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

Restart Claude Desktop. You'll see a 🔌 icon — all 34 tools are now available to Claude.

### VS Code (GitHub Copilot)

Add to your `.vscode/mcp.json` or VS Code settings:

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

### HTTP Mode (Vercel)

A hosted HTTP endpoint is also available for browser-based and remote integrations:

```
https://ai-answer-copier.vercel.app/api/mcp
```

Test it with the MCP Inspector:

```bash
npx @modelcontextprotocol/inspector https://ai-answer-copier.vercel.app/api/mcp
```

---

## Usage Examples

Once connected, just talk to your AI naturally:

> **"Batch-convert 5 markdown documents into DOCX, PDF, and HTML simultaneously and save them to my Downloads folder."**

The AI will use `batch_convert` with all five documents and three format IDs.

> **"Generate 10 physics questions about Newton's Laws and export them as a Kahoot CSV."**

The AI will use `convert_to_csv` automatically.

> **"Convert this markdown into a Word document and save it to my Desktop."**

The AI calls `convert_to_docx` with `output_path: "C:/Users/you/Desktop/questions.docx"`.

> **"Take these lecture notes and produce a print-ready PDF."**

The AI calls `convert_to_pdf` with Puppeteer-powered rendering, full KaTeX math support, and syntax highlighting.

> **"Turn this into a Moodle-compatible XML quiz bank."**

The AI calls `convert_to_xml` with structured question/answer extraction.

---

## Features

### Math & Code as First-Class Citizens

Full support for LaTeX math expressions and syntax-highlighted code blocks. Your `$\sqrt{x^2 + y^2}$` and Python snippets survive every conversion perfectly.

### Smart Binary Handling

Binary formats (PDF, DOCX, XLSX, PNG) don't dump raw base64 into chat. Instead, the server returns actionable guidance so the AI knows to save to a file path — a much better UX.

### GFM (GitHub Flavored Markdown)

Tables, task lists, strikethrough, autolinks — all parsed correctly via `remark-gfm`.

### Puppeteer-Powered PDF & PNG

PDF and image exports use headless Chromium for pixel-perfect rendering with:
- Full CSS styling
- KaTeX math rendering
- Syntax-highlighted code blocks
- Print-optimized layouts

### Self-Contained

Zero runtime dependencies on external APIs. Everything runs locally on your machine. Your data never leaves your computer.

---

## Supported Formats Deep Dive

### For Educators
- **CSV/XLSX** → Direct upload to Kahoot, Quizizz, Google Forms
- **PDF** → Print and distribute physical exam papers
- **DOCX** → Edit in Word, share with colleagues
- **XML/JSON** → Import into Canvas, Moodle, Blackboard LMS

### For Developers
- **JSON** → Structured data for APIs and databases
- **XML** → Enterprise integrations, SOAP services
- **HTML** → Embed in web apps, email templates
- **LaTeX** → Academic publishing, research papers
- **AsciiDoc / RST** → Sphinx, Read the Docs, technical manuals
- **Org Mode** → Emacs workflows

### For Content Creators
- **PNG** → Social media posts, slide decks
- **HTML / Email HTML** → Blog posts, newsletters, email campaigns
- **Markdown** → Documentation, READMEs, wikis
- **RTF** → Universal rich text compatibility

### For Teams & Collaboration
- **Slack** → Paste-ready mrkdwn messages
- **Discord** → Styled markdown for servers and channels
- **JIRA / Confluence** → Atlassian wiki markup
- **MediaWiki** → Wikipedia-style documentation
- **BBCode / Textile** → Forum posts (phpBB, Redmine, Basecamp)

### Import & Quality
- **HTML → Markdown** → Round-trip HTML content back to Markdown
- **Repair** → Auto-fix broken Markdown syntax
- **Lint** → Validate style and syntax
- **Analysis** → Word count, reading time, structure extraction

---

## Architecture

```
┌─────────────────────┐
│  AI Client           │
│  (Claude, Copilot,  │
│   Cursor, etc.)     │
└────────┬────────────┘
         │ MCP Protocol (stdio / HTTP)
         ▼
┌──────────────────────────────────┐
│  AI Answer Copier MCP Server     │
│  34 tools                        │
│                                  │
│  ┌─ Document Conversion (15) ───┐│
│  │  remark/rehype · docx · xlsx ││
│  │  puppeteer · KaTeX · CSV/XML ││
│  │  JSON · RTF · LaTeX · TXT    ││
│  │  Email HTML (inline styles)  ││
│  └──────────────────────────────┘│
│  ┌─ Platform Converters (10) ───┐│
│  │  Slack · Discord · JIRA      ││
│  │  Confluence · AsciiDoc · RST ││
│  │  MediaWiki · BBCode          ││
│  │  Textile · Org Mode          ││
│  └──────────────────────────────┘│
│  ┌─ Batch Convert (1) ──────────┐│
│  │  items[] × formats[]         ││
│  │  22 formats · error-isolated ││
│  └──────────────────────────────┘│
│  ┌─ Import & Repair (3) ────────┐│
│  │  HTML→MD · Repair · Lint     ││
│  └──────────────────────────────┘│
│  ┌─ Analysis (5) ───────────────┐│
│  │  Code Blocks · Links · TOC   ││
│  │  Document Stats · Structure  ││
│  └──────────────────────────────┘│
└──────────────────────────────────┘
```

---

## Development

```bash
# Clone the repository
git clone https://github.com/XJTLUmedia/AI_answer_copier.git
cd AI_answer_copier/mcp-server

# Install dependencies
npm install

# Development mode (hot reload)
npm run dev

# Build for production
npm run build

# Test with MCP Inspector
npm run inspector
```

---

## Live Demo

Try the web interface (no install needed):

**[https://ai-answer-copier.vercel.app](https://ai-answer-copier.vercel.app)**

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

We welcome contributions! Whether it's a new export format, a bug fix, or documentation improvement:

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
