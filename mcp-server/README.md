# AI Answer Copier — Markdown MCP Server

### Turn AI output into any document format. Instantly.

[![npm version](https://img.shields.io/npm/v/@xjtlumedia/markdown-mcp-server.svg)](https://www.npmjs.com/package/@xjtlumedia/markdown-mcp-server)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![MCP Registry](https://img.shields.io/badge/MCP-Registry-purple)](https://registry.modelcontextprotocol.io)
[![Glama](https://glama.ai/mcp/servers/XJTLUmedia/AI_answer_copier/badge)](https://glama.ai/mcp/servers/XJTLUmedia/AI_answer_copier)
[![GitHub](https://img.shields.io/github/stars/XJTLUmedia/AI_answer_copier?style=social)](https://github.com/XJTLUmedia/AI_answer_copier)

**A Model Context Protocol (MCP) server that gives your AI assistant the power to convert Markdown into 14 professional document formats** — PDF, DOCX, HTML, LaTeX, CSV, JSON, XML, XLSX, RTF, PNG, and more. Stop copy-pasting. Let the AI do the exporting.

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

## 14 Export Tools at Your AI's Fingertips

| Tool | Output Format | Use Case |
|---|---|---|
| `harmonize_markdown` | Clean `.md` | Standardize messy AI output |
| `convert_to_txt` | Plain `.txt` | Strip all formatting |
| `convert_to_html` | `.html` | Web pages, email templates |
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
| `generate_html` | Full HTML doc | Self-contained pages with inline styles |

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

Restart Claude Desktop. You'll see a 🔌 icon — all 14 conversion tools are now available to Claude.

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

### For Content Creators
- **PNG** → Social media posts, slide decks
- **HTML** → Blog posts, newsletters
- **Markdown** → Documentation, READMEs, wikis
- **RTF** → Universal rich text compatibility

---

## Architecture

```
┌─────────────────────┐
│  AI Client           │
│  (Claude, Copilot,  │
│   Cursor, etc.)     │
└────────┬────────────┘
         │ MCP Protocol (stdio)
         ▼
┌─────────────────────┐
│  AI Answer Copier    │
│  MCP Server          │
│                      │
│  ┌─ remark/rehype ──┐│
│  │  Markdown Parser  ││
│  └───────────────────┘│
│  ┌─ docx ───────────┐│
│  │  Word Generator   ││
│  └───────────────────┘│
│  ┌─ puppeteer ──────┐│
│  │  PDF/PNG Renderer ││
│  └───────────────────┘│
│  ┌─ xlsx ───────────┐│
│  │  Excel Generator  ││
│  └───────────────────┘│
│  ┌─ Custom Parsers ─┐│
│  │  CSV/JSON/XML/    ││
│  │  RTF/LaTeX/TXT    ││
│  └───────────────────┘│
└─────────────────────┘
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
