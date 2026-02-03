import { useState, useEffect } from 'react';
import {
  Trash2,
  Copy,
  Moon,
  Sun,
  ClipboardCheck,
  Zap,
  Activity,
  FileText,
  FileDown,
  FileType,
  FileCode,
  Printer,
  Sigma,
  Table,
  FileJson,
  Image as ImageIcon,
  Sparkles,
  LayoutTemplate
} from 'lucide-react';
import { TEMPLATES } from './data/templates';
import { motion, AnimatePresence } from 'framer-motion';
import {
  markdownToRichHtml,
  markdownToPlainText,
  harmonizeMarkdown,
  generateBookmarklet
} from './utils/formatter';
import {
  exportAsText,
  exportAsMarkdown,
  exportAsHtml,
  exportAsWord,
  exportAsPdf,
  copyRichText,
  exportAsLaTeX,
  exportAsCSV,
  exportAsJSON,
  exportAsXML,
  exportAsRTF,
  exportAsXLSX,
  exportAsImage
} from './utils/exporter';
import 'katex/dist/katex.min.css';

function App() {
  const [input, setInput] = useState('');
  const [output, setOutput] = useState('');
  const [mode, setMode] = useState<'rich' | 'plain'>('rich');
  const [copied, setCopied] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isDarkMode, setIsDarkMode] = useState(true);
  const [isExporting, setIsExporting] = useState(false);
  const [exportMessage, setExportMessage] = useState('');

  // Toggle dark mode on html element
  useEffect(() => {
    const html = document.documentElement;
    if (isDarkMode) {
      html.classList.add('dark');
    } else {
      html.classList.remove('dark');
    }
  }, [isDarkMode]);

  // Process markdown input
  useEffect(() => {
    let isActive = true;
    const timer = setTimeout(async () => {
      if (!input.trim()) {
        if (isActive) setOutput('');
        return;
      }
      setIsProcessing(true);
      try {
        const result = mode === 'rich'
          ? await markdownToRichHtml(input)
          : await markdownToPlainText(input);

        if (isActive) setOutput(result);
      } catch (err) {
        console.error(err);
      } finally {
        if (isActive) setIsProcessing(false);
      }
    }, 400);

    return () => {
      isActive = false;
      clearTimeout(timer);
    };
  }, [input, mode]);

  // Copy to clipboard for pasting into Word/Docs
  const handleCopy = async () => {
    if (!output) return;
    try {
      if (mode === 'rich') {
        await copyRichText(output);
      } else {
        await navigator.clipboard.writeText(output);
      }
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch (err) {
      console.error(err);
    }
  };

  // Show export success message
  const showExportSuccess = (format: string) => {
    setExportMessage(`✅ Exported as ${format}!`);
    setTimeout(() => setExportMessage(''), 3000);
  };

  // Export handlers
  const handleExportWord = async () => {
    if (!input.trim()) return;
    setIsExporting(true);
    try {
      await exportAsWord(input, 'formatted-document');
      showExportSuccess('.docx');
    } catch (err) {
      console.error(err);
      setExportMessage('❌ Export failed');
    }
    setIsExporting(false);
  };

  const handleExportPdf = async () => {
    if (!output) return;
    setIsExporting(true);
    try {
      await exportAsPdf(output, 'formatted-document');
      showExportSuccess('.pdf');
    } catch (err) {
      console.error(err);
      setExportMessage('❌ PDF export failed');
    }
    setIsExporting(false);
  };

  const handleExportHtml = () => {
    if (!output) return;
    exportAsHtml(output, 'formatted-document');
    showExportSuccess('.html');
  };

  const handleExportMarkdown = () => {
    if (!input.trim()) return;
    exportAsMarkdown(input, 'formatted-document');
    showExportSuccess('.md');
  };

  const handleExportText = () => {
    if (!input.trim()) return;
    exportAsText(input, 'formatted-document');
    showExportSuccess('.txt');
  };

  const handleExportLaTeX = () => {
    if (!input.trim()) return;
    exportAsLaTeX(input, 'document');
    showExportSuccess('.tex');
  };

  const handleExportCSV = () => {
    if (!input.trim()) return;
    exportAsCSV(input, 'data');
    showExportSuccess('.csv');
  };

  const handleExportJSON = () => {
    if (!input.trim()) return;
    exportAsJSON(input, 'document');
    showExportSuccess('.json');
  };

  const handleExportXML = () => {
    if (!input.trim()) return;
    exportAsXML(input, 'document');
    showExportSuccess('.xml');
  };

  const handleExportRTF = () => {
    if (!input.trim()) return;
    exportAsRTF(input, 'document');
    showExportSuccess('.rtf');
  };

  const handleExportXLSX = () => {
    if (!input.trim()) return;
    exportAsXLSX(input, 'data');
    showExportSuccess('.xlsx');
  };

  const handleExportImage = async () => {
    const previewEl = document.querySelector('.markdown-preview');
    if (!previewEl) return;
    setIsExporting(true);
    try {
      await exportAsImage('preview-content', 'formatted-preview');
      showExportSuccess('.png');
    } catch (err) {
      console.error(err);
      setExportMessage('❌ Image export failed');
    }
    setIsExporting(false);
  };

  const handleHarmonize = async () => {
    if (!input.trim()) return;
    setIsProcessing(true);
    try {
      const harmonized = await harmonizeMarkdown(input);
      setInput(harmonized);
      setExportMessage('✨ Syntax Harmonized!');
      setTimeout(() => setExportMessage(''), 2000);
    } catch (err) {
      console.error(err);
    }
    setIsProcessing(false);
  };

  const handleCopyBookmarklet = () => {
    const bookmarklet = generateBookmarklet();
    navigator.clipboard.writeText(bookmarklet);
    setExportMessage('📋 Bookmarklet copied! Drag to Bookmarks bar.');
    setTimeout(() => setExportMessage(''), 4000);
  };

  const loadTemplate = (templateName: string) => {
    const template = TEMPLATES.find(t => t.name === templateName);
    if (template) {
      setInput(template.content);
      setExportMessage(`📑 Loaded ${templateName}`);
      setTimeout(() => setExportMessage(''), 3000);
    }
  };

  return (
    <div className="main-container">

      {/* Theme Toggle */}
      <div className="flex justify-end items-center gap-3">
        <button
          onClick={() => setIsDarkMode(!isDarkMode)}
          className="btn-secondary"
          title="Toggle Theme"
          aria-label="Toggle dark mode"
        >
          {isDarkMode ? <Sun className="w-5 h-5" /> : <Moon className="w-5 h-5" />}
        </button>
      </div>

      {/* Header */}
      <header className="business-header">
        <h1>AI ANSWER COPIER</h1>
        <p>Convert Markdown to Word, PDF, HTML, and more. Paste or type your content below.</p>
      </header>

      <div className="flex flex-col gap-10">

        {/* Main Workspace Card */}
        <div className="card-premium">

          {/* Toolbar */}
          <div className="action-toolbar">
            <div className="toolbar-label">
              <Zap className="w-3.5 h-3.5" />
              <span>INPUT</span>
            </div>

            <div className="segmented-control">
              <button
                onClick={() => setMode('rich')}
                className={`segment-item ${mode === 'rich' ? 'segment-item-active' : 'segment-item-inactive'}`}
              >
                RICH HTML
              </button>
              <button
                onClick={() => setMode('plain')}
                className={`segment-item ${mode === 'plain' ? 'segment-item-active' : 'segment-item-inactive'}`}
              >
                PLAIN TEXT
              </button>
            </div>

            <div className="flex gap-2">
              <button
                onClick={handleHarmonize}
                disabled={!input.trim()}
                className="btn-secondary"
                title="Harmonize Grammar & Syntax"
                aria-label="Harmonize"
              >
                <Sparkles className="w-4 h-4" />
              </button>
              <button
                onClick={() => loadTemplate('Technical Audit Report')}
                className="btn-secondary"
                title="Load Technical Audit Template"
                aria-label="Load template"
              >
                <LayoutTemplate className="w-4 h-4" />
              </button>
              <button
                onClick={() => setInput('')}
                className="btn-secondary"
                title="Clear Input"
                aria-label="Clear input"
              >
                <Trash2 className="w-4 h-4" />
              </button>
            </div>
          </div>

          {/* Editor */}
          <textarea
            className="surface-editor"
            placeholder="Paste or type your Markdown here...

Example:
# Heading 1
## Heading 2

**Bold text** and *italic text*

- Bullet point 1
- Bullet point 2

| Column 1 | Column 2 |
|----------|----------|
| Data 1   | Data 2   |

```code block```"
            value={input}
            onChange={(e) => setInput(e.target.value)}
            aria-label="Markdown input"
          />

          {/* Export Buttons */}
          <div className="export-section">
            <div className="export-header">
              <FileDown className="w-4 h-4" />
              <span>EXPORT OPTIONS</span>
              {exportMessage && (
                <span className="export-message">{exportMessage}</span>
              )}
            </div>

            <div className="export-buttons">
              <button onClick={handleExportWord} disabled={!input.trim() || isExporting} className="export-btn export-btn-word" title="Word (.docx)">
                <FileText className="w-4 h-4" />
                <span>Word</span>
              </button>

              <button onClick={handleExportPdf} disabled={!output} className="export-btn export-btn-pdf" title="PDF">
                <Printer className="w-4 h-4" />
                <span>PDF</span>
              </button>

              <button onClick={handleExportLaTeX} disabled={!input.trim()} className="export-btn bg-indigo-600" style={{ background: 'linear-gradient(135deg, #4f46e5 0%, #3730a3 100%)' }} title="LaTeX (.tex)">
                <Sigma className="w-4 h-4" />
                <span>LaTeX</span>
              </button>

              <button onClick={handleExportHtml} disabled={!output} className="export-btn export-btn-html" title="HTML">
                <FileCode className="w-4 h-4" />
                <span>HTML</span>
              </button>

              <button onClick={handleExportMarkdown} disabled={!input.trim()} className="export-btn export-btn-md" title="Markdown (.md)">
                <FileType className="w-4 h-4" />
                <span>MD</span>
              </button>

              <button onClick={handleExportXLSX} disabled={!input.trim()} className="export-btn" style={{ background: 'linear-gradient(135deg, #10b981 0%, #059669 100%)' }} title="Excel (.xlsx)">
                <Table className="w-4 h-4" />
                <span>Excel</span>
              </button>

              <button onClick={handleExportCSV} disabled={!input.trim()} className="export-btn" style={{ background: 'linear-gradient(135deg, #10b981 0%, #047857 100%)' }} title="CSV">
                <Table className="w-4 h-4" />
                <span>CSV</span>
              </button>

              <button onClick={handleExportJSON} disabled={!input.trim()} className="export-btn" style={{ background: 'linear-gradient(135deg, #f59e0b 0%, #d97706 100%)' }} title="JSON">
                <FileJson className="w-4 h-4" />
                <span>JSON</span>
              </button>

              <button onClick={handleExportXML} disabled={!input.trim()} className="export-btn" style={{ background: 'linear-gradient(135deg, #6b7280 0%, #4b5563 100%)' }} title="XML">
                <FileCode className="w-4 h-4" />
                <span>XML</span>
              </button>

              <button onClick={handleExportRTF} disabled={!input.trim()} className="export-btn" style={{ background: 'linear-gradient(135deg, #ec4899 0%, #db2777 100%)' }} title="RTF">
                <FileText className="w-4 h-4" />
                <span>RTF</span>
              </button>

              <button onClick={handleExportText} disabled={!input.trim()} className="export-btn export-btn-txt" title="Text (.txt)">
                <FileText className="w-4 h-4" />
                <span>Text</span>
              </button>

              <button onClick={handleExportImage} disabled={!output} className="export-btn" style={{ background: 'linear-gradient(135deg, #06b6d4 0%, #0891b2 100%)' }} title="PNG Image">
                <ImageIcon className="w-4 h-4" />
                <span>Image</span>
              </button>

              <button onClick={handleCopy} disabled={!output} className="export-btn export-btn-copy" title="Copy Rich Text">
                {copied ? <ClipboardCheck className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
                <span>{copied ? 'Copied!' : 'Copy'}</span>
              </button>
            </div>
          </div>

          {/* Preview */}
          <div className="surface-preview">
            <div className="flex items-center justify-between mb-6">
              <div className="toolbar-label">
                <Activity className="w-3.5 h-3.5" />
                <span>PREVIEW</span>
              </div>
              {isProcessing && (
                <div
                  className="animate-spin"
                  style={{
                    width: '1rem',
                    height: '1rem',
                    border: '2px solid rgba(99, 102, 241, 0.3)',
                    borderTopColor: '#6366f1',
                    borderRadius: '50%'
                  }}
                />
              )}
            </div>

            <AnimatePresence mode="wait">
              <motion.div
                key={mode + (output ? 'full' : 'empty')}
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0 }}
                className="markdown-preview"
              >
                {mode === 'rich' ? (
                  <div id="preview-content" dangerouslySetInnerHTML={{ __html: output || '<p class="opacity-30 italic">Preview will appear here...</p>' }} />
                ) : (
                  <pre
                    id="preview-content"
                    style={{
                      color: 'var(--text-secondary)',
                      fontFamily: 'var(--font-mono)',
                      whiteSpace: 'pre-wrap',
                      wordBreak: 'break-word'
                    }}
                  >
                    {output || 'Plain text preview will appear here...'}
                  </pre>
                )}
              </motion.div>
            </AnimatePresence>
          </div>
        </div>

        {/* Quick Tips */}
        <div className="tips-card">
          <h4>💡 Quick Tips</h4>
          <ul>
            <li><strong>Word/PDF:</strong> High-fidelity document generation with math and complex layout support.</li>
            <li><strong>Data (JSON/CSV/Excel):</strong> Extract structured tables directly into database formats.</li>
            <li><strong>Technical (LaTeX/XML/RTF):</strong> Seamless conversion for scientific and legacy workflows.</li>
            <li><strong>Harmonize:</strong> Click the ✨ icon to automatically standardize your markdown syntax.</li>
            <li><strong>Web Clipper:</strong> <button onClick={handleCopyBookmarklet} className="text-indigo-400 hover:text-indigo-300 underline font-medium cursor-pointer">Copy Bookmarklet</button> to extract markdown from any website directly.</li>
          </ul>
        </div>

      </div>

      {/* Footer */}
      <footer>
        AI answer copier &copy; {new Date().getFullYear()}
      </footer>
    </div>
  );
}

export default App;
