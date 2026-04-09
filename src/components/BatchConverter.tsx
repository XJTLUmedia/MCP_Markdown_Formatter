import { useState, useRef, useCallback } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import {
  Upload,
  File,
  X,
  Check,
  AlertCircle,
  Download,
  Archive,
  Loader2,
  ChevronDown,
  ChevronRight,
  Plus,
  Trash2
} from 'lucide-react';
import { saveAs } from 'file-saver';
import type { BatchItem, BatchFormat, BatchResult, BatchProgress } from '../utils/batch-converter';
import { BATCH_FORMATS, executeBatch, packageAsZip } from '../utils/batch-converter';

// ── Helpers ──────────────────────────────────────────────────────────

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

const CATEGORIES: { key: BatchFormat['category']; label: string }[] = [
  { key: 'document', label: 'Documents' },
  { key: 'data', label: 'Data' },
  { key: 'platform', label: 'Platform' },
];

const VALID_EXTENSIONS = ['.md', '.markdown', '.txt', '.text'];

// ── Component ────────────────────────────────────────────────────────

interface BatchConverterProps {
  isDarkMode?: boolean;
}

export function BatchConverter(_props: BatchConverterProps) {
  // Phase management
  const [phase, setPhase] = useState<'select' | 'format' | 'processing' | 'results'>('select');

  // File selection
  const [files, setFiles] = useState<BatchItem[]>([]);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [dragOver, setDragOver] = useState(false);

  // Paste input
  const [showPasteInput, setShowPasteInput] = useState(false);
  const [pasteName, setPasteName] = useState('');
  const [pasteContent, setPasteContent] = useState('');

  // Format selection
  const [selectedFormats, setSelectedFormats] = useState<Set<string>>(new Set());

  // Processing
  const [progress, setProgress] = useState<BatchProgress | null>(null);
  const cancelRef = useRef(false);

  // Results
  const [results, setResults] = useState<BatchResult[]>([]);
  const [expandedErrors, setExpandedErrors] = useState<Set<number>>(new Set());
  const [showIndividual, setShowIndividual] = useState(false);

  // ── File handling ────────────────────────────────────────────────

  const filesRef = useRef<BatchItem[]>([]);
  filesRef.current = files;

  const addFiles = useCallback(async (fileList: File[]) => {
    const valid = fileList.filter(f =>
      VALID_EXTENSIONS.some(ext => f.name.toLowerCase().endsWith(ext))
    );
    const newItems: BatchItem[] = [];
    const existingIds = filesRef.current.map(f => f.id);
    for (const file of valid) {
      const content = await file.text();
      const basename = file.name.replace(/\.[^.]+$/, '') || file.name;
      let id = basename;
      let counter = 1;
      while (existingIds.includes(id) || newItems.some(n => n.id === id)) {
        id = `${basename}_${counter}`;
        counter++;
      }
      newItems.push({ id, filename: file.name, content });
    }
    if (newItems.length > 0) {
      setFiles(prev => [...prev, ...newItems]);
    }
  }, []);

  const removeFile = (id: string) => {
    setFiles(prev => prev.filter(f => f.id !== id));
  };

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
    const dropped = Array.from(e.dataTransfer.files);
    addFiles(dropped);
  }, [addFiles]);

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
  }, []);

  const handleBrowse = () => fileInputRef.current?.click();

  const handleFileInput = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      addFiles(Array.from(e.target.files));
      e.target.value = '';
    }
  };

  const handleAddPaste = () => {
    if (!pasteContent.trim() || !pasteName.trim()) return;
    const basename = pasteName.replace(/\.[^.]+$/, '').trim() || 'untitled';
    let id = basename;
    let counter = 1;
    while (files.some(f => f.id === id)) {
      id = `${basename}_${counter}`;
      counter++;
    }
    setFiles(prev => [...prev, { id, filename: `${basename}.md`, content: pasteContent }]);
    setPasteName('');
    setPasteContent('');
    setShowPasteInput(false);
  };

  // ── Format handling ──────────────────────────────────────────────

  const toggleFormat = (formatId: string) => {
    setSelectedFormats(prev => {
      const next = new Set(prev);
      if (next.has(formatId)) next.delete(formatId);
      else next.add(formatId);
      return next;
    });
  };

  const selectAllInCategory = (category: BatchFormat['category']) => {
    const ids = BATCH_FORMATS.filter(f => f.category === category).map(f => f.id);
    setSelectedFormats(prev => {
      const next = new Set(prev);
      ids.forEach(id => next.add(id));
      return next;
    });
  };

  const deselectAllInCategory = (category: BatchFormat['category']) => {
    const ids = BATCH_FORMATS.filter(f => f.category === category).map(f => f.id);
    setSelectedFormats(prev => {
      const next = new Set(prev);
      ids.forEach(id => next.delete(id));
      return next;
    });
  };

  const categoryAllSelected = (category: BatchFormat['category']) => {
    const ids = BATCH_FORMATS.filter(f => f.category === category).map(f => f.id);
    return ids.every(id => selectedFormats.has(id));
  };

  // ── Batch execution ──────────────────────────────────────────────

  const handleStartBatch = async () => {
    setPhase('processing');
    cancelRef.current = false;
    setProgress(null);
    setResults([]);

    const batchResults = await executeBatch(
      files,
      Array.from(selectedFormats),
      (p) => setProgress(p),
      () => cancelRef.current
    );

    setResults(batchResults);
    setPhase('results');
  };

  const handleCancel = () => {
    cancelRef.current = true;
  };

  const handleDownloadZip = async () => {
    const blob = await packageAsZip(results);
    saveAs(blob, 'batch-export.zip');
  };

  const handleDownloadSingle = (result: BatchResult) => {
    if (result.blob) {
      saveAs(result.blob, result.filename);
    }
  };

  const handleReset = () => {
    setFiles([]);
    setSelectedFormats(new Set());
    setProgress(null);
    setResults([]);
    setExpandedErrors(new Set());
    setShowIndividual(false);
    cancelRef.current = false;
    setPhase('select');
  };

  // ── Computed values ──────────────────────────────────────────────

  const totalSize = files.reduce((sum, f) => sum + new Blob([f.content]).size, 0);
  const successCount = results.filter(r => r.success).length;
  const failCount = results.filter(r => !r.success).length;
  const totalConversions = files.length * selectedFormats.size;
  const pct = progress ? Math.round((progress.completed / progress.total) * 100) : 0;

  // ── Render ───────────────────────────────────────────────────────

  return (
    <div className="card-premium">
      {/* Phase indicator */}
      <div className="flex items-center gap-2 mb-6 text-xs font-semibold tracking-widest uppercase"
        style={{ color: 'var(--text-muted)' }}>
        {(['select', 'format', 'processing', 'results'] as const).map((p, i) => (
          <span key={p} className="flex items-center gap-2">
            {i > 0 && <span style={{ color: 'var(--text-muted)', opacity: 0.3 }}>→</span>}
            <span style={{
              color: phase === p ? 'var(--accent)' : 'var(--text-muted)',
              opacity: phase === p ? 1 : 0.5,
            }}>
              {['1. Files', '2. Formats', '3. Converting', '4. Results'][i]}
            </span>
          </span>
        ))}
      </div>

      <AnimatePresence mode="wait">
        {/* ── Phase 1: File Selection ─────────────────────────────── */}
        {phase === 'select' && (
          <motion.div key="select" initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }}>
            {/* Drop zone */}
            <div
              onDrop={handleDrop}
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              className="relative rounded-xl border-2 border-dashed p-10 text-center transition-all cursor-pointer"
              style={{
                borderColor: dragOver ? 'var(--accent)' : 'var(--border)',
                background: dragOver ? 'rgba(99,102,241,0.05)' : 'transparent',
              }}
              onClick={handleBrowse}
            >
              <Upload className="w-10 h-10 mx-auto mb-3" style={{ color: 'var(--text-muted)' }} />
              <p className="text-base font-medium" style={{ color: 'var(--text-primary)' }}>
                Drop markdown files here
              </p>
              <p className="text-sm mt-1" style={{ color: 'var(--text-muted)' }}>
                or click to browse &middot; .md, .markdown, .txt, .text
              </p>
              <input
                ref={fileInputRef}
                type="file"
                multiple
                accept=".md,.markdown,.txt,.text"
                className="hidden"
                onChange={handleFileInput}
              />
            </div>

            {/* Paste option */}
            <div className="mt-4">
              {!showPasteInput ? (
                <button
                  className="btn-secondary flex items-center gap-2 text-sm"
                  onClick={(e) => { e.stopPropagation(); setShowPasteInput(true); }}
                >
                  <Plus className="w-4 h-4" /> Add from text
                </button>
              ) : (
                <div className="rounded-lg p-4" style={{ background: 'var(--surface-secondary)' }}>
                  <input
                    type="text"
                    placeholder="Document name (e.g. notes)"
                    value={pasteName}
                    onChange={e => setPasteName(e.target.value)}
                    className="w-full mb-2 px-3 py-2 rounded-md text-sm"
                    style={{ background: 'var(--surface-primary)', color: 'var(--text-primary)', border: '1px solid var(--border)' }}
                  />
                  <textarea
                    placeholder="Paste markdown content..."
                    value={pasteContent}
                    onChange={e => setPasteContent(e.target.value)}
                    rows={4}
                    className="w-full mb-2 px-3 py-2 rounded-md text-sm resize-none"
                    style={{ background: 'var(--surface-primary)', color: 'var(--text-primary)', border: '1px solid var(--border)', fontFamily: 'var(--font-mono)' }}
                  />
                  <div className="flex gap-2">
                    <button className="btn-secondary text-sm" onClick={handleAddPaste} disabled={!pasteName.trim() || !pasteContent.trim()}>Add</button>
                    <button className="btn-secondary text-sm" onClick={() => { setShowPasteInput(false); setPasteName(''); setPasteContent(''); }}>Cancel</button>
                  </div>
                </div>
              )}
            </div>

            {/* File list */}
            {files.length > 0 && (
              <div className="mt-6">
                <div className="flex items-center justify-between mb-3">
                  <span className="text-sm font-semibold" style={{ color: 'var(--text-primary)' }}>
                    {files.length} file{files.length !== 1 ? 's' : ''} &middot; {formatSize(totalSize)}
                  </span>
                  <button
                    className="btn-secondary text-xs flex items-center gap-1"
                    onClick={() => setFiles([])}
                  >
                    <Trash2 className="w-3 h-3" /> Clear All
                  </button>
                </div>
                <AnimatePresence>
                  {files.map(file => (
                    <motion.div
                      key={file.id}
                      initial={{ opacity: 0, height: 0 }}
                      animate={{ opacity: 1, height: 'auto' }}
                      exit={{ opacity: 0, height: 0 }}
                      className="flex items-center justify-between px-3 py-2 rounded-lg mb-1"
                      style={{ background: 'var(--surface-secondary)' }}
                    >
                      <div className="flex items-center gap-2 min-w-0">
                        <File className="w-4 h-4 flex-shrink-0" style={{ color: 'var(--accent)' }} />
                        <span className="text-sm truncate" style={{ color: 'var(--text-primary)' }}>{file.filename}</span>
                        <span className="text-xs flex-shrink-0" style={{ color: 'var(--text-muted)' }}>
                          {formatSize(new Blob([file.content]).size)}
                        </span>
                      </div>
                      <button
                        className="p-1 rounded hover:bg-red-500/20 transition-colors"
                        onClick={() => removeFile(file.id)}
                      >
                        <X className="w-4 h-4" style={{ color: 'var(--text-muted)' }} />
                      </button>
                    </motion.div>
                  ))}
                </AnimatePresence>
              </div>
            )}

            {/* Next button */}
            <div className="mt-6 flex justify-end">
              <button
                disabled={files.length === 0}
                onClick={() => setPhase('format')}
                className="px-6 py-2.5 rounded-lg font-semibold text-sm text-white transition-all disabled:opacity-40 disabled:cursor-not-allowed"
                style={{ background: files.length > 0 ? 'var(--accent)' : 'var(--text-muted)' }}
              >
                Next: Select Formats →
              </button>
            </div>
          </motion.div>
        )}

        {/* ── Phase 2: Format Selection ───────────────────────────── */}
        {phase === 'format' && (
          <motion.div key="format" initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }}>
            {CATEGORIES.map(({ key, label }) => {
              const formats = BATCH_FORMATS.filter(f => f.category === key);
              const allSelected = categoryAllSelected(key);
              return (
                <div key={key} className="mb-6">
                  <div className="flex items-center justify-between mb-2">
                    <h3 className="text-sm font-bold uppercase tracking-wider" style={{ color: 'var(--text-primary)' }}>{label}</h3>
                    <button
                      className="text-xs font-medium hover:underline"
                      style={{ color: 'var(--accent)' }}
                      onClick={() => allSelected ? deselectAllInCategory(key) : selectAllInCategory(key)}
                    >
                      {allSelected ? 'Deselect All' : 'Select All'}
                    </button>
                  </div>
                  <div className="flex flex-wrap gap-2">
                    {formats.map(fmt => {
                      const selected = selectedFormats.has(fmt.id);
                      return (
                        <button
                          key={fmt.id}
                          onClick={() => toggleFormat(fmt.id)}
                          className="px-3 py-1.5 rounded-lg text-sm font-medium transition-all border"
                          style={{
                            borderColor: selected ? 'var(--accent)' : 'var(--border)',
                            background: selected ? 'rgba(99,102,241,0.15)' : 'transparent',
                            color: selected ? 'var(--accent)' : 'var(--text-secondary)',
                          }}
                        >
                          {selected && <Check className="w-3 h-3 inline mr-1" />}
                          {fmt.label}
                        </button>
                      );
                    })}
                  </div>
                </div>
              );
            })}

            <div className="flex items-center justify-between mt-6 pt-4" style={{ borderTop: '1px solid var(--border)' }}>
              <div className="flex items-center gap-4">
                <button className="btn-secondary text-sm" onClick={() => setPhase('select')}>← Back</button>
                <span className="text-sm" style={{ color: 'var(--text-muted)' }}>
                  {selectedFormats.size} format{selectedFormats.size !== 1 ? 's' : ''} selected
                </span>
              </div>
              <button
                disabled={selectedFormats.size === 0}
                onClick={handleStartBatch}
                className="px-6 py-2.5 rounded-lg font-semibold text-sm text-white transition-all disabled:opacity-40 disabled:cursor-not-allowed"
                style={{ background: selectedFormats.size > 0 ? 'var(--accent)' : 'var(--text-muted)' }}
              >
                Start Conversion ({files.length} × {selectedFormats.size} = {totalConversions})
              </button>
            </div>
          </motion.div>
        )}

        {/* ── Phase 3: Processing ─────────────────────────────────── */}
        {phase === 'processing' && (
          <motion.div key="processing" initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }}>
            {/* Progress bar */}
            <div className="mb-4">
              <div className="flex justify-between text-sm mb-1">
                <span style={{ color: 'var(--text-primary)' }}>{progress?.current || 'Starting...'}</span>
                <span style={{ color: 'var(--text-muted)' }}>{pct}%</span>
              </div>
              <div className="w-full h-3 rounded-full overflow-hidden" style={{ background: 'var(--surface-secondary)' }}>
                <motion.div
                  className="h-full rounded-full"
                  style={{ background: 'var(--accent)' }}
                  initial={{ width: 0 }}
                  animate={{ width: `${pct}%` }}
                  transition={{ duration: 0.3 }}
                />
              </div>
              <div className="text-xs mt-1" style={{ color: 'var(--text-muted)' }}>
                {progress?.completed || 0} / {progress?.total || totalConversions} conversions
              </div>
            </div>

            {/* Per-file status */}
            <div className="space-y-1 max-h-64 overflow-y-auto mb-4">
              {files.map(file => {
                const fileResults = progress?.results.filter(r => r.itemId === file.id) || [];
                const done = fileResults.length;
                const failed = fileResults.filter(r => !r.success).length;
                const total = selectedFormats.size;
                const inProgress = done < total && (progress?.current?.includes(file.filename) ?? false);

                return (
                  <div
                    key={file.id}
                    className="flex items-center gap-2 px-3 py-2 rounded-lg text-sm"
                    style={{ background: 'var(--surface-secondary)' }}
                  >
                    {done === total ? (
                      failed > 0 ? (
                        <AlertCircle className="w-4 h-4 flex-shrink-0 text-amber-500" />
                      ) : (
                        <Check className="w-4 h-4 flex-shrink-0 text-emerald-500" />
                      )
                    ) : inProgress ? (
                      <Loader2 className="w-4 h-4 flex-shrink-0 animate-spin" style={{ color: 'var(--accent)' }} />
                    ) : (
                      <div className="w-4 h-4 flex-shrink-0 rounded-full border-2" style={{ borderColor: 'var(--border)' }} />
                    )}
                    <span className="truncate" style={{ color: 'var(--text-primary)' }}>{file.filename}</span>
                    <span className="ml-auto text-xs flex-shrink-0" style={{ color: 'var(--text-muted)' }}>
                      {done}/{total}
                      {failed > 0 && <span className="text-red-400 ml-1">({failed} failed)</span>}
                    </span>
                  </div>
                );
              })}
            </div>

            {/* Cancel button */}
            <div className="flex justify-center">
              <button
                onClick={handleCancel}
                className="px-5 py-2 rounded-lg text-sm font-medium border transition-colors"
                style={{ borderColor: 'var(--border)', color: 'var(--text-secondary)' }}
              >
                Cancel
              </button>
            </div>
          </motion.div>
        )}

        {/* ── Phase 4: Results ────────────────────────────────────── */}
        {phase === 'results' && (
          <motion.div key="results" initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }}>
            {/* Summary */}
            <div className="text-center mb-6">
              <div className="text-2xl font-bold mb-1" style={{ color: 'var(--text-primary)' }}>
                {failCount === 0 ? '✅' : '⚠️'} Batch Complete
              </div>
              <p className="text-sm" style={{ color: 'var(--text-muted)' }}>
                <span className="text-emerald-500 font-semibold">{successCount}</span> successful
                {failCount > 0 && (
                  <>, <span className="text-red-400 font-semibold">{failCount}</span> failed</>
                )}
                {' '}out of {results.length} conversions
                {cancelRef.current && <span className="ml-1">(cancelled early)</span>}
              </p>
            </div>

            {/* Download ZIP */}
            {successCount > 0 && (
              <div className="flex justify-center mb-6">
                <button
                  onClick={handleDownloadZip}
                  className="px-8 py-3 rounded-xl font-semibold text-white flex items-center gap-2 transition-transform hover:scale-105"
                  style={{ background: 'linear-gradient(135deg, var(--accent) 0%, #4338ca 100%)' }}
                >
                  <Archive className="w-5 h-5" />
                  Download ZIP ({successCount} files)
                </button>
              </div>
            )}

            {/* Individual downloads */}
            {successCount > 0 && (
              <div className="mb-4">
                <button
                  className="flex items-center gap-2 text-sm font-medium mb-2"
                  style={{ color: 'var(--accent)' }}
                  onClick={() => setShowIndividual(!showIndividual)}
                >
                  {showIndividual ? <ChevronDown className="w-4 h-4" /> : <ChevronRight className="w-4 h-4" />}
                  Download Individual Files
                </button>
                <AnimatePresence>
                  {showIndividual && (
                    <motion.div
                      initial={{ opacity: 0, height: 0 }}
                      animate={{ opacity: 1, height: 'auto' }}
                      exit={{ opacity: 0, height: 0 }}
                      className="space-y-1 max-h-48 overflow-y-auto"
                    >
                      {results.filter(r => r.success).map((r, i) => (
                        <div
                          key={i}
                          className="flex items-center justify-between px-3 py-1.5 rounded-lg text-sm"
                          style={{ background: 'var(--surface-secondary)' }}
                        >
                          <span className="truncate" style={{ color: 'var(--text-primary)' }}>{r.filename}</span>
                          <button
                            onClick={() => handleDownloadSingle(r)}
                            className="p-1 rounded hover:bg-indigo-500/20 transition-colors flex-shrink-0"
                          >
                            <Download className="w-4 h-4" style={{ color: 'var(--accent)' }} />
                          </button>
                        </div>
                      ))}
                    </motion.div>
                  )}
                </AnimatePresence>
              </div>
            )}

            {/* Error report */}
            {failCount > 0 && (
              <div className="mb-6">
                <h4 className="text-sm font-bold mb-2 text-red-400">Errors ({failCount})</h4>
                <div className="space-y-1 max-h-40 overflow-y-auto">
                  {results.filter(r => !r.success).map((r, i) => (
                    <div key={i} className="rounded-lg px-3 py-2" style={{ background: 'rgba(239,68,68,0.08)' }}>
                      <div
                        className="flex items-center gap-2 text-sm cursor-pointer"
                        onClick={() => {
                          setExpandedErrors(prev => {
                            const next = new Set(prev);
                            if (next.has(i)) next.delete(i); else next.add(i);
                            return next;
                          });
                        }}
                      >
                        <AlertCircle className="w-4 h-4 text-red-400 flex-shrink-0" />
                        <span style={{ color: 'var(--text-primary)' }}>{r.filename}</span>
                        {expandedErrors.has(i) ? <ChevronDown className="w-3 h-3 ml-auto text-red-400" /> : <ChevronRight className="w-3 h-3 ml-auto text-red-400" />}
                      </div>
                      {expandedErrors.has(i) && (
                        <p className="text-xs mt-1 pl-6 text-red-400/80">{r.error}</p>
                      )}
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Start new batch */}
            <div className="flex justify-center">
              <button
                onClick={handleReset}
                className="px-6 py-2.5 rounded-lg font-semibold text-sm transition-colors"
                style={{ background: 'var(--surface-secondary)', color: 'var(--text-primary)' }}
              >
                Start New Batch
              </button>
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}
