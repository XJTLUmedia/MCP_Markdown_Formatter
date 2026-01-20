import { unified } from 'unified';
import remarkParse from 'remark-parse';
import remarkGfm from 'remark-gfm';
import remarkMath from 'remark-math';
import remarkRehype from 'remark-rehype';
import rehypeKatex from 'rehype-katex';
import rehypeStringify from 'rehype-stringify';
import strip from 'strip-markdown';

import remarkStringify from 'remark-stringify';

/**
 * Converts Markdown to Rich HTML for Word/Google Docs.
 * Includes Math/LaTeX rendering.
 */
export async function markdownToRichHtml(markdown: string): Promise<string> {
  const file = await unified()
    .use(remarkParse)
    .use(remarkGfm)
    .use(remarkMath)
    .use(remarkRehype)
    .use(rehypeKatex)
    .use(rehypeStringify)
    .process(markdown);

  return String(file);
}

/**
 * Harmonizes Markdown (Standardizes syntax, headers, list markers, etc.)
 */
export async function harmonizeMarkdown(markdown: string): Promise<string> {
  const file = await unified()
    .use(remarkParse)
    .use(remarkGfm)
    .use(remarkMath)
    .use(remarkStringify, {
      bullet: '-',
      fence: '`',
      fences: true,
      incrementListMarker: true,
      listItemIndent: 'one',
    })
    .process(markdown);

  return String(file);
}

/**
 * Strips Markdown symbols for clean Plain Text.
 */
export async function markdownToPlainText(markdown: string): Promise<string> {
  const file = await unified()
    .use(remarkParse)
    .use(remarkGfm)
    .use(remarkMath)
    .use(strip)
    .use(remarkStringify)
    .process(markdown);

  return String(file);
}

/**
 * Web Clipper Logic (Bookmarklet Generator)
 * This generates a script that can be dragged to the bookmarks bar.
 * It dynamically loads TurndownService from CDN, converts the page to Markdown,
 * and copies it to the clipboard.
 */
export function generateBookmarklet(): string {
  // Minified bookmarklet that loads TurndownService dynamically
  const code = `(function(){
  var overlay = document.createElement('div');
  overlay.id = 'md-clipper-overlay';
  overlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.7);z-index:999999;display:flex;align-items:center;justify-content:center;';
  overlay.innerHTML = '<div style="background:#1e293b;color:#fff;padding:24px 32px;border-radius:12px;font-family:system-ui,sans-serif;text-align:center;"><div style="font-size:18px;font-weight:600;margin-bottom:8px;">📋 Extracting Content...</div><div style="font-size:14px;opacity:0.7;">Converting to Markdown</div></div>';
  document.body.appendChild(overlay);
  
  function removeOverlay() {
    var el = document.getElementById('md-clipper-overlay');
    if(el) el.remove();
  }
  
  function showResult(success, msg) {
    removeOverlay();
    var result = document.createElement('div');
    result.style.cssText = 'position:fixed;top:20px;right:20px;padding:16px 24px;border-radius:8px;font-family:system-ui,sans-serif;z-index:999999;animation:fadeIn 0.3s;';
    result.style.background = success ? '#10b981' : '#ef4444';
    result.style.color = '#fff';
    result.textContent = msg;
    document.body.appendChild(result);
    setTimeout(function() { result.remove(); }, 3000);
  }
  
  if(typeof TurndownService !== 'undefined') {
    convert();
  } else {
    var script = document.createElement('script');
    script.src = 'https://unpkg.com/turndown@7.2.0/dist/turndown.js';
    script.onload = convert;
    script.onerror = function() {
      showResult(false, '❌ Failed to load converter. Try again.');
    };
    document.head.appendChild(script);
  }
  
  function convert() {
    try {
      var td = new TurndownService({
        headingStyle: 'atx',
        codeBlockStyle: 'fenced',
        bulletListMarker: '-'
      });
      
      td.addRule('removeScripts', {
        filter: ['script', 'style', 'noscript', 'iframe'],
        replacement: function() { return ''; }
      });
      
      var content = document.querySelector('article') || document.querySelector('main') || document.body;
      var markdown = td.turndown(content.innerHTML);
      
      if(navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(markdown).then(function() {
          showResult(true, '✅ Copied! Paste into the Formatter.');
        }).catch(function() {
          fallbackCopy(markdown);
        });
      } else {
        fallbackCopy(markdown);
      }
    } catch(e) {
      showResult(false, '❌ Error: ' + e.message);
    }
  }
  
  function fallbackCopy(text) {
    var el = document.createElement('textarea');
    el.value = text;
    el.style.cssText = 'position:fixed;top:-9999px;';
    document.body.appendChild(el);
    el.select();
    try {
      document.execCommand('copy');
      showResult(true, '✅ Copied! Paste into the Formatter.');
    } catch(e) {
      showResult(false, '❌ Copy failed. Try manually.');
    }
    el.remove();
  }
})();`;

  return `javascript:${encodeURIComponent(code)}`;
}
