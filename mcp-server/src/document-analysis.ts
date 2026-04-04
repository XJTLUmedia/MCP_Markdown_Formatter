// ── Document Analysis Functions ──────────────────────────────────────

export interface CodeBlock {
    language: string;
    code: string;
    startLine: number;
    endLine: number;
}

export interface LinkInfo {
    text: string;
    url: string;
    line: number;
    type: 'inline' | 'reference' | 'image' | 'autolink';
}

export interface DocStats {
    lines: number;
    words: number;
    characters: number;
    paragraphs: number;
    sentences: number;
    headings: number;
    codeBlocks: number;
    tables: number;
    links: number;
    images: number;
    lists: number;
    blockquotes: number;
    estimatedReadingTimeMinutes: number;
}

export interface TocEntry {
    level: number;
    text: string;
    slug: string;
    line: number;
}

// ── Code Block Extraction ────────────────────────────────────────────
export function extractCodeBlocks(md: string): CodeBlock[] {
    const blocks: CodeBlock[] = [];
    const lines = md.split('\n');
    let inBlock = false;
    let currentLang = '';
    let currentCode: string[] = [];
    let startLine = 0;

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        const fenceMatch = line.match(/^(`{3,}|~{3,})([\w+-]*)\s*$/);

        if (fenceMatch) {
            if (!inBlock) {
                inBlock = true;
                currentLang = fenceMatch[2] || 'text';
                currentCode = [];
                startLine = i + 1;
            } else {
                blocks.push({
                    language: currentLang,
                    code: currentCode.join('\n'),
                    startLine: startLine + 1, // 1-indexed
                    endLine: i + 1,
                });
                inBlock = false;
            }
        } else if (inBlock) {
            currentCode.push(line);
        }
    }

    // Handle unclosed fence
    if (inBlock && currentCode.length > 0) {
        blocks.push({
            language: currentLang,
            code: currentCode.join('\n'),
            startLine: startLine + 1,
            endLine: lines.length,
        });
    }

    return blocks;
}

// ── Link Extraction ──────────────────────────────────────────────────
export function extractLinks(md: string): LinkInfo[] {
    const links: LinkInfo[] = [];
    const lines = md.split('\n');
    let inCodeBlock = false;

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (/^(`{3,}|~{3,})/.test(line.trim())) {
            inCodeBlock = !inCodeBlock;
            continue;
        }
        if (inCodeBlock) continue;

        // Images: ![alt](url)
        const imgRegex = /!\[([^\]]*)\]\(([^)]+)\)/g;
        let match;
        while ((match = imgRegex.exec(line)) !== null) {
            links.push({ text: match[1], url: match[2], line: i + 1, type: 'image' });
        }

        // Inline links: [text](url) — but not images
        const linkRegex = /(?<!!)\[([^\]]+)\]\(([^)]+)\)/g;
        while ((match = linkRegex.exec(line)) !== null) {
            links.push({ text: match[1], url: match[2], line: i + 1, type: 'inline' });
        }

        // Reference links: [text][ref]
        const refRegex = /\[([^\]]+)\]\[([^\]]*)\]/g;
        while ((match = refRegex.exec(line)) !== null) {
            links.push({ text: match[1], url: match[2] || match[1], line: i + 1, type: 'reference' });
        }

        // Autolinks: <url>
        const autoRegex = /<(https?:\/\/[^>]+)>/g;
        while ((match = autoRegex.exec(line)) !== null) {
            links.push({ text: match[1], url: match[1], line: i + 1, type: 'autolink' });
        }
    }

    return links;
}

// ── Table of Contents Generation ─────────────────────────────────────
export function generateTOC(md: string, maxDepth: number = 6): string {
    const entries = extractTocEntries(md, maxDepth);
    if (entries.length === 0) return '';

    const minLevel = Math.min(...entries.map(e => e.level));
    const lines = entries.map(entry => {
        const indent = '  '.repeat(entry.level - minLevel);
        return `${indent}- [${entry.text}](#${entry.slug})`;
    });

    return '## Table of Contents\n\n' + lines.join('\n') + '\n';
}

export function extractTocEntries(md: string, maxDepth: number = 6): TocEntry[] {
    const entries: TocEntry[] = [];
    const lines = md.split('\n');
    let inCodeBlock = false;
    const slugCounts = new Map<string, number>();

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (/^(`{3,}|~{3,})/.test(line.trim())) {
            inCodeBlock = !inCodeBlock;
            continue;
        }
        if (inCodeBlock) continue;

        const headingMatch = line.match(/^(#{1,6})\s+(.+)$/);
        if (headingMatch) {
            const level = headingMatch[1].length;
            if (level > maxDepth) continue;
            const text = headingMatch[2].replace(/\s*#+\s*$/, '').trim();
            let slug = text.toLowerCase()
                .replace(/[^\w\s-]/g, '')
                .replace(/\s+/g, '-')
                .replace(/-+/g, '-')
                .replace(/^-|-$/g, '');

            // Handle duplicate slugs
            const count = slugCounts.get(slug) || 0;
            slugCounts.set(slug, count + 1);
            if (count > 0) slug = `${slug}-${count}`;

            entries.push({ level, text, slug, line: i + 1 });
        }
    }

    return entries;
}

// ── Document Statistics ──────────────────────────────────────────────
export function analyzeDocument(md: string): DocStats {
    const lines = md.split('\n');
    const plainText = md
        .replace(/```[\s\S]*?```/g, '')
        .replace(/`[^`]+`/g, '')
        .replace(/<[^>]+>/g, '')
        .replace(/[#*_~`>\[\]|()!]/g, ' ');

    const words = plainText.split(/\s+/).filter(w => w.length > 0).length;
    const sentences = plainText.split(/[.!?]+\s/).filter(s => s.trim().length > 0).length;
    const paragraphs = md.split(/\n\s*\n/).filter(p => p.trim().length > 0).length;

    let headings = 0;
    let codeBlocks = 0;
    let tables = 0;
    let images = 0;
    let lists = 0;
    let blockquotes = 0;
    let inCodeBlock = false;
    let inTable = false;

    for (const line of lines) {
        const trimmed = line.trim();
        if (/^(`{3,}|~{3,})/.test(trimmed)) {
            if (!inCodeBlock) codeBlocks++;
            inCodeBlock = !inCodeBlock;
            continue;
        }
        if (inCodeBlock) continue;

        if (/^#{1,6}\s/.test(trimmed)) headings++;
        if (/^[-*+]\s|^\d+\.\s/.test(trimmed)) lists++;
        if (/^>\s/.test(trimmed)) blockquotes++;
        if (/!\[/.test(trimmed)) images += (trimmed.match(/!\[/g) || []).length;
        if (trimmed.startsWith('|') && trimmed.endsWith('|')) {
            if (!inTable) { tables++; inTable = true; }
        } else {
            inTable = false;
        }
    }

    const linkMatches = md.match(/(?<!!)\[([^\]]+)\]\(([^)]+)\)/g);
    const linkCount = linkMatches ? linkMatches.length : 0;

    return {
        lines: lines.length,
        words,
        characters: md.length,
        paragraphs,
        sentences,
        headings,
        codeBlocks,
        tables,
        links: linkCount,
        images,
        lists,
        blockquotes,
        estimatedReadingTimeMinutes: Math.max(1, Math.ceil(words / 200)),
    };
}

// ── Heading Structure Extraction ─────────────────────────────────────
export function extractStructure(md: string): object {
    const entries = extractTocEntries(md);
    const stats = analyzeDocument(md);
    const codeBlocks = extractCodeBlocks(md);
    const links = extractLinks(md);

    return {
        stats,
        outline: entries.map(e => ({
            level: e.level,
            text: e.text,
            line: e.line,
        })),
        codeBlocks: codeBlocks.map(b => ({
            language: b.language,
            lineCount: b.code.split('\n').length,
            startLine: b.startLine,
            endLine: b.endLine,
        })),
        linkSummary: {
            total: links.length,
            byType: {
                inline: links.filter(l => l.type === 'inline').length,
                image: links.filter(l => l.type === 'image').length,
                reference: links.filter(l => l.type === 'reference').length,
                autolink: links.filter(l => l.type === 'autolink').length,
            },
            uniqueUrls: [...new Set(links.map(l => l.url))].length,
        },
    };
}
