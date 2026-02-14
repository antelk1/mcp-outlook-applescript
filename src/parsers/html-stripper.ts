/** Options for controlling HTML-to-plain-text conversion behavior. */
export interface StripHtmlOptions {
    readonly preserveWhitespace?: boolean;
    readonly maxLength?: number;
}

/** Maps common named HTML entities to their Unicode equivalents. */
const HTML_ENTITIES: Record<string, string> = {
    '&nbsp;': ' ',
    '&amp;': '&',
    '&lt;': '<',
    '&gt;': '>',
    '&quot;': '"',
    '&apos;': "'",
    '&#39;': "'",
    '&copy;': '\u00A9',
    '&reg;': '\u00AE',
    '&trade;': '\u2122',
    '&mdash;': '\u2014',
    '&ndash;': '\u2013',
    '&hellip;': '\u2026',
    '&lsquo;': '\u2018',
    '&rsquo;': '\u2019',
    '&ldquo;': '\u201C',
    '&rdquo;': '\u201D',
    '&bull;': '\u2022',
    '&middot;': '\u00B7',
    '&euro;': '\u20AC',
    '&pound;': '\u00A3',
    '&yen;': '\u00A5',
    '&cent;': '\u00A2',
};

/** HTML tags that produce line breaks when rendered (block-level elements). */
const BLOCK_TAGS = new Set([
    'p',
    'div',
    'br',
    'hr',
    'h1',
    'h2',
    'h3',
    'h4',
    'h5',
    'h6',
    'li',
    'tr',
    'blockquote',
    'pre',
    'article',
    'section',
    'header',
    'footer',
    'nav',
    'aside',
    'table',
    'thead',
    'tbody',
    'tfoot',
]);

/** HTML tags whose entire content should be removed (non-visible elements). */
const INVISIBLE_TAGS = new Set(['script', 'style', 'head', 'meta', 'link', 'noscript']);

/**
 * Converts an HTML string to plain text by stripping tags, decoding entities,
 * and normalizing whitespace. Handles block elements, list items, invisible tags,
 * comments, and CDATA sections.
 * @param html - Raw HTML string, or null/undefined.
 * @param options - Controls whitespace preservation and output length truncation.
 * @returns Clean plain text, or an empty string for null/empty input.
 */
export function stripHtml(html: string | null | undefined, options: StripHtmlOptions = {}): string {
    if (html == null || html === '') {
        return '';
    }
    const { preserveWhitespace = false, maxLength = 0 } = options;
    let text = html;
    // Strip content inside invisible tags (script, style, etc.)
    for (const tag of INVISIBLE_TAGS) {
        const regex = new RegExp(`<${tag}[^>]*>[\\s\\S]*?</${tag}>`, 'gi');
        text = text.replace(regex, '');
    }
    // Strip HTML comments
    text = text.replace(/<!--[\s\S]*?-->/g, '');
    // Strip CDATA sections
    text = text.replace(/<!\[CDATA\[[\s\S]*?\]\]>/g, '');
    // Convert list items to bullet-prefixed lines before general block processing
    text = text.replace(/<li[^>]*>/gi, '\n\u2022 ');
    text = text.replace(/<\/li>/gi, '');
    // Insert newlines around block-level tags (skip li, already handled above)
    for (const tag of BLOCK_TAGS) {
        if (tag === 'li')
            continue;
        text = text.replace(new RegExp(`<${tag}[^>]*>`, 'gi'), '\n');
        text = text.replace(new RegExp(`<${tag}[^>]*/>`, 'gi'), '\n');
        text = text.replace(new RegExp(`</${tag}>`, 'gi'), '\n');
    }
    // Strip all remaining HTML tags
    text = text.replace(/<[^>]+>/g, '');
    // Decode decimal numeric entities (e.g., &#169;)
    text = text.replace(/&#(\d+);/g, (_: string, code: string) => {
        const num = parseInt(code, 10);
        return String.fromCharCode(num);
    });
    // Decode hexadecimal numeric entities (e.g., &#xA9;)
    text = text.replace(/&#x([a-fA-F0-9]+);/g, (_: string, code: string) => {
        const num = parseInt(code, 16);
        return String.fromCharCode(num);
    });
    // Decode named HTML entities from the lookup table
    for (const [entity, replacement] of Object.entries(HTML_ENTITIES)) {
        text = text.split(entity).join(replacement);
    }
    if (!preserveWhitespace) {
        // Collapse and normalize whitespace for readable plain text
        text = text
            .replace(/\t/g, ' ')
            .replace(/ +/g, ' ')
            .replace(/\n\s*\n/g, '\n\n')
            .replace(/^[ \t]+/gm, '')
            .replace(/[ \t]+$/gm, '')
            .trim();
    }
    // Truncate to maxLength with an ellipsis suffix if needed
    if (maxLength > 0 && text.length > maxLength) {
        text = text.substring(0, maxLength - 3) + '...';
    }
    return text;
}

/**
 * Checks whether a string contains HTML markup.
 * @param text - Input string, or null/undefined.
 * @returns True if the string contains at least one HTML-like tag.
 */
export function containsHtml(text: string | null | undefined): boolean {
    if (text == null || text === '') {
        return false;
    }
    return /<[a-zA-Z][^>]*>/.test(text);
}

/**
 * Extracts plain text from a body that may be HTML or already plain text.
 * Strips HTML only when markup is detected; otherwise returns the input as-is
 * (with optional length truncation).
 * @param body - Email or note body content, or null/undefined.
 * @param options - Controls whitespace preservation and output length truncation.
 * @returns Plain text output, or an empty string for null/empty input.
 */
export function extractPlainText(body: string | null | undefined, options: StripHtmlOptions = {}): string {
    if (body == null || body === '') {
        return '';
    }
    if (containsHtml(body)) {
        return stripHtml(body, options);
    }
    const { maxLength = 0 } = options;
    if (maxLength > 0 && body.length > maxLength) {
        return body.substring(0, maxLength - 3) + '...';
    }
    return body;
}
