/**
 * Message bodies in the forms an agent writes and reads cheaply.
 *
 * Writing: an agent writes Markdown (fewer output tokens than HTML, nothing to
 * get wrong) and it is sent as formatted HTML; raw HTML in it is escaped and
 * links go only to http(s) or mailto. Reading: mail HTML is mostly styling,
 * so it comes back as Markdown (structure and links kept, hidden preheaders
 * dropped) or as sanitised HTML, never as the sender's raw markup unless
 * `view: "raw"` asks for Outlook's untouched record.
 */
import { z } from 'zod';
import { McpToolError } from '@chrischall/mcp-utils';
import { escapeHtml, htmlToMarkdown, renderMarkdown, sanitiseHtml } from '@chrischall/mcp-utils/rich';

export type BodyFormat = 'text' | 'markdown' | 'html';

export const writeFormatParam = z
  .enum(['text', 'markdown', 'html'])
  .optional()
  .describe(
    'How the text is written: "text" (default, sent as typed), "markdown" (**bold**, _italic_, lists, [links](https://…), `code`, # headings, > quotes — sent as formatted HTML, and cheaper to write than HTML) or "html".',
  );

export const readFormatParam = z
  .enum(['text', 'markdown', 'html'])
  .optional()
  .describe(
    'How the body comes back: "text" (default, Outlook\'s plain text), "markdown" (keeps lists, tables, emphasis and links with their text; hidden text dropped) or "html" (sanitised: formatting tags only).',
  );

/** `html: true`, the older spelling of `format: "html"`, kept for existing callers. */
export const htmlAliasParam = z.boolean().optional().describe('Same as format "html" (older spelling)');

/** The format asked for, `html: true` being the older spelling of "html". */
export function resolveFormat(format: BodyFormat | undefined, html: boolean | undefined): BodyFormat {
  if (html === true && format !== undefined && format !== 'html') {
    throw new McpToolError(`\`html: true\` and \`format: "${format}"\` disagree.`, {
      hint: 'Pass only `format` ("text", "markdown" or "html").',
    });
  }
  return format ?? (html === true ? 'html' : 'text');
}

/** A body to send, as Outlook stores it. */
export function composeBody(content: string, format: BodyFormat): { ContentType: 'Text' | 'HTML'; Content: string } {
  if (format === 'text') return { ContentType: 'Text', Content: content };
  return { ContentType: 'HTML', Content: format === 'markdown' ? renderMarkdown(content, { headings: true, quotes: true }) : content };
}

/** `text` as HTML that shows it exactly: escaped, line breaks kept. */
export const textAsHtml = (text: string): string => escapeHtml(text).replace(/\r?\n/g, '<br>');

/**
 * A fetched body's content in the format asked for. Only an HTML body is
 * converted: a plain-text mail is already what the sender wrote.
 */
export function readBody(body: { ContentType?: string; Content?: string } | undefined, format: BodyFormat): string | undefined {
  const content = body?.Content;
  if (content === undefined || format === 'text' || body?.ContentType?.toLowerCase() !== 'html') return content;
  return format === 'markdown' ? htmlToMarkdown(content) : sanitiseHtml(content);
}
