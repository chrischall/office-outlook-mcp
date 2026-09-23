import { z } from 'zod';
import type { OutlookClient } from '../client.js';
import { stripOData } from '../view.js';

/**
 * Paging for every list tool.
 *
 * Outlook caps each page (`$top`) and signals the rest with
 * `@odata.nextLink`. A listing that drops that link is indistinguishable from a
 * complete one, so the model says with confidence that an item does not
 * exist. Every list tool therefore returns `nextLink` when there is more, and
 * accepts it back to fetch the next page.
 */
export const nextLinkParam = z
  .string()
  .min(1)
  .optional()
  .describe(
    'Continue a previous listing: pass its `nextLink` value unchanged to fetch the next page. When set, the paging/filter arguments are taken from the link and ignored here.',
  );

/** Fetch the first page at `path`, or the page a previous `nextLink` points at. */
export function fetchPage<T>(
  client: OutlookClient,
  nextLink: string | undefined,
  path: string,
  opts: { text?: boolean; prefer?: string } = {},
): Promise<T> {
  return nextLink !== undefined ? client.getAbsolute<T>(nextLink, opts) : client.get<T>(path, opts);
}

/**
 * Pass a collection through without a typed projection, trimming the OData
 * envelope from each item but keeping the paging link.
 */
export function plainCollection(data: {
  value?: Record<string, unknown>[];
  '@odata.nextLink'?: string;
}): Record<string, unknown> {
  const items = Array.isArray(data?.value) ? data.value : [];
  const out: Record<string, unknown> = { count: items.length, items: items.map(stripOData) };
  const next = data?.['@odata.nextLink'];
  if (next) out.nextLink = next;
  return out;
}
