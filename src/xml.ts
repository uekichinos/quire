/** Characters not permitted anywhere in an XML 1.0 document. */
const ILLEGAL_XML_CHARS = /[\x00-\x08\x0B\x0C\x0E-\x1F]/g

const ENTITY: Record<string, string> = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&apos;',
}

/**
 * Escapes **attribute** content for XML (all five metacharacters) and strips
 * characters illegal in XML 1.0, which would otherwise corrupt the whole `.xlsx`.
 */
export function escapeXml(value: string): string {
  return value.replace(ILLEGAL_XML_CHARS, '').replace(/[&<>"']/g, (c) => ENTITY[c]!)
}

/**
 * Escapes **element text** content — only `&`, `<`, `>` need escaping there, so
 * quotes in formulas and strings stay readable (matching what Excel writes).
 */
export function escapeText(value: string): string {
  return value.replace(ILLEGAL_XML_CHARS, '').replace(/[&<>]/g, (c) => ENTITY[c]!)
}

/** `key="escaped value"` — omits the attribute entirely when `value` is undefined. */
export function attr(key: string, value: string | number | undefined): string {
  if (value === undefined) return ''
  return ` ${key}="${escapeXml(String(value))}"`
}

/** Standard XML prolog used on every part. */
export const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
