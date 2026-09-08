/**
 * A deliberately small, strict XML tokenizer for reading the `.xlsx` parts.
 *
 * It accepts only what machine-generated OOXML uses — the `<?xml?>` declaration,
 * elements, attributes, text, self-closing tags, the five predefined entities
 * and bounded numeric character references — and **hard-errors on anything
 * else**: `<!DOCTYPE>`, `<!ENTITY>`, `<![CDATA[>`, unknown entities, malformed
 * markup. There is no DTD processing, no entity expansion, no namespace
 * resolution — so the whole DOCTYPE / billion-laughs / entity-injection class
 * that dogs general XML parsers simply does not exist here.
 *
 * Hardening: character scanning (no regex on the hot path → no ReDoS),
 * iterative with an explicit tag stack and a depth cap (no stack overflow),
 * numeric references bounded to valid code points, and a guard so
 * `__proto__` / `constructor` / `prototype` can never become object keys.
 */

import { QuireError } from './errors'

export class XmlError extends QuireError {
  constructor(message: string) {
    super(`malformed XML — ${message}`)
    this.name = 'XmlError'
  }
}

export interface XmlHandlers {
  onOpen?(name: string, attrs: Record<string, string>, selfClosing: boolean): void
  onText?(text: string): void
  onClose?(name: string): void
}

const MAX_DEPTH = 256
const FORBIDDEN_KEYS = new Set(['__proto__', 'constructor', 'prototype'])

const NAMED_ENTITIES: Record<string, string> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
}

const ENTITY_RE = /&(#[0-9]+|#x[0-9a-fA-F]+|[A-Za-z][A-Za-z0-9]*);/g

function decodeEntities(input: string): string {
  if (input.indexOf('&') === -1) return input
  return input.replace(ENTITY_RE, (match, body: string) => {
    if (body.charCodeAt(0) === 0x23 /* # */) {
      const hex = body.charCodeAt(1) === 0x78 /* x */
      const cp = parseInt(body.slice(hex ? 2 : 1), hex ? 16 : 10)
      if (!Number.isFinite(cp) || cp < 1 || cp > 0x10ffff || (cp >= 0xd800 && cp <= 0xdfff)) {
        throw new XmlError(`invalid character reference "${match}"`)
      }
      return String.fromCodePoint(cp)
    }
    const named = NAMED_ENTITIES[body]
    if (named === undefined) throw new XmlError(`unknown entity "${match}"`)
    return named
  })
}

function isSpace(code: number): boolean {
  return code === 0x20 || code === 0x09 || code === 0x0a || code === 0x0d
}

/** Finds the `>` that ends a start tag, ignoring any inside quoted attribute values. */
function findTagEnd(src: string, from: number): { contentEnd: number; selfClosing: boolean } {
  let i = from
  let quote = 0
  const n = src.length
  while (i < n) {
    const ch = src.charCodeAt(i)
    if (quote !== 0) {
      if (ch === quote) quote = 0
    } else if (ch === 0x22 || ch === 0x27) {
      quote = ch
    } else if (ch === 0x3e /* > */) {
      const selfClosing = src.charCodeAt(i - 1) === 0x2f /* / */
      return { contentEnd: selfClosing ? i - 1 : i, selfClosing }
    }
    i++
  }
  throw new XmlError('unterminated tag')
}

function parseStartTag(inner: string): { name: string; attrs: Record<string, string> } {
  const attrs: Record<string, string> = Object.create(null)
  const n = inner.length
  let i = 0
  while (i < n && !isSpace(inner.charCodeAt(i))) i++
  const name = inner.slice(0, i)
  if (!name) throw new XmlError('empty tag name')

  while (i < n) {
    while (i < n && isSpace(inner.charCodeAt(i))) i++
    if (i >= n) break

    const nameStart = i
    while (i < n && inner.charCodeAt(i) !== 0x3d /* = */ && !isSpace(inner.charCodeAt(i))) i++
    const attrName = inner.slice(nameStart, i)
    while (i < n && isSpace(inner.charCodeAt(i))) i++
    if (inner.charCodeAt(i) !== 0x3d) throw new XmlError(`malformed attribute "${attrName}"`)
    i++
    while (i < n && isSpace(inner.charCodeAt(i))) i++

    const quote = inner.charCodeAt(i)
    if (quote !== 0x22 && quote !== 0x27) {
      throw new XmlError(`unquoted attribute value for "${attrName}"`)
    }
    i++
    const valStart = i
    while (i < n && inner.charCodeAt(i) !== quote) i++
    if (i >= n) throw new XmlError(`unterminated attribute value for "${attrName}"`)
    const value = decodeEntities(inner.slice(valStart, i))
    i++

    if (attrName && !FORBIDDEN_KEYS.has(attrName)) attrs[attrName] = value
  }

  return { name, attrs }
}

/** Streams `src` through the handlers. Throws `XmlError` on anything malformed or disallowed. */
export function parseXml(src: string, handlers: XmlHandlers): void {
  let i = src.charCodeAt(0) === 0xfeff ? 1 : 0
  const n = src.length
  const stack: string[] = []

  const emitText = (raw: string): void => {
    if (raw.length === 0) return
    // always decode — this is also where entity references are validated
    const decoded = decodeEntities(raw)
    handlers.onText?.(decoded)
  }

  while (i < n) {
    const lt = src.indexOf('<', i)
    if (lt === -1) {
      emitText(src.slice(i))
      break
    }
    if (lt > i) emitText(src.slice(i, lt))

    const c1 = src.charCodeAt(lt + 1)

    if (c1 === 0x3f /* ? */) {
      const end = src.indexOf('?>', lt + 2)
      if (end === -1) throw new XmlError('unterminated processing instruction')
      i = end + 2
      continue
    }

    if (c1 === 0x21 /* ! */) {
      if (src.startsWith('<!--', lt)) {
        const end = src.indexOf('-->', lt + 4)
        if (end === -1) throw new XmlError('unterminated comment')
        i = end + 3
        continue
      }
      throw new XmlError('DOCTYPE, ENTITY and CDATA sections are not allowed')
    }

    if (c1 === 0x2f /* / */) {
      const gt = src.indexOf('>', lt + 2)
      if (gt === -1) throw new XmlError('unterminated end tag')
      const name = src.slice(lt + 2, gt).trim()
      if (!name) throw new XmlError('empty end tag')
      const top = stack.pop()
      if (top !== name) {
        throw new XmlError(`mismatched tag: </${name}> closes <${top ?? 'nothing'}>`)
      }
      handlers.onClose?.(name)
      i = gt + 1
      continue
    }

    const { contentEnd, selfClosing } = findTagEnd(src, lt + 1)
    const { name, attrs } = parseStartTag(src.slice(lt + 1, contentEnd))
    if (stack.length >= MAX_DEPTH) throw new XmlError('maximum nesting depth exceeded')
    handlers.onOpen?.(name, attrs, selfClosing)
    if (selfClosing) {
      handlers.onClose?.(name)
      i = contentEnd + 2
    } else {
      stack.push(name)
      i = contentEnd + 1
    }
  }

  if (stack.length > 0) throw new XmlError(`unclosed tag <${stack[stack.length - 1]}>`)
}
