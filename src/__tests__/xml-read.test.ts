import { describe, expect, it } from 'vitest'
import { parseXml, XmlError, type XmlHandlers } from '../xml-read'

/** Collect a flat event log for easy assertions. */
function events(src: string): string[] {
  const log: string[] = []
  const h: XmlHandlers = {
    onOpen: (name, attrs, sc) =>
      log.push(`open ${name}${sc ? '/' : ''} ${JSON.stringify(attrs)}`),
    onText: (t) => log.push(`text ${JSON.stringify(t)}`),
    onClose: (name) => log.push(`close ${name}`),
  }
  parseXml(src, h)
  return log
}

describe('parseXml — well-formed input', () => {
  it('emits open / text / close in order', () => {
    expect(events('<a><b>hi</b></a>')).toEqual([
      'open a {}',
      'open b {}',
      'text "hi"',
      'close b',
      'close a',
    ])
  })

  it('parses attributes with both quote styles', () => {
    expect(events(`<c r="A1" t='s' s="3"/>`)).toEqual([
      'open c/ {"r":"A1","t":"s","s":"3"}',
      'close c',
    ])
  })

  it('skips the XML declaration and comments', () => {
    expect(events('<?xml version="1.0"?><!-- note --><a/>')).toEqual(['open a/ {}', 'close a'])
  })

  it('strips a leading BOM', () => {
    expect(events('﻿<a/>')).toEqual(['open a/ {}', 'close a'])
  })

  it('decodes the five predefined entities and numeric refs', () => {
    expect(events('<a>&lt;&amp;&gt;&quot;&apos; &#65; &#x41;</a>')).toEqual([
      'open a {}',
      `text "<&>\\"' A A"`,
      'close a',
    ])
  })

  it('preserves significant whitespace in text', () => {
    expect(events('<t xml:space="preserve">  hi  </t>')).toEqual([
      'open t {"xml:space":"preserve"}',
      'text "  hi  "',
      'close t',
    ])
  })

  it('keeps namespaced element and attribute names verbatim', () => {
    expect(events('<x:sheet r:id="rId1"/>')).toEqual([
      'open x:sheet/ {"r:id":"rId1"}',
      'close x:sheet',
    ])
  })

  it('handles a > inside a quoted attribute value', () => {
    expect(events('<f>A1&gt;0</f>')).toEqual(['open f {}', 'text "A1>0"', 'close f'])
    expect(events('<c s="a>b"/>')).toEqual(['open c/ {"s":"a>b"}', 'close c'])
  })
})

describe('parseXml — rejects the dangerous / malformed', () => {
  it('throws on DOCTYPE', () => {
    expect(() => parseXml('<!DOCTYPE x [ <!ENTITY a "b"> ]><x/>', {})).toThrow(XmlError)
    expect(() => parseXml('<!DOCTYPE x><x/>', {})).toThrow(/DOCTYPE/)
  })

  it('throws on a CDATA section', () => {
    expect(() => parseXml('<a><![CDATA[x]]></a>', {})).toThrow(/CDATA/)
  })

  it('throws on an unknown entity', () => {
    expect(() => parseXml('<a>&nbsp;</a>', {})).toThrow(/unknown entity/)
  })

  it('throws on out-of-range and surrogate character references', () => {
    expect(() => parseXml('<a>&#0;</a>', {})).toThrow(/character reference/)
    expect(() => parseXml('<a>&#xD800;</a>', {})).toThrow(/character reference/)
    expect(() => parseXml('<a>&#x110000;</a>', {})).toThrow(/character reference/)
  })

  it('throws on mismatched and unclosed tags', () => {
    expect(() => parseXml('<a></b>', {})).toThrow(/mismatched/)
    expect(() => parseXml('<a><b></a>', {})).toThrow(/mismatched/)
    expect(() => parseXml('<a>', {})).toThrow(/unclosed/)
  })

  it('throws on an unquoted attribute value', () => {
    expect(() => parseXml('<c r=A1/>', {})).toThrow(/unquoted/)
  })

  it('does not let __proto__ / constructor become attribute keys', () => {
    let seen: Record<string, string> = {}
    parseXml('<c __proto__="x" constructor="y" r="A1"/>', {
      onOpen: (_n, attrs) => {
        seen = attrs
      },
    })
    expect(seen).toEqual({ r: 'A1' })
    expect(({} as Record<string, unknown>).x).toBeUndefined()
  })

  it('caps nesting depth (no unbounded stack growth)', () => {
    const deep = '<a>'.repeat(500) + '</a>'.repeat(500)
    expect(() => parseXml(deep, {})).toThrow(/nesting depth/)
  })

  it('handles a large flat document without pathological slowdown', () => {
    const big = '<root>' + '<c r="A1"><v>1</v></c>'.repeat(50_000) + '</root>'
    const t0 = performance.now()
    let cells = 0
    parseXml(big, { onOpen: (n) => n === 'c' && cells++ })
    expect(cells).toBe(50_000)
    expect(performance.now() - t0).toBeLessThan(1000)
  })
})
