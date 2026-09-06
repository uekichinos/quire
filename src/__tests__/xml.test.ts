import { describe, expect, it } from 'vitest'
import { attr, escapeXml } from '../xml'

describe('escapeXml', () => {
  it('escapes the five XML metacharacters', () => {
    expect(escapeXml(`a & b < c > d " e ' f`)).toBe(
      'a &amp; b &lt; c &gt; d &quot; e &apos; f',
    )
  })

  it('strips characters illegal in XML 1.0', () => {
    expect(escapeXml('a\x00b\x08c\x1Fd')).toBe('abcd')
  })

  it('keeps tab, newline and carriage return', () => {
    expect(escapeXml('a\tb\nc\rd')).toBe('a\tb\nc\rd')
  })

  it('leaves unicode, CJK and emoji intact', () => {
    expect(escapeXml('café 你好 📊')).toBe('café 你好 📊')
  })
})

describe('attr', () => {
  it('renders an escaped key="value" pair with a leading space', () => {
    expect(attr('name', 'A & B')).toBe(' name="A &amp; B"')
  })

  it('accepts numbers', () => {
    expect(attr('count', 3)).toBe(' count="3"')
  })

  it('omits the attribute entirely when the value is undefined', () => {
    expect(attr('name', undefined)).toBe('')
  })
})
