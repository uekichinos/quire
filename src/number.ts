/**
 * Renders a JS number for an OOXML `<v>` element. `<v>` must hold a plain
 * decimal string — Excel and strict readers reject exponent notation — so this
 * expands `String(n)`'s `1e+21` / `1e-7` forms into full decimals.
 */
export function numToXml(n: number): string {
  if (Object.is(n, -0)) return '0'
  const s = String(n)
  if (!/[eE]/.test(s)) return s

  const [mantissa, expText] = s.split(/[eE]/) as [string, string]
  const exp = Number(expText)
  const neg = mantissa.startsWith('-')
  const body = neg ? mantissa.slice(1) : mantissa

  const dot = body.indexOf('.')
  const digits = dot === -1 ? body : body.slice(0, dot) + body.slice(dot + 1)
  const pointPos = (dot === -1 ? body.length : dot) + exp

  let out: string
  if (pointPos <= 0) {
    out = `0.${'0'.repeat(-pointPos)}${digits}`
  } else if (pointPos >= digits.length) {
    out = digits + '0'.repeat(pointPos - digits.length)
  } else {
    out = `${digits.slice(0, pointPos)}.${digits.slice(pointPos)}`
  }

  if (out.includes('.')) out = out.replace(/0+$/, '').replace(/\.$/, '')
  return (neg ? '-' : '') + (out || '0')
}
