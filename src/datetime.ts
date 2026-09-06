/**
 * Excel 1900 date system. Serial `1` is 1900-01-01; the fractional part is the
 * time of day. Excel incorrectly treats 1900 as a leap year, so every serial
 * from 1900-03-01 onward is shifted up by one — we reproduce that so the numbers
 * match what Excel shows.
 *
 * Adapted from ExcelJS date handling (MIT).
 */

const EPOCH_1900_UTC = Date.UTC(1899, 11, 31)
const DAY_MS = 86_400_000

/** `Date → Excel serial number`. Uses the date's **local** calendar fields. */
export function dateToSerial(date: Date): number {
  const time = date.getTime()
  if (Number.isNaN(time)) {
    throw new Error('@uekichinos/quire: cannot write an invalid Date')
  }

  const dayStartUtc = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate())
  let days = Math.round((dayStartUtc - EPOCH_1900_UTC) / DAY_MS)

  if (days < 1) {
    throw new Error(
      '@uekichinos/quire: dates before 1900-01-01 are not representable in the 1900 date system',
    )
  }

  // Excel's phantom 1900-02-29.
  if (days >= 60) days += 1

  const msIntoDay =
    date.getHours() * 3_600_000 +
    date.getMinutes() * 60_000 +
    date.getSeconds() * 1_000 +
    date.getMilliseconds()

  return days + msIntoDay / DAY_MS
}
