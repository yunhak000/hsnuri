const kstFormatter = new Intl.DateTimeFormat("en-GB", {
  timeZone: "Asia/Seoul",
  year: "numeric",
  month: "2-digit",
  day: "2-digit",
  hour: "2-digit",
  minute: "2-digit",
  second: "2-digit",
  hourCycle: "h23",
});

export const formatKstDateTime = (date = new Date()): string => {
  if (!Number.isFinite(date.getTime())) return "";
  const parts = Object.fromEntries(
    kstFormatter.formatToParts(date).map(({ type, value }) => [type, value]),
  );
  return `${parts.year}-${parts.month}-${parts.day} ${parts.hour}:${parts.minute}:${parts.second}`;
};

// Excel dates have no timezone. ExcelJS stores their displayed components in UTC;
// applying a local timezone here would change the original order time.
export const formatExcelDateTime = (date: Date): string =>
  Number.isFinite(date.getTime())
    ? date.toISOString().slice(0, 19).replace("T", " ")
    : "";

export const normalizeDateTimeText = (value: string): string => {
  const text = value.trim();

  // Only recognize complete date strings, never order numbers or embedded text.
  const plain = text.match(
    /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[ T](\d{1,2})[:.](\d{2})(?:[:.](\d{2})(?:\.\d+)?)?)?$/,
  );
  if (plain) {
    const [, year, month, day, hour, minute, second] = plain;
    const date = `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
    const time = `${(hour ?? "0").padStart(2, "0")}:${minute ?? "00"}:${second ?? "00"}`;
    const parsed = new Date(`${date}T${time}Z`);
    if (formatExcelDateTime(parsed) !== `${date} ${time}`) return value;
    // A timezone-free spreadsheet time is already a Korean wall-clock time.
    return hour === undefined ? date : `${date} ${time}`;
  }

  const isoWithZone = /^\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:?\d{2})$/i;
  const jsDateString = /^(?:Sun|Mon|Tue|Wed|Thu|Fri|Sat) (?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) \d{2} \d{4} \d{2}:\d{2}:\d{2} GMT[+-]\d{4}(?: \([^\r\n]*\))?$/;
  const kstText = text.match(/^(\d{4}-\d{2}-\d{2})[ T](\d{2}:\d{2}:\d{2})\s+(?:KST|\(Korean Standard Time\))$/i);
  if (isoWithZone.test(text) || jsDateString.test(text) || kstText) {
    const parsed = new Date(
      kstText ? `${kstText[1]}T${kstText[2]}+09:00` : text.replace(/^(\d{4}-\d{2}-\d{2}) /, "$1T"),
    );
    return formatKstDateTime(parsed) || value;
  }

  return value;
};
