class SystemDateService {
  static FIELDS = [
    { internal: "printDay", title: "יום הדפסה" },
    { internal: "printMonth", title: "חודש הדפסה (לועזי)" },
    { internal: "printYear", title: "שנת הדפסה" },
    { internal: "gregorianDay", title: "יום לועזי" },
    { internal: "gregorianMonth", title: "חודש לועזי" },
    { internal: "gregorianYear", title: "שנה לועזית" },
    { internal: "hebrewDay", title: "יום עברי" },
    { internal: "hebrewMonth", title: "חודש עברי" },
    { internal: "hebrewYear", title: "שנה עברית" },
  ];

  static LEGACY_INTERNALS = ["printDate", "gregorianDate", "hebrewDate"];

  static INTERNALS = new Set([
    ...SystemDateService.FIELDS.map((field) => field.internal),
    ...SystemDateService.LEGACY_INTERNALS,
  ]);

  static snapshot(date = new Date()) {
    const gregorian = this._gregorianParts(date);
    const hebrew = this._hebrewParts(date);
    return {
      printDay: gregorian.day,
      printMonth: gregorian.month,
      printYear: gregorian.year,
      gregorianDay: gregorian.day,
      gregorianMonth: gregorian.month,
      gregorianYear: gregorian.year,
      hebrewDay: hebrew.day,
      hebrewMonth: hebrew.month,
      hebrewYear: hebrew.year,
      printDate: this.formatPrintDate(date),
      gregorianDate: this.formatGregorian(date),
      hebrewDate: this.formatHebrew(date),
    };
  }

  static valueOf(internal, date = new Date()) {
    if (!this.INTERNALS.has(internal)) return "";
    return this.snapshot(date)[internal] || "";
  }

  static applyTemplateFieldMap(row, map, date = new Date()) {
    const out = { ...row };
    Object.entries(map || {}).forEach(([tag, field]) => {
      if (field?.source !== "system" || !this.INTERNALS.has(field.internal)) return;
      out[tag] = this.valueOf(field.internal, date);
    });
    return out;
  }

  static formatPrintDate(date) {
    return date.toLocaleString("he-IL");
  }

  static formatGregorian(date) {
    const { day, month, year } = this._gregorianParts(date);
    return `${day}.${this._monthNumber(date)}.${year}`;
  }

  static formatHebrew(date) {
    const { day, month, year } = this._hebrewParts(date);
    return `${day} ${month} ${year}`;
  }

  static _gregorianParts(date) {
    const day = String(date.getDate());
    const month = this._inMonthPrefix(new Intl.DateTimeFormat("he-IL", { month: "long" }).format(date));
    const year = String(date.getFullYear());
    return { day, month, year };
  }

  static _hebrewParts(date) {
    const dayNum = Number(new Intl.DateTimeFormat("he-IL", { calendar: "hebrew", day: "numeric" }).format(date));
    const monthName = new Intl.DateTimeFormat("he-IL", { calendar: "hebrew", month: "long" }).format(date);
    const yearNum = Number(new Intl.DateTimeFormat("he-IL", { calendar: "hebrew", year: "numeric" }).format(date));
    return {
      day: this._toHebrewNumeral(dayNum),
      month: this._inMonthPrefix(monthName),
      year: `ה${this._toHebrewNumeral(yearNum % 1000)}`,
    };
  }

  static _monthNumber(date) {
    return String(date.getMonth() + 1).padStart(2, "0");
  }

  static _inMonthPrefix(monthName) {
    const name = String(monthName || "").trim();
    if (!name) return "";
    return name.startsWith("ב") ? name : `ב${name}`;
  }

  static _toHebrewNumeral(number) {
    const n = Number(number);
    if (!Number.isFinite(n) || n < 1) return "";
    if (n === 15) return 'ט"ו';
    if (n === 16) return 'ט"ז';

    const ones = ["", "א", "ב", "ג", "ד", "ה", "ו", "ז", "ח", "ט"];
    const tens = ["", "י", "כ", "ל", "מ", "נ", "ס", "ע", "פ", "צ"];
    const hundreds = ["", "ק", "ר", "ש", "ת"];
    let rest = n;
    let text = "";

    while (rest >= 400) {
      text += "ת";
      rest -= 400;
    }
    const hundred = Math.floor(rest / 100);
    if (hundred) {
      text += hundreds[hundred];
      rest -= hundred * 100;
    }
    const ten = Math.floor(rest / 10);
    if (ten) {
      text += tens[ten];
      rest -= ten * 10;
    }
    if (rest) text += ones[rest];
    if (text.length === 1) return text;
    return `${text.slice(0, -1)}"${text.slice(-1)}`;
  }
}

window.SystemDateService = SystemDateService;
