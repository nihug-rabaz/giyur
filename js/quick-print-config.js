const QUICK_PRINT_CONFIG = {
  templatePath: "templates/template1.docx",
  outputMode: "browserPrint",
  fields: [
    "שם פרטי",
    "שם משפחה",
    "שם האירוע",
    "תאריך תחילה",
    "תאריך סיום",
  ],
};

window.QUICK_PRINT_CONFIG = QUICK_PRINT_CONFIG;

/*
אם שם העמודה בטבלה שונה משם התגית
תוכל לכתוב שדה כאובייקט במקום מחרוזת:

fields: [
  "שם פרטי",
  { tag: "שם האירוע", column: "שיוך לתהליך" },
]
tag = שם התגית במסמך, column = שם העמודה בטבלה.
*/