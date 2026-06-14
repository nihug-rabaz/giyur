# תוסף הדפסות זימונים — מדריך ארכיטקטורה מלא

תוסף Chrome (Manifest V3) בעברית (RTL) המחבר בין נתוני **SharePoint** לבין יצירת/הדפסת מסמכי **Word** מתבניות. כל הלוגיקה רצה בצד הלקוח; אין שרת. מסמך זה מתאר את כל המנגנונים כך שניתן למסור אותו ל-AI אחר והוא יבין כיצד הכל עובד.

---

## 1. תמונת על

המערכת מאפשרת:
1. **הדפסה/ייצוא מהיר** של מסמכי Word מתבנית `.docx` — לכל שורה בטבלה ממלאים תגיות `[שם תג]` בערכים.
2. **שליפת נתונים מ-SharePoint** (VLOOKUP) — העשרת שורות בנתונים מרשימת SharePoint לפי שדה קישור.
3. **מסכי "רשימות משניות"** (זימונים וישיבות בית דין) — הצגת פריטים מרשימת SharePoint משנית המקושרת לרשימת בסיס, סינון לפי תאריך/סוג/מיקום, והדפסה/ייצוא מרוכז.
4. **דף "הדפסות מרוכזות"** (`hub.html`) — מעטפת אחת שמאגדת את מסכי הזימונים והישיבות.
5. **חלון ניהול** (`admin.html`) — קונפיגורציה מלאה (תבניות, מיפוי שדות, רשימות משניות), עם ייצוא/ייבוא JSON.
6. **בונה רשימת עבודה** (`sp-lookup.html`) — הזנת מזהים, שליפת פריטים, עריכה ויצירת פריטים ב-SharePoint.

### עקרון מפתח: גישה ל-SharePoint דרך לשונית מאומתת
דפדפן התוסף לא יכול לבצע בקשות מאומתות ל-SharePoint ישירות (Cookies / CSRF). לכן `SharePointLookupService` מריץ את ה-`fetch` **בתוך לשונית SharePoint פתוחה ומחוברת** באמצעות `chrome.scripting.executeScript`. אם אין לשונית כזו — הוא פותח אחת, מבצע את הפעולה, וסוגר רק לשוניות שהוא עצמו פתח. קוד היצירה/עדכון משתמש ב-`/_api/contextinfo` כדי לקבל `X-RequestDigest`.

---

## 2. מודל הקונפיגורציה (`chrome.storage.local`)

שתי חנויות הגדרות, שתיהן נשמרות מקומית במחשב, וניתנות לייצוא/ייבוא כ-JSON ממסך הניהול.

### `QuickPrintConfigStore` (`js/quick-print-config-store.js`)
הגדרות ההדפסה המהירה הראשית:
```
{ templatePath, outputMode: "browserPrint" | "download", fields: [{ tag, column }] }
```
`fields` ממפה תג בוורד (`[tag]`) לעמודה בטבלה.

### `SharePointConfigStore` (`js/sharepoint-config-store.js`)
ברירות המחדל נגזרות מ-`js/sharepoint-field-map.js` (`SHAREPOINT_LOOKUP`, `SHAREPOINT_FIELD_MAP`).

מבנה גולמי שמור (מה שמסך הניהול עורך) כולל את שדות הבסיס + מפתחות לכל **פרופיל רשימה משנית**. הפרופילים מוגדרים ב-`SharePointConfigStore.PROFILES = ["summons", "sessions"]`, וכל פרופיל משתמש באותם מפתחות עם תחילית (prefix):

| מפתח (לכל prefix) | משמעות |
|---|---|
| `${p}ListTitle` | כותרת הרשימה המשנית |
| `${p}DateFieldInternal` | שדה תאריך לסינון |
| `${p}LookupFieldInternal` | שדה Lookup לרשומת התיק ברשימת הבסיס |
| `${p}DisplayFields` | שדות (סדורים) שיוצגו בטבלת הרשימה המשנית |
| `${p}BaseDisplayFields` | שדות מרשימת הבסיס שיוצגו לצד הפריט |
| `${p}BasePosition` | `"before"` / `"after"` — מיקום עמודות הבסיס |
| `${p}TypeFieldInternal` | עמודת Choice ל"סוג" (יוצר טאבים) — אופציונלי |
| `${p}Types` | `[{ name, templates: [{name, path}] }]` — סוגים עם תבניות (טאבים) |
| `${p}Templates` | `[{ name, path }]` — תבניות שטוחות (כפתורי הדפסה ללא טאבים) |
| `${p}LocationFieldInternal` | עמודת מיקום (Choice או Lookup) לסינון — אופציונלי |

**`get()`** מחזיר אובייקט מעובד עם `siteUrl/listTitle/listUrl/join/fieldMap/displayFields` ועם `summons` ו-`sessions` שכל אחד הוא **פרופיל** מובנה (`_buildProfile`).

מתודות סטטיות חשובות:
- `buildListUrl(siteUrl, listTitle)` → `.../_api/web/lists/getbytitle('...')/items`.
- `normalizeTypes(types)` — ממיר מבנה ישן (`templatePath` יחיד) למבנה `templates: [{name, path}]`.
- `normalizeTemplates(templates)` — מנקה רשימת תבניות שטוחה ל-`[{name, path}]` (מסנן ללא path).

> שני הפרופילים נבדלים רק בקונפיגורציה ובתוויות; הם רצים על אותו קוד (ראו §5).

---

## 3. שכבת ה-SharePoint — `SharePointLookupService` (`js/sharepoint-lookup-service.js`)

הליבה לכל תקשורת SharePoint. נוצר עם `new SharePointLookupService(config, fieldMap)` כאשר `config` מכיל לפחות `listUrl` (ולעיתים `siteUrl`, `join`, `displayFields`).

### גישה ללשונית מאומתת
- `_canFetchViaSharePointTab()` — בודק אם יש `chrome.tabs`/`chrome.scripting`.
- `_withSharePointTab(run)` — מאתר/פותח לשונית SharePoint, מריץ את `run(tabId)`, וסוגר לשונית רק אם נפתחה כאן.
- `_execInTab(tabId, func, args)` — עוטף `chrome.scripting.executeScript`. **חשוב:** ה-`func` מוסרל ורץ בהקשר הדף, ולכן אסור לו להשתמש בסגירות (closures) — כל הנתונים מועברים דרך `args`.

### מתודות שליפה
- `fetchFields()` — שדות הרשימה. מנסה שאילתה "עשירה" (כולל `Required`, `ReadOnlyField`, `EntityPropertyName`) ונופל לשאילתה מינימלית אם נכשל.
- `fetchItems()` — כל הפריטים (`$top=5000`, עם דפדוף `nextLink`).
- `searchItems(propertyName, value, type)` — חיפוש שרת לפי עמודה בודדת (`$filter`), עם וריאנטים מוטטטים (מספרי/טקסטואלי), `AbortController` של 20 שניות.
- `queryItems(filter, top=200)` — מריץ `$filter` שרירותי (למשל טווח תאריכים).
- `fetchItemsByIds(ids, chunkSize=30)` — שליפה אצווה לפי `Id` (מחלק ל-OR-filters כדי לא לחרוג מאורך URL).
- `fetchLookupOptions(fieldTitle)` — `[{id,label}]` של רשימת היעד של שדה Lookup (לדרופדאון/מסנן).
- `fetchFieldChoices(fieldNameOrTitle)` — ערכי שדה Choice.
- `enrich(headers, rows)` — VLOOKUP: מאנדקס פריטים לפי שדה ה-join, ומזריק עמודות ממופות לכל שורה.

### יצירה/עדכון
- `createItem(fields)` → `_createItemViaSharePointTab`:
  1. `POST /_api/contextinfo` לקבלת `X-RequestDigest`.
  2. קריאת `ListItemEntityTypeFullName`.
  3. `loadFieldInfo()` בונה מיפוי `name → { entity (EntityPropertyName), type, title }`. **חובה להשתמש ב-`EntityPropertyName` כמפתחות ה-JSON ולא ב-InternalName** (במיוחד לשדות בעברית).
  4. `buildPayload()` מטפל בטיפוסים: `Lookup` → `<entity>Id` עם מספר (`resolveLookupId` ממיר טקסט ל-Id מול רשימת היעד); `LookupMulti` → אוסף Id-ים; `User` → `<entity>Id`; אחר → ערך ישיר.
  5. `POST` ליצירה. שדות Lookup שלא נפתרו מדווחים ב-`_skipped`.
- `updateItem(itemId, changes)` → `MERGE` עם `IF-MATCH:*` ו-`X-RequestDigest`.

### עיבוד ערכים
- `readValue(item, internal)` — קריאה סלחנית (התאמת מפתח מטושטשת אם השם הפנימי שונה), ושיטוח Lookup/Person/multi לטקסט.
- `static formatValue(value)` — ממיר מחרוזות תאריך ISO ל-`toLocaleDateString("he-IL")` (תצוגה DD.MM.YYYY). מוחל בכל מקום תצוגה ובייצוא.

---

## 4. שירות הרשימה המשנית — `RelatedListService` (`js/related-list-service.js`)

נבנה מעל `SharePointLookupService` ומקבל **פרופיל** ישירות (`new RelatedListService(profile)`).

- `get configured` — האם יש `listTitle` + `listUrl`.
- `loadFields()` — טוען שדות הרשימה (title/internal/entityProperty/type).
- `search({ from, to, type, locations })` — בונה `$filter`:
  - טווח תאריכים: `dateField ge/le datetime'...'`.
  - סוג (Choice): `typeField eq 'value'`.
  - מיקומים: `_locationClause` — ל-Lookup לפי `<entity>Id eq N`, ל-Choice לפי `<entity> eq 'text'`.
- `locationOptions()` — `[{value,label}]`: ל-Lookup דרך `fetchLookupOptions`, ל-Choice דרך `fetchFieldChoices`.
- `locationKeyOf(item)` — מפתח המיקום (Id ל-Lookup, טקסט ל-Choice) לקיבוץ/תצוגה.
- `lookupId(item, internal)` — מפענח את ה-Id של רשומת הבסיס מתוך שדה ה-Lookup.
- `valueOf(item, internal)` — ערך מנורמל (דרך `readValue`).

---

## 5. מסך גנרי לרשימה משנית — `RelatedListPage` (`js/related-list-page.js`)

**מסך אחד שמשרת גם זימונים וגם ישיבות.** נוצר עם `new RelatedListPage({ profileKey, noun })`:
- `summons.js`: `new RelatedListPage({ profileKey: "summons", noun: "זימונים" })`.
- `sessions.js`: `new RelatedListPage({ profileKey: "sessions", noun: "ישיבות" })`.

המסך קורא את `config[profileKey]` כפרופיל, ואת `config.fieldMap`/`config.displayFields` כרשימת הבסיס (משותפת).

זרימת עבודה:
1. **טאבים** (`_renderTabs`) — אם לפרופיל יש `types`, נוצר טאב לכל סוג. אם אין — אין טאבים.
2. **כפתורי הדפסה** (`_renderPrintButtons`) — אם יש טאב פעיל עם תבניות, מציג את תבניות הטאב; אחרת מציג את `profile.templates` (התבניות השטוחות). כך הישיבות מקבלות כפתורי ממ״ח **ללא טאבים**.
3. **מסנן מיקום** (`_loadLocationOptions`) — צ'קבוקסים (כולם מסומנים) מתוך אפשרויות המיקום. ניקוי = ללא סינון.
4. **חיפוש** (`_search`) — דורש לפחות תאריך אחד; קורא ל-`listService.search` עם הטאב הפעיל והמיקומים הנבחרים.
5. **הצגה**:
   - `_showItems` — שורות הרשימה. אם הוגדרו `baseDisplayFields`, מצרף עמודות מרשימת הבסיס (`_mergeColumns` לפי `basePosition`), בעזרת שליפת אצווה של פריטי הבסיס (`_baseMapFor` → `fetchItemsByIds`).
   - `_showPeople` (מצב "תיקי אנשים") — מציג את תיקי הבסיס המקושרים במקום הפריטים.
6. **דוח Word** (`_downloadReport`) — דרך `TableReportService`. אם יש כמה מיקומים, `_reportGroups` מפצל לטבלה לכל מיקום. כותרת הדוח כוללת את טווח התאריכים (`_dateLabel`).
7. **הדפסת תבנית** (`_printTemplate`) — מדפיס את כל השורות המוצגות עם התבנית שנבחרה דרך `QuickPrintService` (מצב `browserPrint`).
8. **מודאל טקסט ארוך** — בתאי טבלה עם טקסט ארוך מ-`RelatedListPage.LONG_TEXT_LENGTH` (=80) התא מקוצר (CSS `td.cell--clip`) ולחיצה פותחת מודאל עם הטקסט המלא (`_openTextModal`). סגירה ב-✕ / לחיצה ברקע / Esc. המודאל נבנה דינמית פעם אחת ומשתמש במחלקות `.modal-overlay/.modal/.modal-header/.modal-body` הקיימות.

---

## 6. דף "הדפסות מרוכזות" — `hub.html` + `hub.js`

מעטפת המאגדת את שני המסכים **בלי לגעת בלוגיקה שלהם**: היא טוענת את `summons.html` ו-`sessions.html` בתוך `<iframe>` (טעינה עצלה — `src` נקבע רק בכניסה הראשונה למסך).

- `top bar` נסתר שנפתח באנימציה: כפתור צף קבוע (המבורגר + שם המסך הפעיל), לחיצה מוסיפה `body.hub-open` שמחליק את הבר (`transform: translateY`).
- בחירת מסך מחליפה את ה-iframe הפעיל (`.hub-frame.is-active`), מעדכנת את שם המסך, וסוגרת את הבר.
- סגירה גם דרך scrim (רקע) או Esc.

> **אסור למחוק את `summons.html`/`sessions.html`** — הם התוכן שה-hub מציג. אין להם יותר כניסה מהפופאפ (רק דרך ה-hub).

---

## 7. יצירת/הדפסת Word

### `DocxTemplateEngine` (`js/docx-template-engine.js`)
ממלא תבניות `.docx` קיימות (שומר על העיצוב המקורי):
- `fill(rowData, headers, mapping)` — פותח את ה-zip של ה-docx, ועל כל חלקי ה-XML (`document.xml`, headers, footers) מבצע:
  - `collapsePlaceholderRuns` — מאחד `[תג]` שפוצל ע״י Word למספר runs.
  - החלפת `[תג]` בערך (escaped) לפי `mapping` ולפי כותרות.
- `extractPlaceholders(buffer)` — מחזיר את כל התגיות בתבנית (לשימוש מסך הניהול).

### `DocxPrintRenderer` (באותו קובץ)
- `printBlobs(blobs)` — אם זמינה `docx.renderAsync` (ספריית `docx-preview`), מרנדר כל docx ל-HTML מעוצב ומדפיס (שומר עיצוב). אחרת — fallback לטקסט פשוט.
- מדפיס דרך `<iframe>` נסתר עם `print()`.

### `QuickPrintService` (`js/quick-print-service.js`)
- `buildMapping()` — לכל תג מאתר את עמודת הטבלה (התאמה מדויקת/חלקית, ואז `PlaceholderMapper.autoMatch`).
- `printAll(rows)` — לפי `outputMode`: `browserPrint` (דרך `DocxPrintRenderer`) או `download` (קובץ יחיד, או ZIP לרבים דרך `JSZip`).

### `TableReportService` (`js/table-report-service.js`)
בונה דוח Word **טבלאי** (לרוחב, RTL) באמצעות ספריית `docx`:
- `download(headers, groups)` / `build(headers, groups)` — `groups` = `[{title, rows}]` (מערך שורות פשוט = טבלה אחת).
- כל טבלה: `visuallyRightToLeft: true`, כותרות מוצללות, יישור מרכז.
- **קריטי:** הספרייה נלקחת מ-`window.docxBuilder` (ראו §10), לא מ-`window.docx` (ש-`docx-preview` דורסת).

### `PlaceholderMapper` (`js/placeholder-mapper.js`)
התאמה אוטומטית של תג לכותרת: נירמול (הסרת רווחים/פיסוק), התאמה מדויקת → חלקית → לפי קבוצות מילים נרדפות (`ALIASES`).

---

## 8. חלון הניהול — `admin.html` + `admin.js`

`AdminConfigPage` מנהל את כל הקונפיגורציה. רכיבים:

- **`OrderedFieldPicker`** — רכיב לבחירת שדות וקביעת סדר (הוספה/הסרה/↑↓). בשימוש לשדות התצוגה של רשימת הבסיס ושל כל פרופיל.
- **`RelatedListEditor`** — **רכיב לשימוש חוזר** שבונה את כל כרטיס ההגדרות של רשימה משנית (כותרת, טען שדות, שדה תאריך, Lookup, בוחרי שדות תצוגה/בסיס, מיקום עמודות הבסיס, מקטע סוגים/תבניות, שדה מיקום). מופעל פעמיים — לזימונים ולישיבות.
  - דגל `useTypeTabs` (מתוך `labels`):
    - `true` (זימונים) → בונה מקטע **"סוגים ותבניות" (טאבים)**: עמודת Choice + כרטיסי סוג, כל סוג עם תבניות.
    - `false` (ישיבות) → בונה רק מקטע **"תבניות הדפסה (כפתורים)"** — תבניות שטוחות בלי טאבים/עמודת סוג.
  - `load(raw)` / `collect()` קוראים/כותבים את מפתחות הפרופיל לפי ה-prefix.
  - `loadFields()` / `loadTypeChoices()` שולפים שדות/אפשרויות דרך ה-host (שמריץ `chrome.scripting` בלשונית SharePoint).

- שליפת שדות במסך הניהול (`_fetchFields`, `_fetchChoices`) רצה בלשונית SharePoint עם fallback ל-`SharePointLookupService`.
- **ייצוא/ייבוא**: `_export` יוצר JSON (`{ type, version, quickPrint, sharePoint }`), `_import` שומר ומרענן. ההגדרות נשמרות מקומית וגם ניתנות לייצוא לקובץ.

---

## 9. הפופאפ ושאר המסכים

### `popup.html` + `popup.js`
תפריט הפעולות הראשי. כפתורים נוכחיים:
- **הדפסות מרוכזות** → `hub.html` (`OpenPageButton`).
- **פתח דף זימונים** (`OpenSummonsPageButton`) — לוחץ אוטומטית על כפתור "יצירה מרובה" בדף SharePoint.
- **חיפוש פריט ברשימה** → `sp-lookup.html`.
- **הורד דוח Word** (`DownloadReportButton`) — שולח `exportWord` ל-content script.
- **צור מזומנים** (`CreateSummonsButton`).
- כניסה למסך הניהול: `AdminUnlock` (10 לחיצות על הלוגו ואז הקלדת המילה הסודית `"שלום"`) **או** `AdminGate` (כפתור הגדרות + סיסמה `rabaz123`).

### `content.js` + `js/table-reader.js`
Content script שרץ בכל הדפים:
- `scrapeTable` — קורא טבלת זימונים מתוך דיאלוג (`SummonsTableReader`) ושומר ב-`storage`.
- `exportWord` — מזריק `export-page.js` להקשר הדף (לייצוא בהקשר המאומת), אחרי שמעביר את ה-config דרך attribute `data-sp-config`.

### `sp-lookup.html` + `sp-lookup.js`
בונה רשימת עבודה: הזנת מזהה בכל שורה → שליפת פריט מ-SharePoint (חיפוש שרת) → מילוי העמודות. תומך בעריכת פריטים, יצירת פריטים (מודאל טופס עם שדות חובה), וייצוא/הדפסה לתבניות.

---

## 10. מלכודות ופתרונות (חשוב ל-AI)

- **התנגשות ספריות `docx` מול `docx-preview`**: `docx-preview.min.js` דורסת את `window.docx`. לכן `js/docx-builder-bridge.js` (`window.docxBuilder = window.docx;`) נטען **מיד אחרי** `index.umd.js` ולפני `docx-preview`, ו-`TableReportService` משתמש ב-`window.docxBuilder`. ה-bridge הוא קובץ חיצוני כי CSP של התוסף חוסם סקריפטים inline.
- **RTL ב-Word**: טבלאות עם `visuallyRightToLeft: true` + `bidirectional` בפסקאות; section עם `bidi: true`.
- **יצירת פריט ב-SharePoint**: יש להשתמש ב-`EntityPropertyName` כמפתחות; שדות Lookup דורשים `<entity>Id` עם **מזהה מספרי**.
- **`executeScript` ללא closures**: כל הערכים עוברים דרך `args`; הפונקציה מוסרלת לדף.
- **תאריכים**: תמיד דרך `SharePointLookupService.formatValue` לתצוגה אחידה.
- **רשימות גדולות**: חיפוש בצד שרת (`$filter`) עם timeout ו-fallback; שליפת אצווה לפי Id. מומלץ שעמודות חיפוש יהיו Indexed וערך-יחיד.
- **ה-hub** תלוי ב-`summons.html`/`sessions.html` (iframes) — אין למחוק אותם.

---

## 11. מפת קבצים

```
manifest.json                      הגדרות התוסף (MV3)
popup.html / popup.js              תפריט פעולות + שערי ניהול
hub.html / hub.js                  "הדפסות מרוכזות" — מעטפת iframes
summons.html / summons.js          מסך זימונים (entry דק ל-RelatedListPage)
sessions.html / sessions.js        מסך ישיבות בית דין (entry דק)
admin.html / admin.js              חלון ניהול + OrderedFieldPicker + RelatedListEditor
sp-lookup.html / sp-lookup.js      בונה רשימת עבודה + עריכה/יצירה
quick-print.html / quick-print.js  דף הדפסה מהירה
content.js                         content script (scrape / export)
export-page.js / single-export.*   ייצוא בהקשר הדף

js/
  sharepoint-config-store.js       חנות הגדרות SP + פרופילים גנריים
  quick-print-config-store.js      חנות הגדרות הדפסה מהירה
  quick-print-config.js            ברירות מחדל הדפסה מהירה
  sharepoint-field-map.js          ברירות מחדל (siteUrl/list/join/fieldMap)
  sharepoint-lookup-service.js     כל תקשורת SharePoint (fetch/create/update)
  related-list-service.js          שירות פרופיל רשימה משנית
  related-list-page.js             מסך גנרי לרשימה משנית (זימונים/ישיבות)
  docx-template-engine.js          מילוי תבנית docx + רינדור/הדפסה + הורדה
  quick-print-service.js           הדפסה/ייצוא מתבנית לכל השורות
  table-report-service.js          דוח Word טבלאי (קבוצות/מיקומים)
  placeholder-mapper.js            התאמת תג↔עמודה אוטומטית
  docx-builder-bridge.js           חשיפת ספריית docx (window.docxBuilder)
  table-reader.js                  קריאת טבלת זימונים מהדף
  sp-lookup-service.js             שירות לבונה רשימת העבודה
  template-store.js / summons-export-service.js  עזרי תבניות/ייצוא

index.umd.js                       ספריית docx (יצירת Word)
js/docx-preview.min.js             רינדור docx ל-HTML להדפסה
jszip.min.js                       JSZip (קריאת/כתיבת docx ו-zip)
css/app.css                        עיצוב משותף לכל המסכים
templates/template1..3.docx        תבניות Word לדוגמה
```

---

## 12. כללי קוד בפרויקט

- OOP, ללא כפילות קוד (services/components לשימוש חוזר), קוד נקי.
- ללא הערות מיותרות בקוד — לכל היותר 1–2 שורות שמסבירות את רעיון הליבה של מתודה.
- לא לכתוב README בכל שינוי — רק בעדכון גדול.
