# دفتر التفتيش الميداني — Field Inspection Notebook

Read this whole file before doing anything. It is the hand-off from the original claude.ai chat where the project was designed with the user.

## Who and why

- **User:** جاسم خميس جابر الحمادي (Jasim Al Hammadi), inspector at هيئة البيئة (Environment Agency – Abu Dhabi), Al Dhafra region. Name on reports: **جاسم الحمادي**.
- **Language:** talk to him in Arabic (Gulf dialect, short and clear). He uses an **iPhone** (main) and a **Windows PC**. His car has **CarPlay** and a phone holder.
- **Working style he asked for:** when something is unclear, ask **one question at a time**. He wants finished, working results, not drafts. Explain steps one by one when he has to do something himself (he is not a developer).
- **Goal:** he pilots this app himself now. If it works, he presents it to executive management (presentation + demo video) so it is adopted for all inspectors, possibly from **1 January 2027**. Build for one inspector now, but keep it ready to scale to many.

## The inspection work being recorded

Working days **Monday–Friday** (Saturday/Sunday off), roughly 07:00–15:00; Friday is a half day.

| Day | Usual areas |
|---|---|
| Monday, Wednesday | ليوا (one or two محاضر per visit) / المرفأ / الرويس / غياثي / sometimes السلع |
| Tuesday, Thursday | حبشان (East or West, alternating) |
| Friday | بدع زايد |

Each visit: area, site (محضر), date, arrival/departure time, GPS coordinates, observations with photos.
- If a violation is observed **and the offender is present** → «تم تحرير مخالفة».
- If no offender present → «مشاهدة فقط» (record observation, photos, coordinates, time).
- If nothing found → type «لا توجد مشاهدات», action «لا يوجد».
- Violation details (number, offender name) are intentionally **not** recorded yet, only whether one was issued.

**Observation types** (user-approved; he can add more in-app):
حرق مخلفات · تحميل مياه جوفية في صهاريج · تصريف / كب صرف صحي · رمي نفايات عامة · تسرب زيوت أو مواد كيميائية · حفر آبار غير مصرح · لا توجد مشاهدات
Removed on purpose: construction waste (belongs to the municipality), grazing / woodcutting (belongs to the wildlife section), sand excavation (replaced by unauthorized wells).

**Sites:** ليوا has ~50 محاضر (list in `excel/make_xlsx.py` and `DEFAULT_AREAS` in `src/app.html`, from public sources). حبشان: الشرقية / الغربية. المرفأ also includes الطريف. Other areas are single sites for now; the user adds sub-sites from the app.

**Reports required:** weekly → monthly → quarterly (Q1–Q4) → half-year (H1/H2) → annual.

## What exists

### `excel/make_xlsx.py` → the Excel workbook
openpyxl generator. `--clean` writes `excel/template_clean.xlsx` (no sample rows, opens on the weekly report).
Sheets: طريقة الاستخدام · سجل الزيارات · التقرير الأسبوعي · التقرير الشهري · التقرير الدوري · الإعدادات.
- **سجل الزيارات:** one row = one observation; rows 5–1504. Input columns **B–M**: التاريخ, المنطقة, الموقع, وقت الوصول, وقت المغادرة, نوع المشاهدة, وصف المشاهدة, الإجراء, عدد الصور, خط العرض, خط الطول, ملاحظات. Formula columns N–P, plus hidden helpers Q–X (month, year, week start (Monday), new-visit flag = first row for date+site, new-area flag, observation flag, violation flag, weekly index).
- Reports use only Excel-2007 functions plus `_xlfn.TEXTJOIN` (LibreOffice-safe; no FILTER/UNIQUE/XLOOKUP). Inputs: weekly D4 (any date in the week), monthly D4 year + F4 month name, periodic D4 year + F4 period name.
- Notes cells: weekly B77, monthly B63, periodic B87.
- Photos are **not** embedded: they live in a `الصور/` folder next to the workbook, one sub-folder per visit `YYYY-MM-DD_<site>`. Column P links to it.
- Verify with LibreOffice after any change (zero formula errors).

### `src/app.html` → the phone app (single file, vanilla JS, RTL)
Tabs: اليوم · السجل · التصدير · المساعد · الإعدادات.
- Start visit, which stamps arrival time and suggests the area from the weekday schedule. Area chips, site select (add new), coordinates (GPS or paste), observations (type, description, action, photos), departure, finish.
- Photos are resized to 1920 px JPEG with a burned-in stamp: date, time, site, coordinates, inspector name.
- Export builds the **full workbook** by filling `template_clean.xlsx` (embedded as base64, `/*TPL*/""` placeholder) with JSZip XML patching: all visits into the log, report inputs set to the chosen period, the summary note into the matching notes cell. ZIP option adds the photos folder.
- Assistant features (only when Claude `sample` is available): data Q&A chat, formal rewrite of an observation, report summary and recommendations, parsing pasted «مساعد القيادة» text into visits (`### بداية` / `### مشاهدة` blocks; there is a keyword fallback without AI).
- Storage: IndexedDB always; in the claude.ai artifact it also syncs to the `db` capability at `data/users/<uid>/app` (settings) and `data/users/<uid>/app/visits/<visitId>`, with photos in `assets` (photo ids `a:<assetId>:<takenAt>`, local `l:<key>:<takenAt>`).

### Two deployments
1. **claude.ai artifact** (`dist/artifact.html`): this is what the user uses today. Private to him. Capabilities: downloads, db, user, assets, sample. Inside Claude the page **cannot use the microphone, GPS, or open the camera by itself** (sandbox), and the Claude iOS app wipes page storage when closed (hence the cloud `db`).
2. **GitHub Pages** (`index.html`): standalone, installable to the home screen. Here mic (Web Speech API), GPS, and camera work, but there is **no** `db`/`assets`/`sample` (no AI, data only in IndexedDB on the phone).

Build: `python3 build.py` (after `python3 excel/make_xlsx.py --clean` if the template changed). Commit `index.html`. GitHub Pages serves the repo root of `main`.

## Lessons learned (do not repeat)
- iOS Siri Shortcuts were too hard for him to build by hand; I cannot send a signed shortcut from Windows.
- A Claude Project + voice mode «مساعد القيادة» did not fit: he must tap the mic each time and it is not integrated with the app.
- Local-only storage inside the Claude iOS app lost his data. Never rely on it there.

## NEXT TASK: «المساعد الآلي», a voice-driven visit flow in the standalone app

The user's own description: one button on the home screen. He taps it once, and the app **talks** (speechSynthesis, Arabic) and **listens** (webkitSpeechRecognition, `ar-AE`) in a loop, filling the visit as they go:
1. Open a new visit (date and time automatic).
2. «وين المنطقة؟», then «وش الموقع؟», matched against the areas and sites lists (fuzzy, Gulf pronunciation), confirmed back by voice.
3. Take GPS coordinates automatically.
4. «في مشاهدات؟»
   - **Yes:** ask for the type and description, then «تم تحرير مخالفة ولا مشاهدة بس؟». Open the in-app camera (getUserMedia live preview, so a spoken «صوّر» or one tap captures). Ask «في مشاهدة ثانية؟».
   - **No / «لا توجد»:** record «لا توجد مشاهدات».
5. «ننهي الزيارة؟», then stamp departure time, save, and speak a short summary.

Requirements: hands-free, short prompts, read back what was understood, allow «عيد» (repeat), «رجوع» (back) and «إلغاء» (cancel). Must be safe to use while driving (no reading the screen). Large «إيقاف» button. A deterministic state machine plus keyword/fuzzy matching is fine for v1 (no AI available on Pages). The optional Claude API key for free-form understanding is discussed but **not** approved or paid yet. Keep the same data model and Excel export, so both deployments stay compatible.

Paid options discussed, **none bought yet**: Claude API (only if free speech understanding is needed), Apple Developer Program (later, after management approval; needs a Mac), GitHub Pro (not needed).

## Testing
- `node --check` on the extracted `<script>`, then a Playwright run (Chromium) through a visit, photo, export, and reload.
- Recalculate exported workbooks with LibreOffice and confirm values on the weekly report.
