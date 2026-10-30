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

## Field test log
- 2026-10-29, GitHub Pages version installed on his iPhone home screen: GPS fills automatically ✓, the «إملاء صوتي» button works ✓. Bug: iOS delivers dictation in chunks that were glued without spaces; fixed by joining the results with spaces. He said more adjustments are needed; ask him what they are.

## Lessons learned (do not repeat)
- iOS Siri Shortcuts were too hard for him to build by hand; I cannot send a signed shortcut from Windows.
- A Claude Project + voice mode «مساعد القيادة» did not fit: he must tap the mic each time and it is not integrated with the app.
- Local-only storage inside the Claude iOS app lost his data. Never rely on it there.

## «المساعد الآلي» — voice-driven visit (built 2026-10-29, waiting for his field test)
Dark button «المساعد الآلي — زيارة بالصوت» on the اليوم tab (shown only where SpeechRecognition + speechSynthesis exist and **not** inside the claude.ai artifact). Code: section `المساعد الآلي` in `src/app.html` (functions `va*`, overlay `#va`).
- Deterministic state machine: `area → site → hasObs → type → desc → action → photo → more → finish`, plus transient steps `pick` («تقصد X؟» for 0.6–0.8 matches), `newSite` (unknown site → «أضيفه موقع جديد؟», saved to settings), `cancel`.
- Shortcuts: saying a site name at the area step («حبشان الشرقية», «ليوا الثروانية») fills both; naming a type at «في مشاهدات؟» skips the yes. Areas with one site fill the site automatically.
- Matching: `arNorm` (hamza/ta-marbuta/alef-maqsura/diacritics), `arPhon` (Gulf: ج→ي, ق→ك, ذ→ز, ث→س, ظ→ض), filler words and «ال» stripped, Levenshtein over word windows (partial-phrase matches ×0.9). Types also by keyword (`VA_TYPE_KW`) or spoken number. After 2 misses it reads the options.
- Commands anytime (≤3 words): «عيد», «رجوع», «إلغاء» (confirm, deletes the visit), «وقف». «إيقاف» button keeps what was recorded and opens the visit editor (an empty visit is deleted).
- Each answer is read back. Two silences → pauses with a big green «تكلّم» button (also the fallback if iOS refuses to start the mic outside a tap). Camera: `getUserMedia` live preview; «صوّر» or tapping the preview captures; photos go through the same `storePhoto` (1920 px + stamp). GPS taken at start. Departure stamped on «نعم» at «ننهي الزيارة؟», then a spoken summary. Screen wake lock while running.
- Voice: Settings card «صوت المساعد الآلي» picks the Arabic voice (`settings.voice` = voiceURI) and speed (`settings.rate`, 0.8–1.2); default prefers downloaded «Enhanced/Premium» voices. iOS: download one under Settings → Accessibility → Spoken Content → Voices → Arabic.
- Visits get `src:"voice"` (shows the «صوتي» pill). Same data model, Excel export unchanged.
- **Unverified on a real iPhone:** whether iOS lets the mic restart after speech without a tap, and whether speech recognition works while the camera is open. Fallbacks exist (tap «تكلّم», tap the preview, «خلاص، كمّل»). Ask him how it behaved.

## Field test 2026-10-29/30 (voice assistant)
- Works in the field. He finds the conversation too rigid (only fixed keywords) and the default iOS voice unconvincing. Next ideas: free-form understanding via the Claude API (needs his approval and a paid key), and higher-quality pre-recorded prompts. He wants Higgsfield for the demo video (no connector exists; prompts are in the script doc).
- Deliverables made: management deck https://claude.ai/artifact/LcgcN5zez7KvCU19fxWM8E (12 slides, placeholders [__] for his real numbers) and video script doc https://claude.ai/artifact/9b2DDnFSSkfgT72G1dcyZS (11 scenes, English Higgsfield prompts per scene; open question there: AI-generated inspector or himself).
- 2026-10-30: he connected the official Higgsfield MCP connector (https://mcp.higgsfield.ai/mcp, OAuth) at claude.ai. Next session: read the script doc, generate the scenes one by one with Higgsfield (it spends his credits: one generation per scene, redo only what he asks), keep one consistent inspector character, no official logos. Also try Higgsfield voiceover for the narrator and maybe better assistant prompts.
- **Video style rule (religious, firm):** no realistic people. Characters are faceless toy-brick (Lego-like) figures: blank face, no eyes/nose/mouth; kandura, ghutra + agal, dark green vest with a plain emblem patch (the real agency logo is overlaid in editing, AI garbles logos). Higgsfield account: his Plus account (glitchbroccoli1987). Candidate character images: jobs 75c6ba72-664b-403a-8eab-4bea26647eb0 and 8a5a072f-13c6-48d5-bea4-acb1a62483bc (gpt_image_2_5). Video plan: kling3_0 std, 5 s, sound off ≈ 7.5 credits/scene.
- The voice picker commit (460c146) is on `claude/busy-planck-ovjvn7`, not merged to main yet; ask before merging.

## Testing
- Voice flow: Playwright with a mocked `webkitSpeechRecognition` (scripted answers queue) and `speechSynthesis`, Chromium flags `--use-fake-device-for-media-stream --use-fake-ui-for-media-stream`; route the cdnjs JSZip URL to a local copy (cdnjs is blocked in the cloud sandbox).
- `node --check` on the extracted `<script>`, then a Playwright run (Chromium) through a visit, photo, export, and reload.
- Recalculate exported workbooks with LibreOffice and confirm values on the weekly report.
