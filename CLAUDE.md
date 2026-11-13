# دفتر التفتيش الميداني — Field Inspection Notebook

Read this whole file before doing anything. It is the hand-off from the original claude.ai chat where the project was designed with the user.

## Who and why

- **User:** جاسم خميس جابر الحمادي (Jasim Al Hammadi), inspector at هيئة البيئة (Environment Agency – Abu Dhabi), Al Dhafra region. Name on reports: **جاسم الحمادي**.
- **Language:** talk to him in Arabic (Gulf dialect, short and clear). He uses an **iPhone** (main) and a **Windows PC**. His car has **CarPlay** and a phone holder.
- **Working style he asked for:** when something is unclear, ask **one question at a time**. He wants finished, working results, not drafts. Explain steps one by one when he has to do something himself (he is not a developer).
- **Goal:** he pilots this app himself now. If it works, he presents it to executive management (presentation + demo video) so it is adopted for all inspectors, possibly from **1 January 2027**. Build for one inspector now, but keep it ready to scale to many.

## The inspection work being recorded

Working days **Monday–Friday** (Saturday/Sunday off), roughly 07:00–15:00; Friday is a half day.

**No fixed weekday schedule (decided 2026-11-05).** The old Mon/Wed coast, Tue/Thu Habshan, Fri Madinat Zayed table was only for back-filling 2026 reports with a ready Excel sheet; that idea is dropped. The app must not suggest or pre-select an area by weekday: he opens a visit and picks the area himself every time. Do not add a schedule back.

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
- The exported workbook carries **calculated values** for every formula (`fillCachedValues` in `src/app.html`, a small formula engine for the functions the template uses). Without them the iPhone Files preview and mobile mail previews show 0 and 1899/12/31, because they never recalculate. If you add a formula function to the template, add it to the engine and re-run the LibreOffice comparison (all formula cells must match except TODAY()).

### `src/app.html` → the phone app (single file, vanilla JS, RTL)
Tabs: اليوم · السجل · التصدير · المساعد · الإعدادات.
- Start visit, which stamps arrival time (area starts empty; he picks it). Area chips, site select (add new), coordinates (GPS or paste), observations (type, description, action, photos), departure, finish.
- Photos are resized to 1920 px JPEG with a burned-in stamp: date, time, site, coordinates, inspector name.
- Export builds the **full workbook** by filling `template_clean.xlsx` (embedded as base64, `/*TPL*/""` placeholder) with JSZip XML patching: all visits into the log, report inputs set to the chosen period, the summary note into the matching notes cell. ZIP option adds the photos folder.
- Assistant features (only when Claude `sample` is available): data Q&A chat, formal rewrite of an observation, report summary and recommendations, parsing pasted «مساعد القيادة» text into visits (`### بداية` / `### مشاهدة` blocks; there is a keyword fallback without AI).
- Storage: IndexedDB always; in the claude.ai artifact it also syncs to the `db` capability at `data/users/<uid>/app` (settings) and `data/users/<uid>/app/visits/<visitId>`, with photos in `assets` (photo ids `a:<assetId>:<takenAt>`, local `l:<key>:<takenAt>`).

### Two deployments
1. **claude.ai artifact** (`dist/artifact.html`): this is what the user uses today. Private to him. Capabilities: downloads, db, user, assets, sample. Inside Claude the page **cannot use the microphone, GPS, or open the camera by itself** (sandbox), and the Claude iOS app wipes page storage when closed (hence the cloud `db`).
2. **GitHub Pages** (`index.html`): standalone, installable to the home screen. Here mic (Web Speech API), GPS, and camera work, but there is **no** `db`/`assets`/`sample` (no AI, data only in IndexedDB on the phone).

Build: `python3 build.py` (after `python3 excel/make_xlsx.py --clean` if the template changed). Commit `index.html`. GitHub Pages serves the repo root of `main`.

## Update 2026-11-13 (after the management submission) — the seven fixes
Approved with him on 2026-11-13 (management liked the demo; a discussion meeting is coming). Rule: **improve, never remove** anything shown in the videos/presentation. Backup before: branch `backup-2026-11-13`.
1. **Only finished visits go into Excel** (`doneVisits()`); an open visit (or an abandoned empty survey) is left out until finished. Export banner says so.
2. **Delete an open visit from اليوم**: «حذف» on the live card, tap twice to confirm (`delOpen`).
3. **Weekend = duty**: a visit on Sat/Sun shows the pill «مناوب»; Excel log column N shows «السبت (مناوب)»; weekly report row 18 is «السبت والأحد» (عطلة نهاية الأسبوع / مناوب: areas), total moved to row 19 (rows 20+ unchanged, notes still B77).
4. **Working day without visits**: card «أيام دوام بدون زيارات» on اليوم (Mon–Fri of the last 14 days, from `settings.since` or the first visit) — he picks اجتماع / مكتب / إجازة (`settings.dayStatus`, synced with settings). **Not** set automatically (could be a meeting in Abu Dhabi). Excel: new sheet7 «سجل الأيام» (names `D_Date`, `D_Status`), and the weekly day table shows the status in «المناطق» when the day has no visits.
5. **No back-dating**: the date field in the visit/survey editors is read-only; «استيراد من مساعد القيادة» skips items not dated today.
6. **No finished visit without coordinates**: «إنهاء» refuses without lat/lng; the voice assistant keeps the visit open and opens the editor.
7. **Review flag**: `geoIssue(v)` → «يحتاج مراجعة: …» in Excel column M (orange) and a pill in the app, when coordinates are outside the rough `AREA_BOX` of the chosen area, were typed/pasted (`v.geo="paste"`; GPS sets `v.geo="gps"`), or are missing. Boxes are approximate — adjust `AREA_BOX` if a real site falls outside. Areas without a box are not checked.
- **Excel locking** (clean template only): sheets «سجل الزيارات» and «سجل الأيام» are protected (password `daftar2026`, filter/sort allowed). The app gives filled log rows the unlocked styles of hidden reference row 1506 (`_s` override in `patchRow`), so recorded rows stay editable and empty rows cannot be typed into. This is a deterrent, not security (the repo is public; real protection = agency servers).
- **Auto summary (2026-11-13, his decision):** on GitHub Pages there is no AI, so «اكتب الملخص بالمساعد» is hidden. Added «اكتب الملخص تلقائيًا» (`autoSummary()`): a formal paragraph from the period's finished visits (counts, areas, top types, clean visits, surveys, review flags) + up to 4 rule-based recommendations; asks to confirm before replacing existing text. Placeholder no longer promises the AI. The AI summary moves to the second version (with the API proxy) — after initial approval.
- **Period picker (2026-11-13):** export range «ربع / نصف / سنة» (`S.range="period"`, `S.pYear`, `S.pName`, `PERIODS` = the Excel periodic list). The periodic report F4 and the note cell (B87) follow it; a custom range inside one month → monthly report (B63); a custom range over several months → the smallest matching period. The auto summary names the period. Purpose (his words): whoever evaluates inspectors can read each one's work from the Excel for any period.
- Tested: Playwright run (gps/mismatch/no-coords/delete/day status/export), LibreOffice recalculation = app cache on all 19,632 formula cells, 0 errors.

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
- Each answer is confirmed briefly (type, area/site names). Free dictation (the description) is **not** read back — only «سجّلت الوصف.» (his request, 2026-11-06). Two silences → pauses with a big green «تكلّم» button (also the fallback if iOS refuses to start the mic outside a tap). Camera: `getUserMedia` live preview; «صوّر» or tapping the preview captures; photos go through the same `storePhoto` (1920 px + stamp). GPS taken at start. Departure stamped on «نعم» at «ننهي الزيارة؟», then a spoken summary. Screen wake lock while running.
- Visits get `src:"voice"` (shows the «صوتي» pill). Same data model, Excel export unchanged.
- **Unverified on a real iPhone:** whether iOS lets the mic restart after speech without a tap, and whether speech recognition works while the camera is open. Fallbacks exist (tap «تكلّم», tap the preview, «خلاص، كمّل»). Ask him how it behaved.

### ⚠️ Voice experiments of 2026-11-06 were REVERTED (his decision, 17:48)
- The app (src/app.html, index.html) is back to yesterday's working version (commit 3cb89c1): phone voice (Majed) as before, original sentences («وش», «بدأت زيارة الساعة…», «سجّلت: …», summary with departure time).
- Why: Grady clips → earpiece + recognizer deaf after the mic on iPhone Safari; Plan B (best phone voice + picker) still choppy and delayed, and he could not hear a difference. He wants to stop spending time on voice in the web app and revisit it in the native app («البرنامج المطور»).
- All of today's work is kept in branch `backup-2026-11-06d` (clips player, Plan B voice picker) and documented in the three sections below. Those sections describe that branch, NOT the current app.

### Voice — current state (Plan B, 2026-11-06 evening) — READ THIS FIRST
- **The recorded Grady clips are OFF on the web (`VA_CLIPS=false`).** Second field test on his iPhone: the first two sentences played, then Safari moved the clips to the earpiece once the mic had been used, and the recognizer stopped hearing him («حبشان الشرقية» several times; he finished with the on-screen chips). This is Safari's handling of `<audio>` + mic in a web app, not fixable from the page. The phone voice (`speechSynthesis`) is not affected — it worked with the mic in the morning.
- `speak()` now goes straight to `speakTTS()`. `vaVoice()` picks the best Arabic voice Safari exposes (Premium/«استثنائي» > Enhanced/«محسّن» > compact; ar-AE > ar-SA), or the one he chose in Settings (`S.settings.vaVoice` = voiceURI). `vaRate()` = `S.settings.vaRate` (أبطأ 0.9 / عادي 1 / أسرع 1.1).
- Settings → «صوت المساعد الآلي»: lists the Arabic voices Safari gives the page, a speed choice and «جرّب الصوت». It tells us whether the Premium voice he downloaded is available to the web app at all.
- The clip code (preload, trim, WAV join, unlock) and the 115 clips stay for a future native app, where the audio route can be controlled.
- Backup before Plan B: branch `backup-2026-11-06c`.

### Recorded voice (2026-11-06) — replaces the phone's robotic «ماجد» voice
- He found iOS TTS (Majed) choppy and robotic; downloading «ماجد (استثنائي)» on the iPhone did not help (Safari web apps don't get it).
- All assistant speech is now **pre-recorded clips**: voice **Grady**, ElevenLabs v4 via the Higgsfield connector (voice_id `e2a2d2e6-9ed2-59cd-82af-feaa27f8a678`, preset). 115 clips: 56 phrases, numbers 1–10, 6 types, 7 areas, 36 sites. Wording approved by him (Project doc «قائمة-أصوات-المساعد»; links in «روابط-أصوات-المساعد»).
- `speak(text)` → `toClips(text)` maps the on-screen sentence to clip ids (longest phrase first, names at word boundaries, digits 1–10). `hush()` stops everything (replaces `speechSynthesis.cancel()`).
- **Playback fix (2026-11-06, after his first field test):** playing the clips one after another on an `<audio>` failed on iOS after the mic had been used: he heard «مدينة زايد» and then nothing (the next clip never played), and every clip loaded on demand (delay). Now `vaPreload()` fetches and decodes all clips when the assistant starts (CloudFront sends `access-control-allow-origin: *`; decoding uses an `OfflineAudioContext`, so no live audio graph), `trimPCM` cuts the ~0.3 s of silence each clip has, and `wavBlob` joins the pieces of one sentence (80 ms gap) into ONE WAV that is played with a single `play()` on one `<audio>` element. `vaUnlock()` plays a silent WAV inside the start tap so later plays are allowed. A clip that cannot load within 6 s, or a failed play, falls back to `speakTTS`.
- Short answers (area, site, yes/no) now end after 1.0 s of silence instead of 1.3 s (he reported a delay).
- Tested in headless Chromium: all 115 real clips fetch cross-origin and decode (≈3 MB, ≈11 s for all); a full voice visit (fake recognition) plays each sentence once and whole; CDN down → phone voice. Not testable here: iOS Safari itself and car Bluetooth (with the car's Bluetooth connected it did not work at all in his first try).
- Text pieces with no clip (a site/type he adds later, numbers > 10) are skipped silently; the full text stays on screen. If a clip fails to play (offline, files gone) → `speakTTS` (old phone voice); after 2 failures in a row it stays on the phone voice.
- **The clips are played from Higgsfield's CDN** (`CLIP_BASE`, d8j0ntlcm91z4.cloudfront.net), not from this repo: the cloud workspace could not download them (network policy). Next step agreed with him: move them into the repo (`audio/`) and point `CLIP_BASE` there.
- **Any new spoken sentence must be recorded with the same voice and added to `CLIP_SRC`/`CLIP_TXT`**, and the display text must match the clip text exactly, or it will be silent. Wording changes made for the clips: «وش»→«شو», «مو»→«مب», «بدأت الزيارة.» (no time), «سجّلت الوصف.», «أضفت الموقع.», generic «ما لقيت هذا الموقع في القائمة…», summary without departure time and with «مشاهدتين/مخالفتين».
- Not yet field-tested: gaps between joined clips, and whether iOS plays the clips at full volume after the mic was used.

## «زيارة معاينة مزرعة» — farm survey for well-drilling requests (built 2026-11-02, waiting for his field test)
- Button «زيارة معاينة مزرعة (طلب حفر بئر)» on اليوم. Visit with `kind:"survey"` and `sv:{owner,farmNo,reqNo,company,fenced:"yes"|"no",wOk,wWeak,wBur,palms,gh}`; photos live in one obs `{id:"sv",type:"معاينة مزرعة",action:"معاينة"}` so sync/export/photo code is shared. Editor `vSurvey`, check `svCheck` (owner or farm no., area, محضر, fence).
- He does **not** decide eligibility; the form only records facts. Drilling companies are in `settings.companies` (editable in الإعدادات).
- «إنهاء المعاينة» builds a PDF form (`svPages` draws A4 canvases, `jpegPdf` wraps the JPEGs, Google Maps link annotation on the coordinates), then «إرسال النموذج بالإيميل» uses `navigator.share` (Mail) with fallback `offerFile`. The form is a **proposal** to replace the paper form; if management keeps the old paper form, rebuild `svPages` to match it.
- Excel: survey row type «معاينة مزرعة», action «معاينة», description = `svDesc`. It counts as a visit; the observation flag (col V) excludes it. Each report (weekly, monthly, periodic) has a KPI card «معاينات المزارع» (COUNTIFS on type «معاينة مزرعة») so the visit total reconciles (added 2026-11-05).

## Tutorial video (مقطع الشرح)
- v2 (2026-11-02, 4:40): https://d2ol7oe51mr4n9.cloudfront.net/user_3KxcXNsGX1h7cj6wUkRO8K9aqTQ/f9f09cad-5cdf-469a-9193-2424e142a8c0.mp4 (v1 4:00 had «بدع زايد» and no survey).
- Built in the Higgsfield sandbox: Playwright captures of the live GitHub Pages app (fake speech + fake camera), scenes drawn from an HTML template (Reem Kufi titles, Readex Pro text, chapter chips), Excel views rendered from the real export, crossfades, original narration reused + new «المعاينة» chapter. Narrator voice = ElevenLabs preset Harrison (573e5163-59b3-4926-aab1-951ef2985f81). If the app UI changes, recapture and re-render the same way.

## Testing
- Voice flow: Playwright with a mocked `webkitSpeechRecognition` (scripted answers queue) and `speechSynthesis`, Chromium flags `--use-fake-device-for-media-stream --use-fake-ui-for-media-stream`; route the cdnjs JSZip URL to a local copy (cdnjs is blocked in the cloud sandbox).
- `node --check` on the extracted `<script>`, then a Playwright run (Chromium) through a visit, photo, export, and reload.
- Recalculate exported workbooks with LibreOffice and confirm values on the weekly report.
