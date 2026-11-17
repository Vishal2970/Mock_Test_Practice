# MockLab — Excel-based mock test app

A client-side React + TypeScript single-page application. No backend, database, login, or API is required.

## Run locally
1. Install Node.js 18 or newer.
2. Extract this ZIP.
3. In the project folder run:
   ```bash
   npm install
   npm run dev
   ```
4. Open the local URL printed by Vite.

For a production build:
```bash
npm run build
npm run preview
```

## Excel workbook format
Each worksheet becomes a subject (section). The worksheet tab name is the subject name; the workbook filename is ignored. Hidden and empty worksheets are skipped and listed in the import details.

Columns, in this order (a header row is optional but, when present, columns are matched by header name):
1. Question
2. Option A
3. Option B
4. Option C
5. Option D
6. Answer
7. Correct Answer (shown in review/export)
8. Explanation (optional)

- **Answer** is the scoring key: A/B/C/D (also `b)`, `Option B`) or text matching one of the four options.
- **Correct Answer** is a display check. If it disagrees with Answer, or matches no option, the Answer column is still used and the row is listed under warnings.
- With a header row you can place **Explanation** in any column (for example column 7); without headers it is read from column 8.
- Rows are skipped, and listed with sheet and row number, when a question or option is missing, the answer is missing or unrecognised, or two options have the same text.
- Exact duplicate questions inside one worksheet are imported once.

A sample workbook can be downloaded from inside the app.

## Taking a test
- **Sections:** questions are grouped by subject. You choose the count per subject; subjects are shuffled independently, and the four options of every question are shuffled independently. Numbering runs across the whole exam (e.g. Mathematics Q1–30, Reasoning Q31–50).
- **Overall timer (default):** one countdown for the whole exam. All sections stay open and can be revisited; the test is submitted automatically at zero.
- **Section-wise timers (optional):** enter minutes for every subject that has questions. Each section runs on its own clock; when it ends the section locks and the next one starts. Sections are forward-only. “Submit section” moves on early (unused time is forfeited) and can be switched off in the setup screen. If you choose this mode but leave every time empty, the overall timer is used. Entering times for only some subjects is not allowed.
- Timers are based on clock timestamps, so they stay accurate in background tabs, and a page refresh during an exam restores it (sections that expired while the page was closed are locked, and an exam whose time ran out is submitted).
- **Results:** overall and per-subject totals, accuracy (correct ÷ attempted), marks, negative marks, time taken and time per subject, plus a question-wise review. Past results can be reopened from the history list, and “Retake same questions” rebuilds the test with fresh shuffling.
- **Excel export:** sheets *Question Review*, *Subject Analysis* (original columns first, new ones appended) and *Overall Summary*.

## Privacy and storage
The workbook and results are processed in the browser and saved in this browser's localStorage. They are not uploaded to a server. Clearing browser site data removes saved data. Browser storage is not a substitute for a backup.

## Notes
- The app supports four options per question.
- The correct option index is fixed at the moment options are shuffled, so scoring stays correct after shuffling.
- Full question-by-question reviews are stored for the 5 most recent tests; older history entries keep their summary only, to save browser storage.
- A browser cannot enforce fullscreen or prevent tab switching; the fullscreen option is a convenience, not proctoring.
