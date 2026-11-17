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
Each worksheet becomes a subject. The worksheet tab name is the subject name; the workbook filename is ignored.

Use these columns, in this exact order:
1. Question
2. Option A
3. Option B
4. Option C
5. Option D
6. Answer
7. Correct Answer (displayed in question review)

The Answer column accepts A/B/C/D or text matching one of the four options. The Correct Answer column should contain the correct answer text and is displayed explicitly in the review/export; if it matches an option, it is used as the scoring key. If column 7 is blank or does not match an option, column 6 is used as the scoring key. A header row is supported. Blank question rows are ignored. Rows with missing options or an unrecognised answer are skipped.

A sample workbook can also be downloaded from inside the app.

## Features
- Subject detection from worksheet tab names
- Manual question allocation per subject, with available-question validation
- Configurable duration, positive marks, and negative marks
- Shuffled questions and independently shuffled options
- Timer, question palette, answer states, mark-for-review, keyboard shortcuts
- Manual submit and automatic submission at zero
- Overall and subject-wise score analysis
- Question review showing selected and correct options
- Excel export with question review and subject summary
- Browser-local question bank and recent results

## Privacy and storage
The workbook and results are processed in the browser and saved in this browser's localStorage. They are not uploaded to a server. Clearing browser site data removes saved data. Browser storage is not a substitute for a backup.

## Notes
- The app currently supports four options per question.
- The correct answer must be A/B/C/D or match one of the option texts.
- The app stores the correct answer separately from the shuffled option order, so scoring remains correct after shuffling.
- For production use, consider adding a confirmation before replacing an existing question bank and a more robust crash-recovery mechanism for an in-progress exam.
