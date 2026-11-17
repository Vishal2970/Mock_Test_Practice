// Pure exam logic (no React, no DOM except localStorage helpers). Every function that depends on time takes `now`
// explicitly so timer behaviour can be tested without waiting.

export type Question = { id: string; subject: string; question: string; options: string[]; correct: string; correctAnswerDisplay: string; explanation: string };
export type ExamQuestion = Question & { correctIndex: number };
export type Subject = { name: string; count: number; available: number };
export type AnswerState = { selected: number | null; marked: boolean; seconds: number };
export type TimingMode = 'overall' | 'section';
export type SectionConfig = { subject: string; count: number; minutes: number | null };
export type ExamConfig = { mode: TimingMode; overallMinutes: number; sections: SectionConfig[]; allowEarlyNext: boolean; positive: number; negative: number };
export type Section = { subject: string; start: number; end: number; minutes: number | null };
export type ExamSession = {
  v: 1; id: string; fileName: string; config: ExamConfig;
  questions: ExamQuestion[]; answers: AnswerState[]; sections: Section[];
  current: number; activeSection: number; sectionPos: number[];
  startedAt: number; sectionStartedAt: number; qStartedAt: number;
};
export type ReviewItem = { q: Question; selected: number | null; correctIndex: number; seconds: number };
export type SubjectResult = { name: string; total: number; attempted?: number; correct: number; wrong: number; skipped: number; score: number; accuracy?: number; seconds?: number; timeLimit?: number | null };
export type TestResult = {
  id: string; date: string; total: number; attempted: number; correct: number; wrong: number; skipped: number; score: number; maxScore: number; duration: number;
  subjects: SubjectResult[]; review: ReviewItem[];
  // Optional: absent on results saved by earlier versions.
  positive?: number; negative?: number; negativeMarks?: number; accuracy?: number; mode?: TimingMode; autoSubmitted?: boolean; fileName?: string; config?: ExamConfig;
};
export type Settings = { timingMode: TimingMode; duration: string; positive: string; negative: string; sectionMinutes: Record<string, string>; allowEarlyNext: boolean; fullscreen: boolean };
export type StoreData = { history: TestResult[]; questions: Question[]; subjects: Subject[]; counts: Record<string, string>; fileName: string; settings?: Settings };

export const MINUTE = 60000;
export const REVIEW_KEEP = 5;
export const HISTORY_KEEP = 30;
const STORE_KEY = 'mocklab_v1';
const SESSION_KEY = 'mocklab_v1_session';

export const uid = () => Math.random().toString(36).slice(2) + Date.now().toString(36);
export const shuffle = <T,>(items: T[], rng: () => number = Math.random) => {
  const a = [...items];
  for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(rng() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; }
  return a;
};
export const fmt = (s: number) => { s = Math.max(0, Math.round(s)); return `${Math.floor(s / 60).toString().padStart(2, '0')}:${(s % 60).toString().padStart(2, '0')}`; };
const round3 = (n: number) => Math.round(n * 1000) / 1000;

/* ---------------------------------------------------------------- Excel parsing */

export type RawSheet = { name: string; hidden: boolean; rows: unknown[][]; rowOffset: number };
export type RowIssue = { sheet: string; row: number; reason: string };
export type ImportReport = { issues: RowIssue[]; warnings: string[]; duplicates: number; skippedSheets: { name: string; reason: string }[] };

const norm = (v: unknown) => String(v ?? '').toLowerCase().replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
const HEADER_NAMES: Record<string, string[]> = {
  question: ['question', 'questions', 'question text'],
  a: ['option a', 'opt a', 'option 1', 'choice a'], b: ['option b', 'opt b', 'option 2', 'choice b'],
  c: ['option c', 'opt c', 'option 3', 'choice c'], d: ['option d', 'opt d', 'option 4', 'choice d'],
  answer: ['answer', 'ans', 'correct option', 'key', 'right answer'],
  correct: ['correct answer', 'correct answer text', 'correct text'],
  explanation: ['explanation', 'solution', 'rationale', 'explain', 'remarks'],
};
const DEFAULT_COL: Record<string, number> = { question: 0, a: 1, b: 2, c: 3, d: 4, answer: 5, correct: 6, explanation: 7 };
const LETTER = /^\(?\s*(?:option\s*)?([a-d])\s*[).:-]?\s*$/i;

function detectColumns(first: unknown[]): { header: boolean; col: Record<string, number> } {
  const found: Record<string, number> = {};
  first.forEach((cell, i) => { const n = norm(cell); for (const [k, names] of Object.entries(HEADER_NAMES)) if (!(k in found) && names.includes(n)) found[k] = i; });
  const header = HEADER_NAMES.question.includes(norm(first[0])) || Object.keys(found).length >= 3;
  if (!header) return { header: false, col: { ...DEFAULT_COL } };
  // A missing header falls back to its legacy position only when no other header already claimed that column.
  const claimed = new Set(Object.values(found));
  const col: Record<string, number> = {};
  for (const k of Object.keys(DEFAULT_COL)) col[k] = k in found ? found[k] : claimed.has(DEFAULT_COL[k]) ? -1 : DEFAULT_COL[k];
  return { header: true, col };
}

function resolveOption(value: string, opts: string[]): number {
  const m = LETTER.exec(value);
  if (m) return m[1].toUpperCase().charCodeAt(0) - 65;
  return opts.findIndex(o => o.toLowerCase() === value.toLowerCase());
}

/** Turns raw worksheet rows into validated questions plus a report of everything that was skipped or looks suspicious. */
export function parseSheets(sheets: RawSheet[]): { questions: Question[]; report: ImportReport } {
  const questions: Question[] = [];
  const report: ImportReport = { issues: [], warnings: [], duplicates: 0, skippedSheets: [] };
  for (const sheet of sheets) {
    const subject = sheet.name.trim() || 'Untitled';
    if (sheet.hidden) { report.skippedSheets.push({ name: subject, reason: 'hidden worksheet' }); continue; }
    const rows = sheet.rows;
    if (!rows.some(r => r.some(c => String(c ?? '').trim()))) { report.skippedSheets.push({ name: subject, reason: 'empty worksheet' }); continue; }
    const { header, col } = detectColumns(rows[0] || []);
    const cell = (r: unknown[], k: string) => (col[k] < 0 ? '' : String(r[col[k]] ?? '').trim());
    const seen = new Set<string>();
    let added = 0;
    rows.forEach((r, i) => {
      if (i === 0 && header) return;
      if (!r.some(c => String(c ?? '').trim())) return;
      const row = sheet.rowOffset + i + 1;
      const issue = (reason: string) => report.issues.push({ sheet: subject, row, reason });
      const q = cell(r, 'question');
      if (!q) return issue('missing question text');
      const opts = ['a', 'b', 'c', 'd'].map(k => cell(r, k));
      const ans = cell(r, 'answer');
      if (opts.some(o => !o)) return issue('missing option');
      if (!ans) return issue('missing answer');
      if (new Set(opts.map(o => o.toLowerCase())).size < 4) return issue('options must be four different texts');
      const key = resolveOption(ans, opts);
      if (key < 0) return issue('answer must be A/B/C/D or match an option');
      const shown = cell(r, 'correct');
      if (shown) {
        const k2 = resolveOption(shown, opts);
        if (k2 < 0) report.warnings.push(`${subject}, row ${row}: Correct Answer text matches no option; the Answer column is used`);
        else if (k2 !== key) report.warnings.push(`${subject}, row ${row}: Answer and Correct Answer disagree; the Answer column is used`);
      }
      const dupKey = norm(q) + '||' + opts.map(norm).join('|');
      if (seen.has(dupKey)) { report.duplicates++; return; }
      seen.add(dupKey);
      questions.push({ id: uid(), subject, question: q, options: opts, correct: opts[key], correctAnswerDisplay: opts[key], explanation: cell(r, 'explanation') });
      added++;
    });
    if (!added) report.skippedSheets.push({ name: subject, reason: 'no valid questions' });
    else if (/^sheet\d+$/i.test(subject)) report.warnings.push(`Worksheet “${subject}” has a default name and is used as a subject name; rename the tab if that is not intended`);
  }
  return { questions, report };
}

export function groupSubjects(questions: Question[]): Subject[] {
  const order = [...new Set(questions.map(q => q.subject))];
  return order.map(name => ({ name, count: 0, available: questions.filter(q => q.subject === name).length }));
}

/* ---------------------------------------------------------------- Building an exam */

type Group = { subject: string; questions: Question[] };

/** Options are shuffled here; correctIndex is fixed at the same moment, so scoring never depends on option text again. */
export function assembleExam(groups: Group[], config: ExamConfig, fileName: string, now: number, rng: () => number = Math.random): ExamSession {
  const questions: ExamQuestion[] = [];
  const sections: Section[] = [];
  const cfgSections: SectionConfig[] = [];
  for (const g of groups) {
    const start = questions.length;
    for (const q of g.questions) {
      const original = q.options.findIndex(o => o === q.correct);
      if (original < 0) continue;
      const order = shuffle(q.options.map((_, i) => i), rng);
      questions.push({ ...q, options: order.map(i => q.options[i]), correctIndex: order.indexOf(original) });
    }
    if (questions.length === start) continue;
    const minutes = config.sections.find(s => s.subject === g.subject)?.minutes ?? null;
    sections.push({ subject: g.subject, start, end: questions.length, minutes });
    cfgSections.push({ subject: g.subject, count: questions.length - start, minutes });
  }
  const timed = config.mode === 'section' && sections.length > 0 && sections.every(s => (s.minutes ?? 0) > 0);
  const mode: TimingMode = timed ? 'section' : 'overall';
  const finalConfig: ExamConfig = { ...config, mode, sections: cfgSections, overallMinutes: timed ? sections.reduce((n, s) => n + (s.minutes || 0), 0) : config.overallMinutes };
  return {
    v: 1, id: uid(), fileName, config: finalConfig, questions, answers: questions.map(() => ({ selected: null, marked: false, seconds: 0 })), sections,
    current: 0, activeSection: 0, sectionPos: sections.map(s => s.start), startedAt: now, sectionStartedAt: now, qStartedAt: now,
  };
}

/** Picks `count` random questions per section (each subject shuffled independently, subjects kept in worksheet order). */
export function buildExam(bank: Question[], config: ExamConfig, fileName: string, now: number, rng: () => number = Math.random): ExamSession {
  const groups = config.sections.filter(s => s.count > 0).map(s => ({ subject: s.subject, questions: shuffle(bank.filter(q => q.subject === s.subject), rng).slice(0, s.count) }));
  return assembleExam(groups, config, fileName, now, rng);
}

/* ---------------------------------------------------------------- Timing and navigation */

export const sectionOf = (s: ExamSession, qi: number) => s.sections.findIndex(sec => qi >= sec.start && qi < sec.end);
export const canAccess = (s: ExamSession, qi: number) => qi >= 0 && qi < s.questions.length && (s.config.mode === 'overall' || sectionOf(s, qi) === s.activeSection);

/** The timestamp at which the clock that currently governs the exam hits zero. */
export function deadline(s: ExamSession): number {
  return s.config.mode === 'section'
    ? s.sectionStartedAt + (s.sections[s.activeSection].minutes || 0) * MINUTE
    : s.startedAt + s.config.overallMinutes * MINUTE;
}
/** Length of the governing clock, in ms (used for the low-time warning). */
export function clockLength(s: ExamSession): number {
  return s.config.mode === 'section' ? (s.sections[s.activeSection].minutes || 0) * MINUTE : s.config.overallMinutes * MINUTE;
}
/** Time left on the whole exam. In section mode: the active section plus the full length of the sections still to come. */
export function totalLeftMs(s: ExamSession, now: number): number {
  const own = Math.max(0, deadline(s) - now);
  if (s.config.mode !== 'section') return own;
  return own + s.sections.slice(s.activeSection + 1).reduce((n, x) => n + (x.minutes || 0) * MINUTE, 0);
}

export function closeQuestion(s: ExamSession, at: number): ExamSession {
  const dt = Math.max(0, (at - s.qStartedAt) / 1000);
  return { ...s, answers: s.answers.map((a, i) => (i === s.current ? { ...a, seconds: a.seconds + dt } : a)), qStartedAt: Math.max(at, s.qStartedAt) };
}

function enterNextSection(s: ExamSession, at: number): ExamSession {
  const c = closeQuestion(s, at);
  const next = c.activeSection + 1;
  return { ...c, activeSection: next, sectionStartedAt: at, current: c.sections[next].start, qStartedAt: at };
}

/** Advances the exam to `now`: locks expired sections (each next section starts exactly when the previous one ended) and detects final expiry. */
export function tick(s: ExamSession, now: number): { session: ExamSession; finished: boolean; finishedAt: number } {
  let cur = s;
  for (;;) {
    const d = deadline(cur);
    if (now < d) return { session: cur, finished: false, finishedAt: now };
    if (cur.config.mode !== 'section' || cur.activeSection >= cur.sections.length - 1) return { session: closeQuestion(cur, d), finished: true, finishedAt: d };
    cur = enterNextSection(cur, d);
  }
}

/** Used after a page reload: catch up with the clock, then restart question timing from now. */
export function resumeSession(s: ExamSession, now: number) {
  const t = tick(s, now);
  return t.finished ? t : { ...t, session: { ...t.session, qStartedAt: now } };
}

export function goTo(s: ExamSession, qi: number, now: number): ExamSession {
  if (!canAccess(s, qi)) return s;
  const c = closeQuestion(s, now);
  const sectionPos = [...c.sectionPos];
  sectionPos[sectionOf(c, qi)] = qi;
  return { ...c, current: qi, sectionPos };
}
/** Overall mode only: tabs jump back to where the user last was in that section. */
export function goToSection(s: ExamSession, si: number, now: number): ExamSession {
  return s.config.mode === 'section' || !s.sections[si] ? s : goTo(s, s.sectionPos[si], now);
}
export function canLeaveSection(s: ExamSession) {
  return s.config.mode === 'section' && s.config.allowEarlyNext && s.activeSection < s.sections.length - 1;
}
/** Section mode: leave the active section now. The unused time is forfeited and the next section's clock starts. */
export function leaveSection(s: ExamSession, now: number): ExamSession {
  return canLeaveSection(s) ? enterNextSection(s, now) : s;
}
const patch = (s: ExamSession, fn: (a: AnswerState) => AnswerState): ExamSession =>
  canAccess(s, s.current) ? { ...s, answers: s.answers.map((a, i) => (i === s.current ? fn(a) : a)) } : s;
export const selectOption = (s: ExamSession, opt: number) => (opt >= 0 && opt < (s.questions[s.current]?.options.length ?? 0) ? patch(s, a => ({ ...a, selected: opt })) : s);
export const clearAnswer = (s: ExamSession) => patch(s, a => ({ ...a, selected: null }));
export const toggleMark = (s: ExamSession) => patch(s, a => ({ ...a, marked: !a.marked }));

export function sectionStats(s: ExamSession, si: number) {
  const sec = s.sections[si];
  const rows = s.answers.slice(sec.start, sec.end);
  return { total: rows.length, answered: rows.filter(a => a.selected !== null).length, marked: rows.filter(a => a.marked).length };
}

/* ---------------------------------------------------------------- Scoring */

export function scoreSession(s: ExamSession, finishedAt: number, auto: boolean): TestResult {
  const { positive, negative } = s.config;
  const review: ReviewItem[] = s.questions.map((q, i) => ({ q, selected: s.answers[i].selected, correctIndex: q.correctIndex, seconds: Math.round(s.answers[i].seconds) }));
  const tally = (rows: ReviewItem[]) => {
    const correct = rows.filter(r => r.selected !== null && r.selected === r.correctIndex).length;
    const wrong = rows.filter(r => r.selected !== null && r.selected !== r.correctIndex).length;
    return { correct, wrong, attempted: correct + wrong, skipped: rows.length - correct - wrong };
  };
  const subjects: SubjectResult[] = s.sections.map(sec => {
    const rows = review.slice(sec.start, sec.end);
    const t = tally(rows);
    return { name: sec.subject, total: rows.length, attempted: t.attempted, correct: t.correct, wrong: t.wrong, skipped: t.skipped, score: round3(t.correct * positive - t.wrong * negative), accuracy: t.attempted ? round3((t.correct / t.attempted) * 100) : 0, seconds: rows.reduce((n, r) => n + r.seconds, 0), timeLimit: sec.minutes };
  });
  const t = tally(review);
  return {
    id: uid(), date: new Date(finishedAt).toISOString(), total: review.length, attempted: t.attempted, correct: t.correct, wrong: t.wrong, skipped: t.skipped,
    score: round3(t.correct * positive - t.wrong * negative), maxScore: round3(review.length * positive), duration: Math.max(0, Math.round((finishedAt - s.startedAt) / 1000)),
    subjects, review, positive, negative, negativeMarks: round3(t.wrong * negative), accuracy: t.attempted ? round3((t.correct / t.attempted) * 100) : 0,
    mode: s.config.mode, autoSubmitted: auto, fileName: s.fileName, config: s.config,
  };
}

/** Full question-by-question reviews are kept for the most recent results only; older entries keep their summary. */
export const limitReviews = (history: TestResult[]) => history.slice(0, HISTORY_KEEP).map((h, i) => (i < REVIEW_KEEP || !h.review.length ? h : { ...h, review: [] }));

/* ---------------------------------------------------------------- Browser storage */

export function loadStore(): StoreData | null {
  try {
    const raw = localStorage.getItem(STORE_KEY);
    if (!raw) return null;
    const d = JSON.parse(raw);
    const history: TestResult[] = Array.isArray(d.history) ? d.history.filter((h: TestResult) => h && Array.isArray(h.subjects)).map((h: TestResult) => ({ ...h, review: Array.isArray(h.review) ? h.review : [] })) : [];
    const questions: Question[] = Array.isArray(d.questions) ? d.questions : [];
    return { history, questions, subjects: questions.length && Array.isArray(d.subjects) ? d.subjects : groupSubjects(questions), counts: d.counts || {}, fileName: d.fileName || 'Restored question bank', settings: d.settings };
  } catch { return null; }
}
/** Returns false only if the data could not be stored at all (after dropping stored reviews to make room). */
export function saveStore(data: StoreData): boolean {
  try { localStorage.setItem(STORE_KEY, JSON.stringify(data)); return true; } catch { /* quota: retry smaller */ }
  try { localStorage.setItem(STORE_KEY, JSON.stringify({ ...data, history: data.history.map(h => ({ ...h, review: [] })) })); return true; } catch { return false; }
}
export const clearStore = () => { try { localStorage.removeItem(STORE_KEY); localStorage.removeItem(SESSION_KEY); } catch { /* ignore */ } };

export function saveSession(s: ExamSession): boolean { try { localStorage.setItem(SESSION_KEY, JSON.stringify(s)); return true; } catch { return false; } }
export function clearSession() { try { localStorage.removeItem(SESSION_KEY); } catch { /* ignore */ } }
export function loadSession(): ExamSession | null {
  try {
    const raw = localStorage.getItem(SESSION_KEY);
    if (!raw) return null;
    const s = JSON.parse(raw) as ExamSession;
    const ok = s && s.v === 1 && s.config && Array.isArray(s.questions) && s.questions.length > 0 && Array.isArray(s.answers) && s.answers.length === s.questions.length
      && Array.isArray(s.sections) && s.sections.length > 0 && s.sections[s.sections.length - 1].end === s.questions.length && Array.isArray(s.sectionPos) && s.sectionPos.length === s.sections.length
      && Number.isInteger(s.current) && s.current >= 0 && s.current < s.questions.length && Number.isInteger(s.activeSection) && s.activeSection >= 0 && s.activeSection < s.sections.length
      && [s.startedAt, s.sectionStartedAt, s.qStartedAt].every(Number.isFinite);
    if (!ok) { clearSession(); return null; }
    return s;
  } catch { return null; }
}
