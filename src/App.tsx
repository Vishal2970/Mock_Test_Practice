import { useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileSpreadsheet, Clock3, Play, CheckCircle2, Flag, ChevronLeft, ChevronRight, Send, RotateCcw, Download, Trash2, X, AlertCircle, BookOpen, BarChart3, Timer, Shuffle, Home, Keyboard, Lock, Maximize, Minimize, Eraser, TriangleAlert, Cloud, Folder, RefreshCw, LogOut } from 'lucide-react';
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid } from 'recharts';
import * as E from './examEngine';
import type { Question, Subject, ExamSession, ExamConfig, TestResult, RawSheet, ImportReport, AnswerState, TimingMode } from './examEngine';
import { downloadDriveFile, DriveApiError, GoogleAuthCancelledError, getDriveFile, getDriveFolder, listDriveItems, prepareGoogleIdentityServices, requestDriveAccessToken, revokeDriveAccessToken } from './googleDrive';
import type { DriveItem } from './googleDrive';

type Phase = 'setup' | 'exam' | 'result';
type Confirm = null | 'section' | 'final';
type DriveSyncFile = DriveItem;
const GOOGLE_CLIENT_ID = import.meta.env.VITE_GOOGLE_CLIENT_ID?.trim() || '';
const DRIVE_ROOT_FOLDER_ID = import.meta.env.VITE_GOOGLE_DRIVE_ROOT_FOLDER_ID?.trim() || '';
const { fmt } = E;
const letter = (i: number) => String.fromCharCode(65 + i);
const palState = (a: AnswerState) => (a.selected !== null ? (a.marked ? 'answered-marked' : 'answered') : a.marked ? 'marked' : 'unanswered');
const enterFs = () => { try { document.documentElement.requestFullscreen?.()?.catch(() => {}); } catch { /* not supported */ } };
const leaveFs = () => { try { if (document.fullscreenElement) document.exitFullscreen().catch(() => {}); } catch { /* not supported */ } };

function App() {
  // Everything persisted is read once, so first render already shows the saved bank, settings and any exam in progress.
  const [boot] = useState(() => ({ store: E.loadStore(), session: E.loadSession() }));
  const st = boot.store;
  const [questions, setQuestions] = useState<Question[]>(st?.questions || []);
  const [fileName, setFileName] = useState(st?.fileName || '');
  const [subjects, setSubjects] = useState<Subject[]>(st?.subjects || []);
  const [counts, setCounts] = useState<Record<string, string>>(st?.counts || {});
  const [duration, setDuration] = useState(st?.settings?.duration || '90');
  const [positive, setPositive] = useState(st?.settings?.positive || '1');
  const [negative, setNegative] = useState(st?.settings?.negative || '0.25');
  const [timingMode, setTimingMode] = useState<TimingMode>(st?.settings?.timingMode || 'overall');
  const [sectionMinutes, setSectionMinutes] = useState<Record<string, string>>(st?.settings?.sectionMinutes || {});
  const [allowEarly, setAllowEarly] = useState(st?.settings?.allowEarlyNext ?? true);
  const [fullscreen, setFullscreen] = useState(st?.settings?.fullscreen ?? false);
  const [phase, setPhase] = useState<Phase>(boot.session ? 'exam' : 'setup');
  const [session, setSession] = useState<ExamSession | null>(boot.session);
  const sessionRef = useRef<ExamSession | null>(boot.session);
  const [now, setNow] = useState(() => Date.now());
  const [result, setResult] = useState<TestResult | null>(null);
  const [resultNote, setResultNote] = useState('');
  const [history, setHistory] = useState<TestResult[]>(st?.history || []);
  const [confirm, setConfirm] = useState<Confirm>(null);
  const confirmRef = useRef<Confirm>(null);
  confirmRef.current = confirm;
  const [notice, setNotice] = useState('');
  const [toast, setToast] = useState('');
  const [importing, setImporting] = useState(false);
  const [report, setReport] = useState<ImportReport | null>(null);
  const [storageWarn, setStorageWarn] = useState(false);
  const [sessionWarn, setSessionWarn] = useState(false);
  const [isFs, setIsFs] = useState(false);
  const [reviewFilter, setReviewFilter] = useState<'all' | 'wrong' | 'correct' | 'skipped'>('all');
  const [reviewSubject, setReviewSubject] = useState('All subjects');
  const [showExplanation, setShowExplanation] = useState<Record<number, boolean>>({});
  const fileRef = useRef<HTMLInputElement>(null);
  const [driveToken, setDriveToken] = useState<string | null>(null);
  const [driveExpiresAt, setDriveExpiresAt] = useState(0);
  const [driveRootName, setDriveRootName] = useState('');
  const [driveFolderId, setDriveFolderId] = useState(DRIVE_ROOT_FOLDER_ID);
  const [drivePath, setDrivePath] = useState<DriveItem[]>([]);
  const [driveItems, setDriveItems] = useState<DriveItem[]>([]);
  const [driveBusy, setDriveBusy] = useState(false);
  const [driveAuthPreparing, setDriveAuthPreparing] = useState(false);
  const [driveMessage, setDriveMessage] = useState('');
  const [driveSyncFile, setDriveSyncFile] = useState<DriveSyncFile | null>(null);
  const driveSyncFileRef = useRef<DriveSyncFile | null>(driveSyncFile);
  driveSyncFileRef.current = driveSyncFile;

  const totalRequested = useMemo(() => Object.values(counts).reduce((n, v) => n + (Number(v) || 0), 0), [counts]);
  const totalAvailable = subjects.reduce((n, s) => n + s.available, 0);
  const minutesOf = (name: string) => Number(sectionMinutes[name]) || 0;
  const activeSubjects = subjects.filter(s => (Number(counts[s.name]) || 0) > 0);
  const filledTimes = activeSubjects.filter(s => minutesOf(s.name) > 0).length;
  const sectionTimed = timingMode === 'section' && filledTimes > 0;
  const partialTimes = sectionTimed && filledTimes < activeSubjects.length;
  const totalSectionMinutes = activeSubjects.reduce((n, s) => n + minutesOf(s.name), 0);
  const timingValid = sectionTimed ? !partialTimes : Number(duration) > 0;
  const configValid = subjects.length > 0 && totalRequested > 0 && totalRequested <= totalAvailable && subjects.every(s => (Number(counts[s.name]) || 0) <= s.available) && timingValid && Number(positive) >= 0 && Number(negative) >= 0;

  /* ------------------------------------------------------------ persistence */
  const storeRef = useRef<E.StoreData>({ history, questions, subjects, counts, fileName });
  storeRef.current = { history, questions, subjects, counts, fileName, settings: { timingMode, duration, positive, negative, sectionMinutes, allowEarlyNext: allowEarly, fullscreen } };
  useEffect(() => {
    const id = window.setTimeout(() => setStorageWarn(!E.saveStore(storeRef.current)), 400);
    return () => window.clearTimeout(id);
  }, [history, questions, subjects, counts, fileName, timingMode, duration, positive, negative, sectionMinutes, allowEarly, fullscreen]);
  useEffect(() => { if (session) setSessionWarn(!E.saveSession(session)); }, [session]);
  useEffect(() => {
    const onHide = () => { const s = sessionRef.current; if (s) E.saveSession(E.closeQuestion(s, Date.now())); E.saveStore(storeRef.current); };
    const onUnload = (e: BeforeUnloadEvent) => { if (sessionRef.current) { e.preventDefault(); e.returnValue = ''; } };
    const onFs = () => setIsFs(!!document.fullscreenElement);
    window.addEventListener('pagehide', onHide); window.addEventListener('beforeunload', onUnload); document.addEventListener('fullscreenchange', onFs);
    return () => { window.removeEventListener('pagehide', onHide); window.removeEventListener('beforeunload', onUnload); document.removeEventListener('fullscreenchange', onFs); };
  }, []);
  useEffect(() => { if (!toast) return; const id = window.setTimeout(() => setToast(''), 8000); return () => window.clearTimeout(id); }, [toast]);
  useEffect(() => {
    if (!GOOGLE_CLIENT_ID || !DRIVE_ROOT_FOLDER_ID) return;
    let active = true;
    setDriveAuthPreparing(true);
    void prepareGoogleIdentityServices().catch(error => {
      console.warn('[MockPractice Drive OAuth] Google Identity Services preload failed.', {
        message: error instanceof Error ? error.message : 'Unknown error',
      });
    }).finally(() => { if (active) setDriveAuthPreparing(false); });
    return () => { active = false; };
  }, []);
  useEffect(() => {
    if (!driveToken || !driveExpiresAt) return;
    const id = window.setTimeout(() => {
      setDriveToken(null); setDriveExpiresAt(0); setDriveItems([]); setDrivePath([]); setDriveSyncFile(null); driveSyncFileRef.current = null;
      setDriveMessage('Your Google Drive session expired. Sign in again to continue syncing.');
    }, Math.max(0, driveExpiresAt - Date.now()));
    return () => window.clearTimeout(id);
  }, [driveToken, driveExpiresAt]);
  useEffect(() => {
    if (!driveToken) return;
    let active = true;
    const sync = async () => {
      try {
        const items = await listDriveItems(driveToken, driveFolderId);
        if (active) setDriveItems(items);
        if (!driveSyncFile) return;
        const latest = await getDriveFile(driveToken, driveSyncFile.id);
        if (latest.modifiedTime && latest.modifiedTime !== driveSyncFile.modifiedTime) {
          if (!/\.(xlsx|xls)$/i.test(latest.name)) {
            if (active) { setDriveSyncFile(null); setDriveMessage('The synced Drive file is no longer an Excel workbook.'); }
            return;
          }
          const bytes = await downloadDriveFile(driveToken, latest.id);
          const currentSync = driveSyncFileRef.current;
          if (active && currentSync?.id === driveSyncFile.id && currentSync.modifiedTime === driveSyncFile.modifiedTime && processWorkbook(bytes, latest.name, latest)) {
            setDriveMessage(`Automatically synced the updated Drive workbook “${latest.name}”.`);
          }
        }
      } catch (error) {
        if (!active) return;
        if (error instanceof DriveApiError && error.status === 401) {
          setDriveToken(null); setDriveExpiresAt(0); setDriveItems([]); setDrivePath([]); setDriveSyncFile(null); driveSyncFileRef.current = null;
          setDriveMessage('Your Google Drive session expired. Sign in again to continue syncing.');
        } else if (error instanceof DriveApiError && (error.status === 403 || error.status === 404)) {
          setDriveItems([]); setDriveSyncFile(null); driveSyncFileRef.current = null;
          setDriveMessage('Access to this Drive folder or file was denied. Check its sharing permissions.');
        } else {
          setDriveMessage(`Drive synchronization failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
        }
      }
    };
    const id = window.setInterval(() => { void sync(); }, 60000);
    return () => { active = false; window.clearInterval(id); };
  }, [driveToken, driveFolderId, driveSyncFile]);

  /* ------------------------------------------------------------ exam engine glue (all state transitions go through the ref) */
  function commit(next: ExamSession | null) { sessionRef.current = next; setSession(next); }
  function noteAdvance(prev: ExamSession, next: ExamSession) {
    if (next.activeSection > prev.activeSection) setToast(`${prev.sections[prev.activeSection].subject} time is up and the section is locked. You are now in ${next.sections[next.activeSection].subject}.`);
  }
  function finish(s: ExamSession, at: number, auto: boolean) {
    if (!sessionRef.current) return;
    const r = E.scoreSession(E.closeQuestion(s, at), at, auto);
    commit(null); E.clearSession();
    setResult(r); setHistory(h => E.limitReviews([r, ...h]));
    setResultNote(auto ? 'Time is up. Your test was submitted automatically.' : '');
    setReviewFilter('all'); setReviewSubject('All subjects'); setShowExplanation({}); setConfirm(null); setToast(''); setPhase('result'); leaveFs();
  }
  /** Runs one user action against the latest session, first catching up with the clock so a locked section can never be edited. */
  function act(fn: (s: ExamSession, at: number) => ExamSession) {
    const s = sessionRef.current; if (!s) return;
    const at = Date.now(); const t = E.tick(s, at);
    if (t.finished) { finish(t.session, t.finishedAt, true); return; }
    if (t.session !== s) noteAdvance(s, t.session);
    commit(fn(t.session, at)); setNow(at);
  }
  function submitNow() {
    const s = sessionRef.current; if (!s) return;
    const at = Date.now(); const t = E.tick(s, at);
    if (t.finished) finish(t.session, t.finishedAt, true); else finish(t.session, at, false);
  }
  function beginExam(s: ExamSession) {
    commit(s); setNow(Date.now()); setResult(null); setResultNote(''); setNotice(''); setToast(''); setConfirm(null); setPhase('exam');
    if (fullscreen) enterFs();
  }

  useEffect(() => { // an exam found in storage: catch up with the clock (may lock sections or auto-submit)
    const s = sessionRef.current; if (!s) return;
    const at = Date.now(); const r = E.resumeSession(s, at);
    if (r.finished) { finish(r.session, r.finishedAt, true); return; }
    commit(r.session); setNow(at);
    setToast(r.session.activeSection > s.activeSection ? `Exam restored. Time passed while you were away, so you are now in ${r.session.sections[r.session.activeSection].subject}.` : 'Your exam was restored after the page reloaded.');
  }, []);
  useEffect(() => {
    if (phase !== 'exam') return;
    const id = window.setInterval(() => {
      const s = sessionRef.current; if (!s) return;
      const at = Date.now(); setNow(at);
      const t = E.tick(s, at);
      if (t.finished) finish(t.session, t.finishedAt, true);
      else if (t.session !== s) { noteAdvance(s, t.session); commit(t.session); }
    }, 250);
    return () => window.clearInterval(id);
  }, [phase]);
  useEffect(() => {
    if (phase !== 'exam') return;
    const onKey = (e: KeyboardEvent) => {
      if (confirmRef.current) { if (e.key === 'Escape') setConfirm(null); return; }
      const tag = (e.target as HTMLElement | null)?.tagName;
      if (tag === 'INPUT' || tag === 'SELECT' || tag === 'TEXTAREA' || e.ctrlKey || e.metaKey || e.altKey) return;
      const k = e.key.toLowerCase();
      if (['a', 'b', 'c', 'd'].includes(k)) act(s => E.selectOption(s, k.charCodeAt(0) - 97));
      else if (e.key === 'ArrowRight') act((s, at) => E.goTo(s, s.current + 1, at));
      else if (e.key === 'ArrowLeft') act((s, at) => E.goTo(s, s.current - 1, at));
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [phase]);

  /* ------------------------------------------------------------ setup actions */
  function processWorkbook(data: ArrayBuffer, name: string, syncFile: DriveSyncFile | null = null) {
    const wb = XLSX.read(data, { type: 'array' });
    const sheets: RawSheet[] = wb.SheetNames.map((sheetName, i) => {
      const ws = wb.Sheets[sheetName];
      return { name: sheetName, hidden: (wb.Workbook?.Sheets?.[i]?.Hidden || 0) > 0, rowOffset: ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']).s.r : 0, rows: XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: '', raw: false }) };
    });
    const { questions: parsed, report: rep } = E.parseSheets(sheets);
    setReport(rep.issues.length || rep.warnings.length || rep.duplicates || rep.skippedSheets.length ? rep : null);
    if (!parsed.length) { setNotice('No valid questions found. Check the worksheet rows and columns (details below).'); return false; }
    const grouped = E.groupSubjects(parsed);
    setQuestions(parsed); setSubjects(grouped); setCounts(Object.fromEntries(grouped.map(s => [s.name, '0']))); setFileName(name);
    driveSyncFileRef.current = syncFile;
    setDriveSyncFile(syncFile);
    const extra = [rep.issues.length && `${rep.issues.length} row(s) skipped`, rep.duplicates && `${rep.duplicates} duplicate(s) removed`, rep.skippedSheets.length && `${rep.skippedSheets.length} worksheet(s) skipped`].filter(Boolean).join(', ');
    setNotice(`Imported ${parsed.length} questions from ${grouped.length} worksheet(s).` + (extra ? ` ${extra}; see import details.` : ''));
    return true;
  }
  function importFile(file: File) {
    if (!/\.(xlsx|xls)$/i.test(file.name)) { setNotice('Please choose an .xlsx or .xls file.'); return; }
    if (questions.length && !window.confirm(`Replace the current question bank with “${file.name}”?`)) return;
    if (file.size > 25 * 1024 * 1024 && !window.confirm('This workbook is very large and may take a while to read. Continue?')) return;
    setNotice(''); setDriveMessage(''); setImporting(true);
    const reader = new FileReader();
    reader.onerror = () => { setImporting(false); setNotice('Could not read this file.'); };
    reader.onload = e => {
      window.setTimeout(() => { // let the "reading" state paint before the synchronous parse
        try {
          const data = e.target?.result;
          if (!(data instanceof ArrayBuffer)) throw new Error('Workbook data was not an ArrayBuffer.');
          processWorkbook(data, file.name);
        } catch { setNotice('Could not read this workbook. Please upload a valid .xlsx or .xls file.'); }
        finally { setImporting(false); }
      }, 30);
    };
    reader.readAsArrayBuffer(file);
  }
  async function signInToDrive() {
    if (!GOOGLE_CLIENT_ID || !DRIVE_ROOT_FOLDER_ID) {
      setDriveMessage('Google Drive is not configured. Set VITE_GOOGLE_CLIENT_ID and VITE_GOOGLE_DRIVE_ROOT_FOLDER_ID; manual uploads are still available.');
      return;
    }
    setDriveBusy(true); setDriveMessage('');
    try {
      const auth = await requestDriveAccessToken(GOOGLE_CLIENT_ID);
      const root = await getDriveFolder(auth.token, DRIVE_ROOT_FOLDER_ID);
      const items = await listDriveItems(auth.token, root.id);
      setDriveToken(auth.token); setDriveExpiresAt(Date.now() + auth.expiresIn * 1000);
      setDriveRootName(root.name); setDriveFolderId(root.id); setDrivePath([]); setDriveItems(items);
      setDriveMessage('');
    } catch (error) {
      if (error instanceof GoogleAuthCancelledError) {
        setDriveMessage('Google sign-in was cancelled before Google returned an access token. No Drive session was started; if you completed authorization, try again.');
      } else if (error instanceof DriveApiError && (error.status === 403 || error.status === 404)) {
        setDriveMessage('Access denied: this Google account cannot access the configured MockPractice root folder.');
      } else if (error instanceof DriveApiError && error.status === 400) {
        setDriveMessage('Google Drive configuration error: the configured MockPractice root ID must point to a folder.');
      } else {
        setDriveMessage(`Google Drive sign-in failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
      }
    } finally { setDriveBusy(false); }
  }
  async function openDriveFolder(folder: DriveItem, path: DriveItem[]) {
    if (!driveToken) return;
    setDriveBusy(true); setDriveMessage('');
    try {
      const items = await listDriveItems(driveToken, folder.id);
      setDriveFolderId(folder.id); setDrivePath(path); setDriveItems(items);
    } catch (error) {
      setDriveItems([]);
      setDriveMessage(error instanceof DriveApiError && (error.status === 403 || error.status === 404)
        ? 'Access denied: your account cannot read this Drive folder.'
        : `Could not open this Drive folder: ${error instanceof Error ? error.message : 'Unknown error'}`);
    } finally { setDriveBusy(false); }
  }
  async function importDriveWorkbook(file: DriveItem) {
    if (!driveToken || !/\.(xlsx|xls)$/i.test(file.name)) return;
    if (questions.length && !window.confirm(`Replace the current question bank with “${file.name}”?`)) return;
    setImporting(true); setDriveBusy(true); setDriveMessage(''); setNotice('');
    try {
      const bytes = await downloadDriveFile(driveToken, file.id);
      if (processWorkbook(bytes, file.name, file)) setDriveMessage(`Imported “${file.name}”. Changes will sync automatically while you are signed in.`);
    } catch (error) {
      if (error instanceof DriveApiError && error.status === 401) {
        setDriveToken(null); setDriveExpiresAt(0); setDriveItems([]); setDrivePath([]); setDriveSyncFile(null); driveSyncFileRef.current = null;
        setDriveMessage('Your Google Drive session expired. Sign in again to continue.');
      } else if (error instanceof DriveApiError && (error.status === 403 || error.status === 404)) {
        setDriveItems(items => items.filter(item => item.id !== file.id));
        if (driveSyncFileRef.current?.id === file.id) {
          setDriveSyncFile(null); driveSyncFileRef.current = null;
        }
        setDriveMessage('Access denied: your Google account cannot read this Excel file.');
      } else {
        setDriveMessage(`Could not download this Drive file: ${error instanceof Error ? error.message : 'Unknown error'}`);
      }
    } finally { setImporting(false); setDriveBusy(false); }
  }
  async function signOutOfDrive() {
    const token = driveToken;
    setDriveToken(null); setDriveExpiresAt(0); setDriveItems([]); setDrivePath([]); setDriveSyncFile(null); driveSyncFileRef.current = null;
    if (token) {
      try {
        await revokeDriveAccessToken(token);
        setDriveMessage('Signed out of Google Drive.');
      } catch {
        setDriveMessage('Signed out of this app. Google token revocation could not be confirmed.');
      }
    }
  }
  function updateCount(name: string, value: string) {
    const numeric = value === '' ? '' : String(Math.max(0, Math.floor(Number(value) || 0)));
    setCounts(old => ({ ...old, [name]: numeric }));
  }
  function updateMinutes(name: string, value: string) {
    setSectionMinutes(old => ({ ...old, [name]: value === '' ? '' : String(Math.max(0, Number(value) || 0)) }));
  }
  function startTest() {
    if (!configValid) { setNotice('Check your subject counts, available questions, timing and marks.'); return; }
    const config: ExamConfig = {
      mode: sectionTimed ? 'section' : 'overall', overallMinutes: sectionTimed ? totalSectionMinutes : Number(duration), allowEarlyNext: allowEarly, positive: Number(positive), negative: Number(negative),
      sections: subjects.map(s => ({ subject: s.name, count: Number(counts[s.name]) || 0, minutes: sectionTimed ? minutesOf(s.name) : null })).filter(s => s.count > 0),
    };
    beginExam(E.buildExam(questions, config, fileName, Date.now()));
  }
  function retake() {
    if (!result || !result.review.length) return;
    const base: ExamConfig = result.config || { mode: 'overall', overallMinutes: Math.max(1, Number(duration) || 90), allowEarlyNext: true, positive: result.positive ?? Number(positive), negative: result.negative ?? Number(negative), sections: [] };
    const order = [...new Set(result.review.map(r => r.q.subject))];
    const groups = order.map(subject => ({ subject, questions: E.shuffle(result.review.filter(r => r.q.subject === subject).map(r => r.q)) }));
    beginExam(E.assembleExam(groups, base, result.fileName || fileName, Date.now()));
  }
  function openHistory(h: TestResult) {
    setResult(h); setResultNote(''); setReviewFilter('all'); setReviewSubject('All subjects'); setShowExplanation({}); setPhase('result');
  }
  function reset() { setPhase('setup'); setResult(null); setResultNote(''); setNotice(''); leaveFs(); }
  function clearEverything() {
    if (phase === 'exam') return;
    if (!window.confirm('Clear this browser’s saved question bank and test history?')) return;
    E.clearStore(); setQuestions([]); setSubjects([]); setCounts({}); setSectionMinutes({}); setHistory([]); setFileName(''); setReport(null); setDriveSyncFile(null); driveSyncFileRef.current = null; reset();
  }

  /* ------------------------------------------------------------ export */
  function exportReview() {
    if (!result) return;
    const pos = result.positive ?? Number(positive), neg = result.negative ?? Number(negative);
    const rows = result.review.map(({ q, selected, correctIndex, seconds }, i) => ({
      'Subject': q.subject, 'Question': q.question, 'Option A': q.options[0], 'Option B': q.options[1], 'Option C': q.options[2], 'Option D': q.options[3],
      'Your Answer': selected === null ? 'Not attempted' : q.options[selected], 'Correct Answer': q.options[correctIndex],
      'Status': selected === null ? 'Skipped' : selected === correctIndex ? 'Correct' : 'Incorrect', 'Explanation': q.explanation, 'Time Taken (seconds)': seconds,
      'Q No.': i + 1, 'Your Option': selected === null ? '' : letter(selected), 'Correct Option': letter(correctIndex), 'Marks': selected === null ? 0 : selected === correctIndex ? pos : -neg,
    }));
    const subj = result.subjects.map(s => ({ name: s.name, total: s.total, correct: s.correct, wrong: s.wrong, skipped: s.skipped, score: s.score, attempted: s.attempted ?? s.correct + s.wrong, 'accuracy %': Math.round(s.accuracy ?? (s.correct / Math.max(1, s.correct + s.wrong)) * 100), 'time spent (seconds)': s.seconds ?? '', 'time limit (min)': s.timeLimit ?? '' }));
    const acc = Math.round(result.accuracy ?? (result.correct / Math.max(1, result.attempted)) * 100);
    const summary = [['Metric', 'Value'], ['Date', new Date(result.date).toLocaleString()], ['Question bank', result.fileName || fileName], ['Timing mode', result.mode === 'section' ? 'Section-wise timers' : 'Overall timer'], ['Total questions', result.total], ['Attempted', result.attempted], ['Correct', result.correct], ['Incorrect', result.wrong], ['Skipped', result.skipped], ['Accuracy % (of attempted)', acc], ['Marks per correct', pos], ['Marks lost per wrong', neg], ['Negative marks deducted', result.negativeMarks ?? result.wrong * neg], ['Total marks', result.score], ['Maximum marks', result.maxScore], ['Total time taken', fmt(result.duration)], ['Auto-submitted', result.autoSubmitted ? 'Yes' : 'No']];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Question Review');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(subj), 'Subject Analysis');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summary), 'Overall Summary');
    XLSX.writeFile(wb, `MockLab_Result_${new Date(result.date).toISOString().slice(0, 10)}.xlsx`);
  }
  function downloadTemplate() {
    const head = ['Question', 'Option A', 'Option B', 'Option C', 'Option D', 'Answer', 'Correct Answer', 'Explanation'];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([head, ['What is 2 + 2?', '3', '4', '5', '6', 'B', '4', '2 + 2 equals 4.']]), 'Mathematics');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([head, ['Example question?', 'A', 'B', 'C', 'D', 'A', 'A', 'Optional explanation.']]), 'Reasoning');
    XLSX.writeFile(wb, 'MockLab_Question_Template.xlsx');
  }

  const visibleReview = result?.review.map((r, i) => ({ ...r, index: i })).filter(r => (reviewSubject === 'All subjects' || r.q.subject === reviewSubject) && (reviewFilter === 'all' || (reviewFilter === 'wrong' && r.selected !== null && r.selected !== r.correctIndex) || (reviewFilter === 'correct' && r.selected === r.correctIndex) || (reviewFilter === 'skipped' && r.selected === null))) || [];

  /* ------------------------------------------------------------ exam screen */
  function renderExam(s: ExamSession) {
    const sectionMode = s.config.mode === 'section';
    const view = E.sectionOf(s, s.current), vs = s.sections[view];
    const left = Math.max(0, E.deadline(s) - now), len = E.clockLength(s), warn = left < Math.min(300000, len * 0.2);
    const totalLeft = E.totalLeftMs(s, now);
    const q = s.questions[s.current], a = s.answers[s.current];
    const total = s.questions.length;
    const answeredAll = s.answers.filter(x => x.selected !== null).length;
    const markedAll = s.answers.filter(x => x.marked).length;
    const vst = E.sectionStats(s, view);
    const lastInSection = s.current === vs.end - 1, lastSection = view === s.sections.length - 1;
    const canLeave = E.canLeaveSection(s);
    let nextLabel = 'Next', nextDisabled = false, nextTitle = '';
    let nextAction = () => act((x, at) => E.goTo(x, x.current + 1, at));
    if (lastInSection) {
      if (lastSection) nextDisabled = true;
      else if (!sectionMode) nextLabel = 'Next section';
      else if (canLeave) { nextLabel = 'Submit section'; nextAction = () => setConfirm('section'); }
      else { nextLabel = 'Next section'; nextDisabled = true; nextTitle = 'The next section opens when this section’s time ends'; }
    }
    const nextSection = s.sections[s.activeSection + 1];
    return <section className="exam-wrap">
      <div className="exam-top">
        <div><div className="eyebrow">MOCK EXAM IN PROGRESS</div><h1>{sectionMode ? vs.subject : 'Test in progress'}</h1><p>{s.fileName} · {s.sections.length} section{s.sections.length > 1 ? 's' : ''} · {total} questions</p></div>
        <div className="timers">
          {sectionMode && <div className={`timer-card ${warn ? 'timer-warning' : ''}`}><Timer size={19} /><div><small>SECTION TIME LEFT</small><b>{fmt(Math.ceil(left / 1000))}</b></div></div>}
          <div className={`timer-card ${!sectionMode && warn ? 'timer-warning' : ''}`}><Clock3 size={19} /><div><small>{sectionMode ? 'TOTAL TIME LEFT (MAX)' : 'TIME REMAINING'}</small><b>{fmt(Math.ceil(totalLeft / 1000))}</b></div></div>
          <button className="icon-button" title={isFs ? 'Exit fullscreen' : 'Fullscreen'} aria-label={isFs ? 'Exit fullscreen' : 'Enter fullscreen'} onClick={() => (isFs ? leaveFs() : enterFs())}>{isFs ? <Minimize size={17} /> : <Maximize size={17} />}</button>
        </div>
      </div>
      <nav className="section-tabs" aria-label="Sections">{s.sections.map((sec, i) => {
        const ss = E.sectionStats(s, i), locked = sectionMode && i < s.activeSection, upcoming = sectionMode && i > s.activeSection;
        return <button key={sec.subject} className={`section-tab ${i === view ? 'active' : ''} ${locked ? 'locked' : ''}`} disabled={locked || upcoming} aria-current={i === view ? 'true' : undefined}
          title={locked ? 'Time is over or the section was submitted' : upcoming ? 'Opens after the current section' : sec.subject} onClick={() => act((x, at) => E.goToSection(x, i, at))}>
          <span className="section-tab-name">{locked && <Lock size={13} />}{sec.subject}</span>
          <small>{ss.answered}/{ss.total} answered · Q{sec.start + 1}–{sec.end}{sectionMode && sec.minutes ? ` · ${sec.minutes} min` : ''}</small>
        </button>;
      })}</nav>
      {toast && <div className="exam-toast" role="status"><AlertCircle size={16} /><span>{toast}</span><button aria-label="Dismiss" onClick={() => setToast('')}><X size={14} /></button></div>}
      {sessionWarn && <div className="notice warn-box"><TriangleAlert size={16} />This browser could not save your progress. Do not refresh or close this page.</div>}
      <div className="exam-progress"><div><span>{answeredAll} of {total} answered</span><span>{Math.round(answeredAll / total * 100)}% complete</span></div><div className="progress-track"><i style={{ width: `${answeredAll / total * 100}%` }} /></div></div>
      <div className="exam-layout">
        <section className="panel question-panel">
          <div className="question-meta"><span>{vs.subject}</span><span>Question {s.current + 1} <small>· {s.current - vs.start + 1} of {vs.end - vs.start} in this section</small></span></div>
          <div className="question-text">{q.question}</div>
          <div className="options">{q.options.map((o, i) => <button key={i} className={`option ${a.selected === i ? 'selected' : ''}`} aria-pressed={a.selected === i} onClick={() => act(x => E.selectOption(x, i))}><span className="option-letter">{letter(i)}</span><span>{o}</span>{a.selected === i && <CheckCircle2 size={18} />}</button>)}</div>
          <div className="question-controls">
            <button className="secondary-button" disabled={s.current === (sectionMode ? vs.start : 0)} onClick={() => act((x, at) => E.goTo(x, x.current - 1, at))}><ChevronLeft size={17} /> Previous</button>
            <button className="secondary-button" disabled={a.selected === null} onClick={() => act(x => E.clearAnswer(x))}><Eraser size={15} /> Clear</button>
            <button className={`mark-button ${a.marked ? 'is-marked' : ''}`} onClick={() => act(x => E.toggleMark(x))}><Flag size={16} />{a.marked ? 'Marked for review' : 'Mark for review'}</button>
            <button className="primary-button next-button" title={nextTitle} disabled={nextDisabled} onClick={nextAction}>{nextLabel} <ChevronRight size={17} /></button>
          </div>
        </section>
        <aside className="panel palette-panel">
          <div className="palette-head"><div><h3>{vs.subject}</h3><p>Questions {vs.start + 1}–{vs.end}</p></div><span>{vs.end - vs.start}</span></div>
          <div className="palette-legend"><span><i className="legend-answered" /> Answered</span><span><i className="legend-unanswered" /> Not answered</span><span><i className="legend-marked" /> For review</span><span><i className="legend-both" /> Answered + review</span></div>
          <div className="palette-grid">{s.answers.slice(vs.start, vs.end).map((x, k) => { const i = vs.start + k; return <button key={s.questions[i].id} aria-label={`Question ${i + 1}, ${palState(x)}`} className={`palette-number ${palState(x)} ${s.current === i ? 'current' : ''}`} onClick={() => act((y, at) => E.goTo(y, i, at))}>{i + 1}</button>; })}</div>
          <div className="palette-summary"><div><span>Answered</span><b>{vst.answered}</b></div><div><span>Not answered</span><b>{vst.total - vst.answered}</b></div><div><span>Marked</span><b>{vst.marked}</b></div></div>
          {canLeave && <button className="secondary-button section-submit" onClick={() => setConfirm('section')}><ChevronRight size={16} /> Submit section</button>}
          <button className="submit-button" onClick={() => setConfirm('final')}><Send size={16} /> Submit test</button>
          <div className="keyboard-hint"><Keyboard size={14} /> A–D to answer · ← → to navigate</div>
        </aside>
      </div>
      {confirm === 'section' && nextSection && <div className="modal-backdrop"><div className="modal" role="dialog" aria-modal="true" aria-labelledby="mdl-title"><button className="modal-close" aria-label="Close" onClick={() => setConfirm(null)}><X size={18} /></button><div className="modal-icon"><Lock size={22} /></div>
        <h2 id="mdl-title">Submit {vs.subject}?</h2><p>You cannot return to this section. The remaining {fmt(Math.ceil(left / 1000))} is forfeited and {nextSection.subject} starts right away{nextSection.minutes ? ` with ${nextSection.minutes} minutes` : ''}.</p>
        <div className="modal-stats"><span>Answered <b>{vst.answered}</b></span><span>Unanswered <b>{vst.total - vst.answered}</b></span><span>Marked <b>{vst.marked}</b></span></div>
        <div className="modal-actions"><button className="secondary-button" autoFocus onClick={() => setConfirm(null)}>Stay in section</button><button className="primary-button" onClick={() => { setConfirm(null); act((x, at) => E.leaveSection(x, at)); }}>Submit &amp; continue</button></div></div></div>}
      {confirm === 'final' && <div className="modal-backdrop"><div className="modal" role="dialog" aria-modal="true" aria-labelledby="mdl-title"><button className="modal-close" aria-label="Close" onClick={() => setConfirm(null)}><X size={18} /></button><div className="modal-icon"><Send size={22} /></div>
        <h2 id="mdl-title">Submit your test?</h2><p>Once submitted, you cannot change your answers.{sectionMode && !lastSection ? ' Sections you have not opened yet will count as skipped.' : ''}</p>
        <div className="modal-stats"><span>Answered <b>{answeredAll}</b></span><span>Unanswered <b>{total - answeredAll}</b></span><span>Marked <b>{markedAll}</b></span></div>
        {s.sections.length > 1 && <table className="modal-table"><tbody>{s.sections.map((sec, i) => { const ss = E.sectionStats(s, i); return <tr key={sec.subject}><td>{sec.subject}</td><td>{ss.answered}/{ss.total} answered</td><td>{ss.marked} marked</td></tr>; })}</tbody></table>}
        <div className="modal-actions"><button className="secondary-button" autoFocus onClick={() => setConfirm(null)}>Continue test</button><button className="primary-button" onClick={submitNow}>Submit now</button></div></div></div>}
    </section>;
  }

  /* ------------------------------------------------------------ render */
  const pos = result?.positive ?? Number(positive), neg = result?.negative ?? Number(negative);
  const resAcc = result ? Math.round(result.accuracy ?? (result.correct / Math.max(1, result.attempted)) * 100) : 0;
  return <div className="app-shell">
    <header className="topbar"><div className="brand"><div className="brand-icon"><BookOpen size={20} /></div><div><b>MockLab</b><span>Excel-powered mock tests</span></div></div><div className="top-actions"><span className="local-pill"><span /> Stored on this device</span><button className="icon-button" disabled={phase === 'exam'} title="Clear local question bank and history" aria-label="Clear saved data" onClick={clearEverything}><Trash2 size={17} /></button></div></header>
    <main className="main">
      {storageWarn && <div className="notice warn-box"><TriangleAlert size={16} />Browser storage is full or blocked, so your latest data is not being saved. Export your results and clear old history.</div>}
      {phase === 'setup' && <>
        <section className="welcome"><div><div className="eyebrow">YOUR PERSONAL EXAM ROOM</div><h1>Prepare smarter.<br /><em>Perform better.</em></h1><p>Turn your Excel question bank into a timed, section-wise, shuffled mock test — no account, no server, no database.</p></div><div className="hero-art"><div className="hero-sheet"><FileSpreadsheet size={36} /><div className="sheet-lines"><i /><i /><i /><i /></div><b> XLSX</b></div><div className="float-check"><CheckCircle2 size={21} /></div><div className="float-clock"><Clock3 size={22} /></div></div></section>
        <div className="setup-grid">
          <section className="panel upload-panel"><div className="section-heading"><div className="step">01</div><div><h2>Choose a question-bank source</h2><p>Google Drive is optional; manual upload never requires login.</p></div></div>
            <div className="source-options">
              <div className="source-option"><div className="source-option-title"><Cloud size={18} /><b>Google Drive</b><span>Optional</span></div><p>Browse shared exam folders and sync Excel changes while signed in.</p><button className="secondary-button" disabled={driveBusy || (!driveToken && driveAuthPreparing)} onClick={driveToken ? () => { void signOutOfDrive(); } : () => { void signInToDrive(); }}>{driveToken ? <><LogOut size={16} /> Sign out of Google</> : driveAuthPreparing ? <><RefreshCw size={16} /> Preparing Google sign-in…</> : <><Cloud size={16} /> Sign in with Google</>}</button></div>
              <div className="source-option"><div className="source-option-title"><Upload size={18} /><b>Upload Excel</b><span>No login</span></div><p>Continue without signing into Google. Your workbook stays in this browser.</p><button className="secondary-button" onClick={() => fileRef.current?.click()}><Upload size={16} /> Upload Excel</button></div>
            </div>
            {driveMessage && <div className="notice drive-notice" role="status"><AlertCircle size={16} />{driveMessage}</div>}
            {driveToken && <section className="drive-browser" aria-label="Google Drive files">
              <div className="drive-browser-heading"><div><h3><Cloud size={17} /> {driveRootName || 'Google Drive'}</h3><p>Drive access is read-only. File access is checked by Google for this account.</p></div><span className="sync-label"><RefreshCw size={13} /> Auto-sync every 60s</span></div>
              <nav className="drive-breadcrumbs" aria-label="Drive folders"><button onClick={() => { void openDriveFolder({ id: DRIVE_ROOT_FOLDER_ID, name: driveRootName, mimeType: 'application/vnd.google-apps.folder' }, []); }}>{driveRootName || 'Root folder'}</button>{drivePath.map((folder, index) => <span key={folder.id}><ChevronRight size={13} /><button onClick={() => { void openDriveFolder(folder, drivePath.slice(0, index + 1)); }}>{folder.name}</button></span>)}</nav>
              {driveBusy && <div className="drive-loading"><RefreshCw size={15} /> Loading Drive…</div>}
              <div className="drive-items">
                {driveItems.filter(item => item.mimeType === 'application/vnd.google-apps.folder').map(folder => <button className="drive-item" key={folder.id} disabled={driveBusy} onClick={() => { void openDriveFolder(folder, [...drivePath, folder]); }}><Folder size={17} /><span>{folder.name}</span><ChevronRight size={15} /></button>)}
                {driveItems.filter(item => /\.(xlsx|xls)$/i.test(item.name)).map(file => <button className="drive-item drive-file" key={file.id} disabled={driveBusy || importing} onClick={() => { void importDriveWorkbook(file); }}><FileSpreadsheet size={17} /><span>{file.name}</span><small>{file.modifiedTime ? new Date(file.modifiedTime).toLocaleDateString() : 'Excel workbook'}</small><Download size={15} /></button>)}
                {!driveBusy && driveItems.filter(item => item.mimeType === 'application/vnd.google-apps.folder' || /\.(xlsx|xls)$/i.test(item.name)).length === 0 && <p className="drive-empty">No exam folders or Excel files in this folder.</p>}
              </div>
              {driveSyncFile && <p className="drive-sync-file">Automatically syncing: {driveSyncFile.name}</p>}
            </section>}
            <div className="manual-upload-heading"><div className="step">02</div><div><h2>Upload Excel — Continue without login</h2><p>Choose an Excel workbook (.xlsx or .xls)</p></div></div>
            <div className={`dropzone ${questions.length ? 'has-file' : ''}`} onDragOver={e => e.preventDefault()} onDrop={e => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) importFile(f); }} onClick={() => fileRef.current?.click()}>
              <input ref={fileRef} type="file" accept=".xlsx,.xls" hidden onChange={e => { const f = e.target.files?.[0]; if (f) importFile(f); e.currentTarget.value = ''; }} />
              {questions.length ? <><div className="file-badge"><FileSpreadsheet size={25} /></div><b>{fileName}</b><span>{questions.length} questions imported · {subjects.length} subjects detected</span><button className="text-button" onClick={e => { e.stopPropagation(); fileRef.current?.click(); }}>Replace workbook</button></> : <><div className="upload-circle"><Upload size={23} /></div><b>Drop your Excel file here</b><span>or click to browse from your device</span><small>Subject names are read from worksheet tabs</small></>}
            </div>
            <div className="format-note"><AlertCircle size={15} /><span>Columns: Question · Option A · B · C · D · Answer · Correct Answer · Explanation (optional)</span></div><button className="template-link" onClick={downloadTemplate}><Download size={15} /> Download sample Excel template</button>
            {importing && <div className="notice info"><AlertCircle size={16} />Reading workbook…</div>}
            {notice && <div className="notice"><AlertCircle size={16} />{notice}</div>}
            {report && <details className="import-report"><summary>Import details ({report.issues.length} skipped row(s), {report.duplicates} duplicate(s), {report.warnings.length} warning(s))</summary>
              {report.skippedSheets.length > 0 && <><h4>Skipped worksheets</h4><ul>{report.skippedSheets.map(x => <li key={x.name}>{x.name}: {x.reason}</li>)}</ul></>}
              {report.issues.length > 0 && <><h4>Skipped rows</h4><ul>{report.issues.slice(0, 100).map((x, i) => <li key={i}>{x.sheet}, row {x.row}: {x.reason}</li>)}{report.issues.length > 100 && <li>…and {report.issues.length - 100} more</li>}</ul></>}
              {report.warnings.length > 0 && <><h4>Warnings</h4><ul>{report.warnings.slice(0, 100).map((x, i) => <li key={i}>{x}</li>)}{report.warnings.length > 100 && <li>…and {report.warnings.length - 100} more</li>}</ul></>}
              {report.duplicates > 0 && <p>{report.duplicates} exact duplicate question(s) within a worksheet were ignored.</p>}
            </details>}
            {subjects.length > 0 && <div className="detected"><div className="detected-title">Detected worksheets <span>{totalAvailable} total questions</span></div>{subjects.map(s => <div className="subject-row" key={s.name}><div className="subject-dot" /><b>{s.name}</b><span>{s.available} questions</span></div>)}</div>}
          </section>
          <section className="panel config-panel"><div className="section-heading"><div className="step">03</div><div><h2>Configure your test</h2><p>Set question count and exam rules</p></div></div>
            {subjects.length === 0 ? <div className="empty-config"><div><Shuffle size={25} /></div><b>Your subjects will appear here</b><span>Upload an Excel workbook to configure the test.</span></div> : <>
              <div className="field-label">QUESTIONS PER SUBJECT <span>{timingMode === 'section' ? 'Count and time for each section' : 'Enter each count manually'}</span></div>
              <div className="count-list">{subjects.map(s => <div className="count-row" key={s.name}><div className="subject-name"><span className="subject-dot" /><span>{s.name}</span><small>Max {s.available}</small></div>
                <div className="count-inputs"><input type="number" min="0" max={s.available} aria-label={`${s.name} question count`} value={counts[s.name] ?? '0'} onChange={e => updateCount(s.name, e.target.value)} />
                  {timingMode === 'section' && <div className="input-unit minutes-unit"><input type="number" min="0" placeholder="–" aria-label={`${s.name} minutes`} disabled={!(Number(counts[s.name]) > 0)} value={sectionMinutes[s.name] ?? ''} onChange={e => updateMinutes(s.name, e.target.value)} /><small>min</small></div>}</div></div>)}</div>
              <div className="total-strip"><div><span>Total selected</span><b>{totalRequested} <small>/ {totalAvailable}</small></b></div><div className={totalRequested > 0 && totalRequested <= totalAvailable ? 'total-ok' : 'total-warn'}>{totalRequested > 0 && totalRequested <= totalAvailable ? 'Within available bank' : 'Check total count'}</div></div>
              <div className="field-label spaced">TIMING MODE <span>Optional</span></div>
              <div className="mode-switch" role="radiogroup" aria-label="Timing mode">
                <button role="radio" aria-checked={timingMode === 'overall'} className={timingMode === 'overall' ? 'active' : ''} onClick={() => setTimingMode('overall')}><b>Overall timer</b><span>One countdown; move between sections freely</span></button>
                <button role="radio" aria-checked={timingMode === 'section'} className={timingMode === 'section' ? 'active' : ''} onClick={() => setTimingMode('section')}><b>Section-wise timers</b><span>A time limit per subject; sections lock in order</span></button>
              </div>
              {timingMode === 'section' && <>
                {filledTimes === 0 && <p className="mode-hint">No section times entered, so the overall timer ({Number(duration) || 0} min) will be used.</p>}
                {partialTimes && <p className="mode-hint warn">Enter a time for every subject with questions, or clear all times to use the overall timer.</p>}
                <label className="check-row"><input type="checkbox" checked={allowEarly} onChange={e => setAllowEarly(e.target.checked)} /> Allow moving to the next section before time runs out (unused time is forfeited)</label>
              </>}
              <div className="field-label spaced">TEST SETTINGS</div>
              <div className="settings-grid"><label className="setting"><span><Clock3 size={15} /> {sectionTimed ? 'Total (sum of sections)' : timingMode === 'section' ? 'Overall (fallback)' : 'Duration'}</span><div className="input-unit"><input type="number" min="1" readOnly={sectionTimed} value={sectionTimed ? String(totalSectionMinutes) : duration} onChange={e => setDuration(e.target.value)} /><small>min</small></div></label><label className="setting"><span>+ Correct answer</span><div className="input-unit"><input type="number" min="0" step="0.25" value={positive} onChange={e => setPositive(e.target.value)} /><small>marks</small></div></label><label className="setting"><span>− Wrong answer</span><div className="input-unit"><input type="number" min="0" step="0.25" value={negative} onChange={e => setNegative(e.target.value)} /><small>marks</small></div></label></div>
              <label className="check-row"><input type="checkbox" checked={fullscreen} onChange={e => setFullscreen(e.target.checked)} /> Start in fullscreen</label>
              <div className="shuffle-note"><Shuffle size={16} /><span>Questions are shuffled within each subject; the four options are shuffled for every question.</span></div>
              <button className="primary-button start-button" disabled={!configValid} onClick={startTest}><Play size={17} fill="currentColor" /> Start mock test <span>{totalRequested || 0} questions · {sectionTimed ? totalSectionMinutes : Number(duration) || 0} min</span></button>
              {!configValid && <p className="validation">{partialTimes ? 'Enter a time for every section, or clear all section times.' : 'Select at least one question, stay within each subject’s available count and enter valid settings.'}</p>}
            </>}
          </section>
        </div>
        <section className="history-section"><div className="history-head"><div><h2>Recent test history</h2><p>Select a test to open its result</p></div><span>{history.length} tests</span></div>{history.length === 0 ? <div className="history-empty"><BarChart3 size={20} /> Your completed tests will appear here.</div> : <div className="history-list">{history.slice(0, 5).map(h => <div className="history-item" role="button" tabIndex={0} key={h.id} onClick={() => openHistory(h)} onKeyDown={e => { if (e.key === 'Enter') openHistory(h); }}><div className="history-date"><b>{new Date(h.date).toLocaleDateString()}</b><span>{new Date(h.date).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</span></div><div className="history-score"><b>{h.score.toFixed(2)} <small>/ {h.maxScore.toFixed(2)}</small></b><span>{h.correct} correct · {h.wrong} wrong · {h.skipped} skipped{h.mode === 'section' ? ' · sectional' : ''}</span></div><div className="history-accuracy">{Math.round(h.correct / Math.max(1, h.attempted) * 100)}%<small>accuracy</small></div></div>)}</div>}</section>
      </>}
      {phase === 'exam' && session && renderExam(session)}
      {phase === 'result' && result && <section className="result-wrap">
        <div className="result-heading"><div><div className="eyebrow">EXAM COMPLETED</div><h1>Your performance, <em>decoded.</em></h1><p>{new Date(result.date).toLocaleString()} · {result.total} questions · {fmt(result.duration)} taken{result.mode === 'section' ? ' · section-wise timers' : ''}</p></div><div className="result-actions"><button className="secondary-button" onClick={reset}><Home size={16} /> Dashboard</button><button className="primary-button" onClick={exportReview}><Download size={16} /> Export Excel</button></div></div>
        {resultNote && <div className="notice"><AlertCircle size={16} />{resultNote}</div>}
        <div className="score-grid"><div className="score-card main-score"><span>YOUR SCORE</span><strong>{result.score.toFixed(2)}</strong><small>out of {result.maxScore.toFixed(2)} marks</small><div className="score-track"><i style={{ width: `${Math.max(0, Math.min(100, result.score / Math.max(1, result.maxScore) * 100))}%` }} /></div></div><div className="score-card"><span>ACCURACY</span><strong>{resAcc}<small>%</small></strong><small>of attempted questions</small></div><div className="score-card"><span>ATTEMPTED</span><strong>{result.attempted}<small>/{result.total}</small></strong><small>{result.skipped} not attempted</small></div><div className="score-card"><span>CORRECT / WRONG</span><strong>{result.correct}<small> / {result.wrong}</small></strong><small>+{pos} / −{neg} marking</small></div></div>
        <div className="score-grid score-grid-2"><div className="score-card"><span>TOTAL QUESTIONS</span><strong>{result.total}</strong><small>{result.subjects.length} section(s)</small></div><div className="score-card"><span>SKIPPED</span><strong>{result.skipped}</strong><small>no marks awarded</small></div><div className="score-card"><span>NEGATIVE MARKS</span><strong>−{(result.negativeMarks ?? result.wrong * neg).toFixed(2)}</strong><small>from {result.wrong} wrong answer(s)</small></div><div className="score-card"><span>TIME TAKEN</span><strong>{fmt(result.duration)}</strong><small>mm:ss</small></div></div>
        <div className="analysis-grid"><section className="panel analysis-panel"><div className="panel-title"><div><h2>Subject-wise performance</h2><p>Answers by subject</p></div><BarChart3 size={20} /></div><div className="chart-wrap"><ResponsiveContainer width="100%" height="100%"><BarChart data={result.subjects} margin={{ top: 10, right: 10, left: -18, bottom: 4 }}><CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#edf0f5" /><XAxis dataKey="name" tick={{ fontSize: 11, fill: '#687386' }} axisLine={false} tickLine={false} /><YAxis allowDecimals={false} tick={{ fontSize: 11, fill: '#687386' }} axisLine={false} tickLine={false} /><Tooltip /><Bar dataKey="correct" name="Correct" fill="#21a179" radius={[5, 5, 0, 0]} /><Bar dataKey="wrong" name="Wrong" fill="#e66b70" radius={[5, 5, 0, 0]} /><Bar dataKey="skipped" name="Skipped" fill="#d8dee9" radius={[5, 5, 0, 0]} /></BarChart></ResponsiveContainer></div></section>
          <section className="panel subject-table-panel"><div className="panel-title"><div><h2>Subject summary</h2><p>Marks, accuracy and time per subject</p></div></div><div className="subject-table">{result.subjects.map(s => { const att = s.attempted ?? s.correct + s.wrong; const acc = Math.round(s.accuracy ?? (s.correct / Math.max(1, att)) * 100); return <div className="subject-result" key={s.name}><b>{s.name}</b><span>{s.total} questions</span><strong>{s.score.toFixed(2)} marks</strong><div><i>{att} attempted</i><i className="good">{s.correct} correct</i><i className="bad">{s.wrong} wrong</i><i>{s.skipped} skipped</i><i>{acc}% accuracy</i>{s.seconds !== undefined && <i>{fmt(s.seconds)} spent{s.timeLimit ? ` of ${s.timeLimit} min` : ''}</i>}</div></div>; })}</div></section></div>
        <section className="panel review-panel"><div className="review-header"><div><h2>Question-wise review</h2><p>Review every answer and learn from mistakes</p></div><span>{visibleReview.length} questions</span></div><div className="review-filters"><div className="filter-tabs">{(['all', 'wrong', 'correct', 'skipped'] as const).map(f => <button className={reviewFilter === f ? 'active' : ''} onClick={() => setReviewFilter(f)} key={f}>{f === 'all' ? 'All' : f === 'wrong' ? 'Incorrect' : f === 'correct' ? 'Correct' : 'Skipped'}</button>)}</div><select value={reviewSubject} onChange={e => setReviewSubject(e.target.value)} aria-label="Filter by subject"><option>All subjects</option>{result.subjects.map(s => <option key={s.name}>{s.name}</option>)}</select></div>
          {result.review.length === 0 && <p className="history-empty">The detailed review of this older test is no longer stored in the browser (only the 5 most recent tests keep it). Its summary above is unchanged.</p>}
          <div className="review-list">{visibleReview.map(({ q, selected, correctIndex, seconds, index }) => <article className="review-question" key={`${q.id}-${index}`}><div className="review-q-head"><span className="review-index">Q{index + 1}</span><span className="review-subject">{q.subject}</span><span className={`result-tag ${selected === null ? 'skipped' : selected === correctIndex ? 'correct' : 'wrong'}`}>{selected === null ? 'Skipped' : selected === correctIndex ? 'Correct' : 'Incorrect'}</span><small>{fmt(seconds)}</small></div><h3>{q.question}</h3><div className="review-options">{q.options.map((o, i) => <div className={`review-option ${i === correctIndex ? 'right-answer' : ''} ${selected === i && i !== correctIndex ? 'wrong-answer' : ''}`} key={i}><span>{letter(i)}</span>{o}{i === correctIndex && <b>Correct answer</b>}{selected === i && <small>Your answer</small>}</div>)}</div><div className="correct-answer-note"><b>Your answer:</b> {selected === null ? 'Not attempted' : `${letter(selected)}. ${q.options[selected]}`} · <b>Correct answer:</b> {letter(correctIndex)}. {q.options[correctIndex]}</div>{q.explanation && <><button className="explanation-toggle" onClick={() => setShowExplanation(v => ({ ...v, [index]: !v[index] }))}>{showExplanation[index] ? 'Hide explanation' : 'Show explanation'}</button>{showExplanation[index] && <p className="explanation">{q.explanation}</p>}</>}</article>)}</div></section>
        <div className="result-bottom"><button className="secondary-button" onClick={reset}><Home size={16} /> Back to dashboard</button>{result.review.length > 0 && <button className="secondary-button" onClick={retake}><RotateCcw size={16} /> Retake same questions</button>}<button className="primary-button" onClick={() => { setPhase('setup'); setResult(null); setResultNote(''); setNotice('Configure another test using the same question bank.'); }}><Shuffle size={16} /> Create another test</button></div>
      </section>}
    </main><footer><span>MOCKLAB</span><span>Private by design · Your question bank stays in this browser</span></footer>
  </div>;
}
export default App;
