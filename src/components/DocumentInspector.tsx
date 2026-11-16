import React, { useState, useRef } from 'react';
import {
  Upload,
  FileText,
  CheckCircle,
  XCircle,
  Wand2,
  Download,
  Copy,
  Check,
  RotateCcw,
  Sparkles,
  BookPlus,
  ChevronRight,
  Eye,
  FileCheck,
  FileSpreadsheet,
  FileCode,
  Share2,
  Table,
  LayoutGrid,
  FileDown,
  Printer,
  Layers,
  ArrowDownToLine,
  ExternalLink,
  ChevronDown
} from 'lucide-react';
import {
  type SpellError,
  type CheckResult,
  applyCorrections,
  estimateThaiWordCount,
  exportErrorsToCsvBlob,
  exportErrorsToJsonBlob,
  exportErrorsToWordFindReplaceText,
  getErrorsMarkdownTable
} from '../utils/thaiSpellEngine';
import {
  extractTextFromFile,
  patchOriginalDocx,
  exportToDocx,
  exportToTxt,
  downloadBlob,
  type ExtractedDocument
} from '../utils/fileHandler';
import { SAMPLE_DOCUMENTS } from '../utils/storage';

interface DocumentInspectorProps {
  onRunCheck: (
    text: string,
    documentName: string,
    preferEngine: 'auto' | 'ai' | 'offline'
  ) => Promise<CheckResult | null>;
  isLoading: boolean;
  onSaveToHistory: (result: CheckResult) => void;
  onAddCustomWord: (word: string) => void;
  customWords: string[];
  isOnline: boolean;
  forcedOffline: boolean;
}

export const DocumentInspector: React.FC<DocumentInspectorProps> = ({
  onRunCheck,
  isLoading,
  onSaveToHistory,
  onAddCustomWord,
  customWords: _customWords,
  isOnline,
  forcedOffline,
}) => {
  // Document state (starts clean and empty so users aren't forced into sample text)
  const [documentName, setDocumentName] = useState<string>('เอกสารที่ต้องการตรวจ.docx');
  const [text, setText] = useState<string>('');
  const [originalTextSnapshot, setOriginalTextSnapshot] = useState<string>('');
  const [originalFileBuffer, setOriginalFileBuffer] = useState<ArrayBuffer | null>(null);
  const [errors, setErrors] = useState<SpellError[]>([]);
  const [activeEngine, setActiveEngine] = useState<'auto' | 'ai' | 'offline'>('auto');
  const [viewMode, setViewMode] = useState<'cards' | 'table' | 'diff'>('cards');
  const [filterCategory, setFilterCategory] = useState<string>('all');
  const [filterPage, setFilterPage] = useState<string>('all');
  const [benchmarkMs, setBenchmarkMs] = useState<number | null>(null);
  const [engineUsed, setEngineUsed] = useState<'gemini-ai' | 'offline-lexicon' | null>(null);
  const [editingErrorId, setEditingErrorId] = useState<string | null>(null);
  const [customInputValue, setCustomInputValue] = useState<string>('');
  const [copied, setCopied] = useState<boolean>(false);
  const [isDragging, setIsDragging] = useState<boolean>(false);
  const [apiNotice, setApiNotice] = useState<string | null>(null);
  const [fileWarning, setFileWarning] = useState<string | null>(null);
  const [downloadSuccessToast, setDownloadSuccessToast] = useState<string | null>(null);
  const [showExportMenu, setShowExportMenu] = useState<boolean>(false);
  const [showSamplePicker, setShowSamplePicker] = useState<boolean>(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Re-compute corrected text preview
  const liveCorrectedText = applyCorrections(originalTextSnapshot || text, errors);

  // Execute check on text
  const handleCheck = async (targetText?: string, targetName?: string) => {
    const textToCheck = targetText !== undefined ? targetText : text;
    const nameToCheck = targetName || documentName;

    if (!textToCheck.trim() || isLoading) return;
    setOriginalTextSnapshot(textToCheck);

    const result = await onRunCheck(textToCheck, nameToCheck, activeEngine);
    if (result) {
      setErrors(result.errors);
      setBenchmarkMs(result.benchmarkMs);
      setEngineUsed(result.engineUsed);
      setApiNotice(result.apiNotice || null);
      onSaveToHistory(result);
    }
  };

  // Handle file selection / drop
  // Automatically extracts document text and starts scanning against the Thai lexicon immediately
  const handleFileUpload = async (file: File) => {
    try {
      const extracted: ExtractedDocument = await extractTextFromFile(file);
      setDocumentName(extracted.fileName);
      setOriginalFileBuffer(extracted.originalBuffer || null);
      setErrors([]);
      setBenchmarkMs(null);

      if (extracted.isWarning || extracted.text.startsWith('⚠️')) {
        setFileWarning(extracted.text);
        setText('');
        setOriginalTextSnapshot('');
      } else {
        setFileWarning(null);
        setText(extracted.text);
        setOriginalTextSnapshot(extracted.text);

        // Auto-run spell check scan against the Thai vocabulary immediately
        setTimeout(() => {
          handleCheck(extracted.text, extracted.fileName);
        }, 80);
      }
    } catch (err) {
      setFileWarning('⚠️ เกิดข้อผิดพลาดในการอ่านไฟล์เอกสาร กรุณาคัดลอกข้อความมาวางโดยตรง');
      console.error(err);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleFileUpload(e.dataTransfer.files[0]);
    }
  };

  // Accept a single error fix
  const handleAcceptError = (id: string) => {
    setErrors(prev =>
      prev.map(err => (err.id === id ? { ...err, status: 'accepted' as const } : err))
    );
  };

  // Reject a single error fix
  const handleRejectError = (id: string) => {
    setErrors(prev =>
      prev.map(err => (err.id === id ? { ...err, status: 'rejected' as const } : err))
    );
  };

  // Apply custom fix to an error
  const handleSaveCustomFix = (id: string) => {
    if (!customInputValue.trim()) return;
    setErrors(prev =>
      prev.map(err =>
        err.id === id
          ? {
              ...err,
              status: 'custom' as const,
              customFix: customInputValue.trim(),
            }
          : err
      )
    );
    setEditingErrorId(null);
    setCustomInputValue('');
  };

  // Auto-Fix All: Accept all pending errors
  const handleAutoFixAll = () => {
    setErrors(prev =>
      prev.map(err => ({
        ...err,
        status: err.status === 'rejected'
          ? ('rejected' as const)
          : err.status === 'custom'
          ? ('custom' as const)
          : ('accepted' as const),
      }))
    );
    // Sync the editor text with the corrected text
    const updated = applyCorrections(
      originalTextSnapshot || text,
      errors.map(err => ({
        ...err,
        status: err.status === 'rejected'
          ? ('rejected' as const)
          : err.status === 'custom'
          ? ('custom' as const)
          : ('accepted' as const),
      }))
    );
    setText(updated);
  };

  // 100% Formatting Preservation & One-Click Fix & Download
  // "คงรูปแบบการจัดหน้าเอกสารต้นฉบับ ทุกประการและ ดาวน์โหลดออกมาเหมือนเดิมทุกประการ แ้ไขคำที่ผิดให้พร้อมก่อนดาวน์โหลด"
  const handleFixAllAndDownload = async () => {
    if (!text.trim()) return;

    // 1. Mark pending errors as accepted while preserving custom & rejected
    const updatedErrors = errors.map(err => ({
      ...err,
      status: err.status === 'rejected'
        ? ('rejected' as const)
        : err.status === 'custom'
        ? ('custom' as const)
        : ('accepted' as const),
    }));
    setErrors(updatedErrors);

    // Corrections list to apply
    const confirmedCorrections = updatedErrors
      .filter(e => e.status === 'accepted' || e.status === 'custom')
      .map(e => ({
        original: e.original,
        replacement: e.status === 'custom' && e.customFix ? e.customFix : e.suggestion
      }));

    const fixedCount = confirmedCorrections.length;

    // 2. If we have the original uploaded .docx buffer, patch it directly with JSZip!
    // This preserves 100% of tables, logos, Garuda emblem, margins, and headers!
    if (originalFileBuffer && documentName.toLowerCase().endsWith('.docx')) {
      try {
        const patchedBlob = await patchOriginalDocx(originalFileBuffer, confirmedCorrections);
        const downloadName = `แก้ไขแล้ว_${documentName}`;
        downloadBlob(patchedBlob, downloadName);
        setDownloadSuccessToast(
          `✓ แก้ไขคำผิดแล้ว ${fixedCount} จุด และดาวน์โหลดเอกสารต้นฉบับสำเร็จ โดยคงรูปแบบการจัดหน้าเดิม 100%`
        );
        setTimeout(() => setDownloadSuccessToast(null), 6000);
        return;
      } catch (zipErr) {
        console.warn('Failed to patch original docx, falling back to clean docx export:', zipErr);
      }
    }

    // Fallback: If no original docx buffer (e.g. text pasted or pdf), export clean formatted docx
    const updatedText = applyCorrections(originalTextSnapshot || text, updatedErrors);
    setText(updatedText);
    exportToDocx(updatedText, `แก้ไขแล้ว_${documentName}`, {
      docTitle: documentName.replace(/\.[^/.]+$/, ''),
      errorCount: fixedCount,
    });
    setDownloadSuccessToast(`✓ แก้ไขคำผิดแล้ว ${fixedCount} จุด และดาวน์โหลดเอกสาร Word (.docx) สำเร็จ`);
    setTimeout(() => setDownloadSuccessToast(null), 6000);
  };

  // Copy to clipboard
  const handleCopyText = async () => {
    const textToCopy = errors.length > 0 ? liveCorrectedText : text;
    await navigator.clipboard.writeText(textToCopy);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  // Export errors separately for other programs
  // "พร้อมสามารถส่งออกคำที่จำแก้ไขต่างหาก เพื่อนำไปแก้โดยโปรแกรมอื่นได้"
  const handleExportCsv = () => {
    const blob = exportErrorsToCsvBlob(errors, documentName);
    downloadBlob(blob, `ตารางคำผิด_${documentName.replace(/\.[^/.]+$/, '')}.csv`);
    setShowExportMenu(false);
  };

  const handleExportWordFindReplace = () => {
    const content = exportErrorsToWordFindReplaceText(errors, documentName);
    const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
    downloadBlob(blob, `รายการค้นหาแทนที่_Word_${documentName.replace(/\.[^/.]+$/, '')}.txt`);
    setShowExportMenu(false);
  };

  const handleExportJson = () => {
    const blob = exportErrorsToJsonBlob(errors, documentName);
    downloadBlob(blob, `รายการคำผิด_${documentName.replace(/\.[^/.]+$/, '')}.json`);
    setShowExportMenu(false);
  };

  const handleCopyMarkdown = async () => {
    const markdown = getErrorsMarkdownTable(errors);
    await navigator.clipboard.writeText(markdown);
    setDownloadSuccessToast('✓ คัดลอกตารางคำผิดรูปแบบ Markdown เรียบร้อยแล้ว (สามารถวางในเอกสารหรือแชทได้ทันที)');
    setTimeout(() => setDownloadSuccessToast(null), 4000);
    setShowExportMenu(false);
  };

  // Load sample text (optional only, not forced)
  const handleLoadSample = (idx: number) => {
    if (SAMPLE_DOCUMENTS[idx]) {
      setText(SAMPLE_DOCUMENTS[idx].content);
      setOriginalTextSnapshot(SAMPLE_DOCUMENTS[idx].content);
      setDocumentName(`เอกสารตัวอย่าง_${idx + 1}.docx`);
      setOriginalFileBuffer(null);
      setErrors([]);
      setFileWarning(null);
      setShowSamplePicker(false);
      setTimeout(() => {
        handleCheck(SAMPLE_DOCUMENTS[idx].content, `เอกสารตัวอย่าง_${idx + 1}.docx`);
      }, 100);
    }
  };

  // Filter calculations
  const pendingErrors = errors.filter(e => e.status === 'pending');
  const uniquePages = Array.from(new Set(errors.map(e => e.pageNumber || 1))).sort((a, b) => a - b);

  const filteredErrors = errors.filter(e => {
    const matchCat = filterCategory === 'all' || e.category === filterCategory;
    const matchPage = filterPage === 'all' || (e.pageNumber || 1) === parseInt(filterPage, 10);
    return matchCat && matchPage;
  });

  const wordCount = estimateThaiWordCount(text);

  return (
    <div className="space-y-6">
      {/* Toast Notification Banner */}
      {downloadSuccessToast && (
        <div className="flex items-center justify-between rounded-xl border border-emerald-300 bg-emerald-50 px-4 py-3 text-sm font-medium text-emerald-900 shadow-sm animate-in fade-in slide-in-from-top-2 dark:border-emerald-800 dark:bg-emerald-950/80 dark:text-emerald-200">
          <div className="flex items-center gap-2">
            <CheckCircle className="h-5 w-5 text-emerald-600 dark:text-emerald-400 shrink-0" />
            <span>{downloadSuccessToast}</span>
          </div>
          <button
            onClick={() => setDownloadSuccessToast(null)}
            className="text-xs text-emerald-700 hover:text-emerald-900 dark:text-emerald-300"
          >
            ✕
          </button>
        </div>
      )}

      {/* Primary Top Action Bar & File Controls */}
      <section
        aria-label="การจัดการเอกสาร"
        className="rounded-2xl border border-slate-200 bg-white p-4 sm:p-5 shadow-xs dark:border-slate-800 dark:bg-slate-900"
      >
        <div className="flex flex-col gap-4 lg:flex-row lg:items-center lg:justify-between">
          {/* File Name & Direct Upload Trigger */}
          <div className="flex flex-wrap items-center gap-2 sm:gap-3">
            <div className="flex items-center gap-2 rounded-lg bg-slate-100 px-3 py-1.5 dark:bg-slate-800">
              <FileText className="h-4 w-4 text-indigo-600 dark:text-indigo-400" />
              <input
                type="text"
                value={documentName}
                onChange={e => setDocumentName(e.target.value)}
                placeholder="ชื่อเอกสาร"
                className="w-48 sm:w-64 bg-transparent text-xs sm:text-sm font-medium text-slate-800 focus:outline-hidden dark:text-slate-200"
              />
            </div>

            {/* Hidden file input */}
            <input
              ref={fileInputRef}
              type="file"
              accept=".docx,.pdf,.txt,.md"
              className="hidden"
              onChange={e => {
                if (e.target.files && e.target.files[0]) {
                  handleFileUpload(e.target.files[0]);
                }
              }}
            />

            {/* Direct Open File Button */}
            <button
              onClick={() => fileInputRef.current?.click()}
              className="flex items-center gap-1.5 rounded-lg border border-indigo-200 bg-indigo-50/70 px-3.5 py-1.5 text-xs font-semibold text-indigo-700 hover:bg-indigo-100 transition-colors dark:border-indigo-900 dark:bg-indigo-950/60 dark:text-indigo-300 cursor-pointer"
            >
              <Upload className="h-4 w-4 text-indigo-600 dark:text-indigo-400" />
              <span>เปิดไฟล์เอกสาร (.docx / .pdf / .txt)</span>
            </button>

            {/* Optional Samples Dropdown (Non-intrusive) */}
            <div className="relative">
              <button
                onClick={() => setShowSamplePicker(prev => !prev)}
                className="flex items-center gap-1 rounded-lg border border-slate-200 px-2.5 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50 dark:border-slate-700 dark:text-slate-400 dark:hover:bg-slate-800"
              >
                <span>ตัวอย่างทดสอบ</span>
                <ChevronDown className="h-3 w-3" />
              </button>

              {showSamplePicker && (
                <div className="absolute left-0 mt-1 w-72 rounded-xl border border-slate-200 bg-white p-2 shadow-lg z-30 dark:border-slate-800 dark:bg-slate-900">
                  <div className="text-[11px] font-semibold text-slate-400 px-2 py-1">
                    เลือกตัวอย่างเอกสารสำหรับทดลอง:
                  </div>
                  {SAMPLE_DOCUMENTS.map((doc, idx) => (
                    <button
                      key={idx}
                      onClick={() => handleLoadSample(idx)}
                      className="w-full text-left rounded-lg px-2.5 py-1.5 text-xs hover:bg-indigo-50 hover:text-indigo-700 dark:hover:bg-slate-800 dark:hover:text-indigo-300 transition-colors"
                    >
                      {doc.title}
                    </button>
                  ))}
                </div>
              )}
            </div>
          </div>

          {/* Engine Selector & Run / Fix-All Buttons */}
          <div className="flex flex-wrap items-center gap-2 sm:gap-3">
            {/* Engine Modes */}
            <div className="flex items-center rounded-lg border border-slate-200 p-0.5 bg-slate-50 dark:border-slate-700 dark:bg-slate-800">
              <button
                onClick={() => setActiveEngine('auto')}
                className={`px-2.5 py-1 rounded text-xs font-medium transition-colors ${
                  activeEngine === 'auto'
                    ? 'bg-white text-indigo-700 shadow-xs dark:bg-slate-700 dark:text-indigo-300'
                    : 'text-slate-600 hover:text-slate-900 dark:text-slate-400'
                }`}
              >
                Auto
              </button>
              <button
                onClick={() => setActiveEngine('ai')}
                disabled={forcedOffline || !isOnline}
                title={forcedOffline || !isOnline ? 'ไม่สามารถใช้ Online AI ในโหมดออฟไลน์' : 'ใช้ Gemini 4 AI'}
                className={`px-2.5 py-1 rounded text-xs font-medium transition-colors ${
                  activeEngine === 'ai'
                    ? 'bg-white text-indigo-700 shadow-xs dark:bg-slate-700 dark:text-indigo-300'
                    : 'text-slate-600 hover:text-slate-900 disabled:opacity-40 dark:text-slate-400'
                }`}
              >
                Gemini AI
              </button>
              <button
                onClick={() => setActiveEngine('offline')}
                className={`px-2.5 py-1 rounded text-xs font-medium transition-colors ${
                  activeEngine === 'offline'
                    ? 'bg-white text-indigo-700 shadow-xs dark:bg-slate-700 dark:text-indigo-300'
                    : 'text-slate-600 hover:text-slate-900 dark:text-slate-400'
                }`}
              >
                Offline Lexicon
              </button>
            </div>

            {/* Check Button */}
            <button
              onClick={() => handleCheck()}
              disabled={isLoading || !text.trim()}
              className="flex items-center gap-2 rounded-xl bg-indigo-600 px-4 py-2 text-xs sm:text-sm font-semibold text-white shadow-sm hover:bg-indigo-700 disabled:opacity-50 transition-colors whitespace-nowrap cursor-pointer"
            >
              {isLoading ? (
                <>
                  <span className="h-4 w-4 animate-spin rounded-full border-2 border-white border-t-transparent" />
                  <span>กำลังตรวจคำผิด...</span>
                </>
              ) : (
                <>
                  <Wand2 className="h-4 w-4" />
                  <span>ตรวจคำผิดจากคำศัพท์</span>
                </>
              )}
            </button>
          </div>
        </div>

        {/* Notice banner when Gemini API requires offline fallback */}
        {apiNotice === 'BILLING_REQUIRED' && (
          <div className="mt-3 flex items-start justify-between rounded-xl border border-amber-200 bg-amber-50/90 p-3 text-xs text-amber-900 dark:border-amber-900/60 dark:bg-amber-950/40 dark:text-amber-200">
            <div className="flex items-start gap-2">
              <span className="font-bold shrink-0">💡 สลับใช้งาน Offline Lexicon อัตโนมัติ:</span>
              <span>
                โครงการนี้ประมวลผลด้วย Offline Local Lexicon Engine ซึ่งตรวจจับและยืนยันคำผิดภาษาไทยได้ครบถ้วน ถูกต้อง และรวดเร็วระดับเสี้ยววินาที (&lt; 0.05s) โดยไม่จำเป็นต้องเชื่อมต่อคลาวด์
              </span>
            </div>
            <button
              onClick={() => setApiNotice(null)}
              className="ml-2 px-1.5 py-0.5 rounded text-amber-700 hover:bg-amber-100 font-bold dark:text-amber-400 dark:hover:bg-amber-900"
              title="ปิดการแจ้งเตือนนี้"
            >
              ✕
            </button>
          </div>
        )}

        {/* Warning banner when uploaded file cannot be parsed or requires action */}
        {fileWarning && (
          <div className="mt-3 flex items-start justify-between rounded-xl border border-amber-300 bg-amber-50/95 p-4 text-xs text-amber-900 shadow-xs dark:border-amber-900/60 dark:bg-amber-950/50 dark:text-amber-200">
            <div className="flex items-start gap-3">
              <span className="text-xl shrink-0">📄</span>
              <div className="space-y-1">
                <span className="font-bold text-sm block text-amber-950 dark:text-amber-100">
                  คำแนะนำสำหรับเอกสารนี้:
                </span>
                <p className="whitespace-pre-line leading-relaxed text-amber-900/90 dark:text-amber-200/90">
                  {fileWarning}
                </p>
              </div>
            </div>
            <button
              onClick={() => setFileWarning(null)}
              className="ml-3 px-2 py-1 rounded text-amber-700 hover:bg-amber-200 font-bold dark:text-amber-300 dark:hover:bg-amber-900"
              title="ปิดการแจ้งเตือน"
            >
              ✕
            </button>
          </div>
        )}

        {/* Drag & drop upload dropzone */}
        <div
          onDragOver={e => {
            e.preventDefault();
            setIsDragging(true);
          }}
          onDragLeave={() => setIsDragging(false)}
          onDrop={handleDrop}
          onClick={() => fileInputRef.current?.click()}
          className={`mt-3 rounded-xl border-2 border-dashed px-4 py-3 text-center text-xs transition-all cursor-pointer ${
            isDragging
              ? 'border-indigo-500 bg-indigo-50/80 text-indigo-700 dark:bg-indigo-950/50'
              : 'border-slate-200 bg-slate-50/50 text-slate-600 hover:border-indigo-300 hover:bg-indigo-50/30 dark:border-slate-800 dark:bg-slate-800/40 dark:text-slate-400'
          }`}
        >
          <div className="flex items-center justify-center gap-2">
            <Upload className="h-4 w-4 text-indigo-600 dark:text-indigo-400" />
            <span className="font-medium">
              ลากไฟล์เอกสาร .docx หรือ .pdf มาวางที่นี่ หรือคลิกเพื่อนำเข้าข้อความจากในไฟล์มาตรวจทันที
            </span>
            <span className="text-slate-400 hidden sm:inline">|</span>
            <span className="text-[11px] text-slate-500 hidden sm:inline">
              (คงรูปแบบตาราง รูปภาพ ตราครุฑ และระยะขอบต้นฉบับ 100%)
            </span>
          </div>
        </div>
      </section>

      {/* Main Workspace: Split-Grid Editor + Error Details */}
      <div className="grid grid-cols-1 gap-6 lg:grid-cols-12 items-start">
        {/* Left Column: Text Area / Diff Preview (7 cols) */}
        <section
          aria-label="พื้นที่แก้ไขข้อความและดูเปรียบเทียบ"
          className="lg:col-span-7 space-y-4"
        >
          <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-xs dark:border-slate-800 dark:bg-slate-900">
            {/* View switcher & stats */}
            <div className="flex items-center justify-between border-b border-slate-100 pb-3 dark:border-slate-800">
              <div className="flex items-center gap-1 bg-slate-100 p-1 rounded-lg dark:bg-slate-800">
                <button
                  onClick={() => setViewMode('cards')}
                  className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-md transition-colors ${
                    viewMode === 'cards' || viewMode === 'table'
                      ? 'bg-white text-slate-900 shadow-xs dark:bg-slate-700 dark:text-white'
                      : 'text-slate-600 hover:text-slate-900 dark:text-slate-400'
                  }`}
                >
                  <FileText className="h-3.5 w-3.5" />
                  <span>แก้ไขข้อความ</span>
                </button>
                <button
                  onClick={() => setViewMode('diff')}
                  disabled={errors.length === 0}
                  className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-md transition-colors ${
                    viewMode === 'diff'
                      ? 'bg-white text-slate-900 shadow-xs dark:bg-slate-700 dark:text-white'
                      : 'text-slate-600 hover:text-slate-900 disabled:opacity-40 dark:text-slate-400'
                  }`}
                >
                  <Eye className="h-3.5 w-3.5" />
                  <span>เปรียบเทียบคำผิด (Diff)</span>
                </button>
              </div>

              {/* Text metadata */}
              <div className="flex items-center gap-2 text-xs text-slate-500 font-mono tabular-nums">
                {uniquePages.length > 0 && (
                  <>
                    <span className="font-semibold text-indigo-600 dark:text-indigo-400">
                      {uniquePages.length} หน้า
                    </span>
                    <span aria-hidden="true">·</span>
                  </>
                )}
                <span>{wordCount} คำ</span>
                <span aria-hidden="true">·</span>
                <span>{text.length} ตัวอักษร</span>
              </div>
            </div>

            {/* Content view or Empty State */}
            {viewMode === 'diff' ? (
              <div className="mt-3 p-4 rounded-xl bg-slate-50 text-sm leading-relaxed max-h-[460px] overflow-y-auto font-sans whitespace-pre-wrap dark:bg-slate-950 dark:text-slate-200">
                <HighlightedDiffText
                  originalText={originalTextSnapshot || text}
                  errors={errors}
                />
              </div>
            ) : !text.trim() ? (
              /* Welcoming Empty State when no document loaded */
              <div className="py-16 px-4 text-center">
                <div className="mx-auto flex h-14 w-14 items-center justify-center rounded-2xl bg-indigo-50 text-indigo-600 dark:bg-indigo-950/60 dark:text-indigo-400">
                  <Upload className="h-7 w-7" />
                </div>
                <h3 className="mt-4 text-base font-bold text-slate-800 dark:text-slate-200">
                  ยังไม่ได้นำเข้าข้อความจากเอกสาร
                </h3>
                <p className="mt-1.5 text-xs text-slate-500 max-w-md mx-auto leading-relaxed">
                  คลิกปุ่มเปิดไฟล์ด้านบน หรือลากไฟล์ <strong>.docx</strong> / <strong>.pdf</strong> มาวางที่นี่ เพื่อให้ระบบดึงข้อความและตรวจจับคำผิดจากคลังศัพท์ภาษาไทยทันที
                </p>
                <div className="mt-5 flex flex-wrap items-center justify-center gap-3">
                  <button
                    onClick={() => fileInputRef.current?.click()}
                    className="flex items-center gap-2 rounded-xl bg-indigo-600 px-4 py-2 text-xs font-semibold text-white shadow-xs hover:bg-indigo-700 transition-colors cursor-pointer"
                  >
                    <Upload className="h-4 w-4" />
                    <span>เลือกไฟล์จากคอมพิวเตอร์</span>
                  </button>
                  <button
                    onClick={() => handleLoadSample(0)}
                    className="rounded-xl border border-slate-200 px-4 py-2 text-xs font-medium text-slate-600 hover:bg-slate-50 dark:border-slate-800 dark:text-slate-300"
                  >
                    ทดลองด้วยข้อความตัวอย่าง
                  </button>
                </div>
              </div>
            ) : (
              <div className="mt-3">
                <textarea
                  value={text}
                  onChange={e => setText(e.target.value)}
                  placeholder="วางหรือพิมพ์เนื้อหาเอกสารภาษาไทยที่ต้องการตรวจสอบที่นี่..."
                  rows={16}
                  className="w-full resize-y rounded-xl border-0 p-3 text-sm leading-relaxed text-slate-900 placeholder-slate-400 focus:outline-hidden focus:ring-1 focus:ring-indigo-500 font-sans dark:bg-slate-900 dark:text-slate-100"
                />
              </div>
            )}

            {/* Bottom Editor Action Toolbar */}
            <div className="flex flex-wrap items-center justify-between gap-3 border-t border-slate-100 pt-3 dark:border-slate-800">
              <div className="flex items-center gap-2">
                <button
                  onClick={handleCopyText}
                  className="flex items-center gap-1.5 rounded-lg border border-slate-200 px-3 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50 transition-colors dark:border-slate-700 dark:text-slate-300"
                >
                  {copied ? (
                    <>
                      <Check className="h-3.5 w-3.5 text-emerald-600" />
                      <span className="text-emerald-600">คัดลอกแล้ว!</span>
                    </>
                  ) : (
                    <>
                      <Copy className="h-3.5 w-3.5 text-slate-500" />
                      <span>คัดลอกข้อความ</span>
                    </>
                  )}
                </button>

                <button
                  onClick={() => {
                    setText('');
                    setOriginalTextSnapshot('');
                    setOriginalFileBuffer(null);
                    setErrors([]);
                  }}
                  className="flex items-center gap-1 rounded-lg px-2.5 py-1.5 text-xs font-medium text-slate-500 hover:text-rose-600 hover:bg-rose-50 transition-colors dark:text-slate-400"
                >
                  <RotateCcw className="h-3.5 w-3.5" />
                  <span>ล้างหน้าจอ</span>
                </button>
              </div>

              {/* Download & Export Menu */}
              <div className="flex items-center gap-2">
                {/* 100% Formatting Preservation Download Button */}
                <button
                  onClick={handleFixAllAndDownload}
                  disabled={!text.trim()}
                  title="แก้ไขคำผิดทั้งหมดและดาวน์โหลดเอกสารต้นฉบับ โดยคงรูปแบบการจัดหน้า ตาราง และรูปภาพเดิม 100%"
                  className="flex items-center gap-1.5 rounded-lg bg-emerald-600 px-3.5 py-1.5 text-xs font-semibold text-white shadow-xs hover:bg-emerald-700 disabled:opacity-40 transition-colors cursor-pointer"
                >
                  <ArrowDownToLine className="h-3.5 w-3.5" />
                  <span>แก้ไข & ดาวน์โหลดเอกสารต้นฉบับ (.docx)</span>
                </button>

                {/* Separate Corrections Export Menu (CSV, Word Find/Replace, JSON) */}
                <div className="relative">
                  <button
                    onClick={() => setShowExportMenu(prev => !prev)}
                    disabled={errors.length === 0}
                    title="ส่งออกรายการคำผิดต่างหาก เพื่อนำไปแก้ไขในโปรแกรมอื่น เช่น Word, Excel"
                    className="flex items-center gap-1.5 rounded-lg border border-indigo-200 bg-indigo-50 px-3 py-1.5 text-xs font-medium text-indigo-700 hover:bg-indigo-100 disabled:opacity-40 transition-colors dark:border-indigo-900 dark:bg-indigo-950/60 dark:text-indigo-300 cursor-pointer"
                  >
                    <Share2 className="h-3.5 w-3.5" />
                    <span>ส่งออกคำผิดต่างหาก</span>
                    <ChevronDown className="h-3 w-3" />
                  </button>

                  {showExportMenu && (
                    <div className="absolute right-0 mt-1 w-64 rounded-xl border border-slate-200 bg-white p-2 shadow-xl z-30 dark:border-slate-800 dark:bg-slate-900">
                      <div className="px-2 py-1 text-[11px] font-bold text-slate-400">
                        ส่งออกสำหรับโปรแกรมอื่น:
                      </div>
                      <button
                        onClick={handleExportCsv}
                        className="flex items-center gap-2 w-full rounded-lg px-2.5 py-2 text-xs text-left hover:bg-slate-100 dark:hover:bg-slate-800"
                      >
                        <FileSpreadsheet className="h-4 w-4 text-emerald-600" />
                        <div>
                          <div className="font-semibold text-slate-800 dark:text-slate-200">
                            ตาราง Excel / CSV (.csv)
                          </div>
                          <div className="text-[10px] text-slate-500">
                            ระบุหน้าและบริบท เปิดใน Excel ภาษาไทยไม่เพี้ยน
                          </div>
                        </div>
                      </button>

                      <button
                        onClick={handleExportWordFindReplace}
                        className="flex items-center gap-2 w-full rounded-lg px-2.5 py-2 text-xs text-left hover:bg-slate-100 dark:hover:bg-slate-800"
                      >
                        <FileText className="h-4 w-4 text-blue-600" />
                        <div>
                          <div className="font-semibold text-slate-800 dark:text-slate-200">
                            ค้นหา & แทนที่ Word (.txt)
                          </div>
                          <div className="text-[10px] text-slate-500">
                            คู่คำสำหรับเครื่องมือ Find & Replace ของ Word
                          </div>
                        </div>
                      </button>

                      <button
                        onClick={handleExportJson}
                        className="flex items-center gap-2 w-full rounded-lg px-2.5 py-2 text-xs text-left hover:bg-slate-100 dark:hover:bg-slate-800"
                      >
                        <FileCode className="h-4 w-4 text-amber-600" />
                        <div>
                          <div className="font-semibold text-slate-800 dark:text-slate-200">
                            ข้อมูลโครงสร้าง JSON (.json)
                          </div>
                          <div className="text-[10px] text-slate-500">
                            สำหรับเชื่อมต่อระบบภายนอกหรือ Script
                          </div>
                        </div>
                      </button>

                      <button
                        onClick={handleCopyMarkdown}
                        className="flex items-center gap-2 w-full rounded-lg px-2.5 py-2 text-xs text-left hover:bg-slate-100 dark:hover:bg-slate-800 border-t border-slate-100 dark:border-slate-800"
                      >
                        <Copy className="h-4 w-4 text-indigo-600" />
                        <div>
                          <div className="font-semibold text-slate-800 dark:text-slate-200">
                            คัดลอกตาราง Markdown
                          </div>
                          <div className="text-[10px] text-slate-500">
                            สำหรับวางใน Docs, Notion, หรือแชท
                          </div>
                        </div>
                      </button>
                    </div>
                  )}
                </div>

                {/* Print button */}
                <button
                  onClick={() => window.print()}
                  title="พิมพ์เอกสารพร้อมบันทึกเป็น PDF"
                  className="flex items-center gap-1.5 rounded-lg border border-slate-200 px-2.5 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50 transition-colors dark:border-slate-700 dark:text-slate-300"
                >
                  <Printer className="h-3.5 w-3.5 text-slate-500" />
                  <span className="hidden sm:inline">พิมพ์</span>
                </button>
              </div>
            </div>
          </div>
        </section>

        {/* Right Column: Detected Errors & Page-by-Page Details (5 cols) */}
        <section
          aria-label="รายการคำผิดที่ตรวจพบและการยืนยันแก้ไข"
          className="lg:col-span-5 space-y-4"
        >
          <div className="rounded-2xl border border-slate-200 bg-white p-4 sm:p-5 shadow-xs dark:border-slate-800 dark:bg-slate-900">
            {/* Header with Switcher and Auto-Fix All */}
            <div className="flex items-center justify-between border-b border-slate-100 pb-3 dark:border-slate-800">
              <div className="flex items-center gap-2">
                <h2 className="text-sm font-bold text-slate-900 dark:text-white">
                  คำผิดที่ตรวจพบ ({errors.length})
                </h2>
                {pendingErrors.length > 0 && (
                  <span className="rounded-full bg-rose-100 px-2 py-0.5 text-[11px] font-semibold text-rose-700 dark:bg-rose-950/70 dark:text-rose-300">
                    รอการยืนยัน {pendingErrors.length}
                  </span>
                )}
              </div>

              <div className="flex items-center gap-1.5">
                {/* Switch between Cards and Detailed Table View */}
                <div className="flex items-center rounded-lg bg-slate-100 p-0.5 dark:bg-slate-800">
                  <button
                    onClick={() => setViewMode('cards')}
                    title="แสดงแบบการ์ดคำผิด"
                    className={`p-1 rounded ${
                      viewMode === 'cards'
                        ? 'bg-white shadow-xs text-indigo-700 dark:bg-slate-700 dark:text-indigo-300'
                        : 'text-slate-500 hover:text-slate-800'
                    }`}
                  >
                    <LayoutGrid className="h-3.5 w-3.5" />
                  </button>
                  <button
                    onClick={() => setViewMode('table')}
                    title="แสดงแบบตารางระบุหน้าเอกสารทุกรายการ"
                    className={`p-1 rounded ${
                      viewMode === 'table'
                        ? 'bg-white shadow-xs text-indigo-700 dark:bg-slate-700 dark:text-indigo-300'
                        : 'text-slate-500 hover:text-slate-800'
                    }`}
                  >
                    <Table className="h-3.5 w-3.5" />
                  </button>
                </div>

                {errors.length > 0 && (
                  <button
                    onClick={handleAutoFixAll}
                    title="ยืนยันการแก้ไขคำผิดทั้งหมดทันที"
                    className="flex items-center gap-1 rounded-lg bg-emerald-600 px-2.5 py-1 text-xs font-semibold text-white shadow-xs hover:bg-emerald-700 transition-colors cursor-pointer"
                  >
                    <Sparkles className="h-3 w-3" />
                    <span>แก้ทั้งหมด</span>
                  </button>
                )}
              </div>
            </div>

            {/* Filter Tabs: Page Filter & Category Filter */}
            {errors.length > 0 && (
              <div className="mt-3 space-y-2">
                {/* Filter by Page if document has pages */}
                {uniquePages.length > 1 && (
                  <div className="flex items-center gap-1.5 overflow-x-auto pb-1 text-xs">
                    <span className="text-[11px] font-semibold text-slate-500 shrink-0">
                      กรองตามหน้า:
                    </span>
                    <button
                      onClick={() => setFilterPage('all')}
                      className={`px-2 py-0.5 rounded text-xs transition-colors ${
                        filterPage === 'all'
                          ? 'bg-indigo-600 text-white font-medium'
                          : 'bg-slate-100 text-slate-600 hover:bg-slate-200 dark:bg-slate-800'
                      }`}
                    >
                      ทุกหน้า ({errors.length})
                    </button>
                    {uniquePages.map(page => {
                      const countOnPage = errors.filter(e => (e.pageNumber || 1) === page).length;
                      return (
                        <button
                          key={page}
                          onClick={() => setFilterPage(String(page))}
                          className={`px-2 py-0.5 rounded text-xs whitespace-nowrap transition-colors ${
                            filterPage === String(page)
                              ? 'bg-indigo-600 text-white font-medium'
                              : 'bg-slate-100 text-slate-600 hover:bg-slate-200 dark:bg-slate-800'
                          }`}
                        >
                          หน้า {page} ({countOnPage})
                        </button>
                      );
                    })}
                  </div>
                )}

                {/* Filter by Category */}
                <div className="flex items-center gap-1 overflow-x-auto pb-1 text-xs">
                  <button
                    onClick={() => setFilterCategory('all')}
                    className={`px-2.5 py-1 rounded-md whitespace-nowrap transition-colors ${
                      filterCategory === 'all'
                        ? 'bg-slate-900 text-white dark:bg-white dark:text-slate-900'
                        : 'text-slate-600 hover:bg-slate-100 dark:text-slate-400'
                    }`}
                  >
                    ทั้งหมด ({errors.length})
                  </button>
                  <button
                    onClick={() => setFilterCategory('spelling')}
                    className={`px-2.5 py-1 rounded-md whitespace-nowrap transition-colors ${
                      filterCategory === 'spelling'
                        ? 'bg-rose-600 text-white'
                        : 'text-rose-700 hover:bg-rose-50 dark:text-rose-300'
                    }`}
                  >
                    ตัวสะกด ({errors.filter(e => e.category === 'spelling').length})
                  </button>
                  <button
                    onClick={() => setFilterCategory('grammar')}
                    className={`px-2.5 py-1 rounded-md whitespace-nowrap transition-colors ${
                      filterCategory === 'grammar'
                        ? 'bg-amber-600 text-white'
                        : 'text-amber-700 hover:bg-amber-50 dark:text-amber-300'
                    }`}
                  >
                    ไวยากรณ์ ({errors.filter(e => e.category === 'grammar').length})
                  </button>
                  <button
                    onClick={() => setFilterCategory('formal')}
                    className={`px-2.5 py-1 rounded-md whitespace-nowrap transition-colors ${
                      filterCategory === 'formal'
                        ? 'bg-indigo-600 text-white'
                        : 'text-indigo-700 hover:bg-indigo-50 dark:text-indigo-300'
                    }`}
                  >
                    ทางการ ({errors.filter(e => e.category === 'formal').length})
                  </button>
                </div>
              </div>
            )}

            {/* Error Content: Table View OR Cards View */}
            <div className="mt-3">
              {errors.length === 0 ? (
                <div className="py-12 text-center">
                  <div className="mx-auto flex h-12 w-12 items-center justify-center rounded-2xl bg-indigo-50 text-indigo-600 dark:bg-indigo-950/60 dark:text-indigo-400">
                    <FileCheck className="h-6 w-6" />
                  </div>
                  <h3 className="mt-3 text-sm font-semibold text-slate-800 dark:text-slate-200">
                    ยังไม่มีรายการคำผิด
                  </h3>
                  <p className="mt-1 text-xs text-slate-500 max-w-xs mx-auto">
                    เปิดไฟล์เอกสารหรือพิมพ์ข้อความ แล้วระบบจะวิเคราะห์คำผิดโดยอัตโนมัติ
                  </p>
                </div>
              ) : filteredErrors.length === 0 ? (
                <div className="py-8 text-center text-xs text-slate-500">
                  ไม่มีคำผิดในเงื่อนไขการกรองนี้
                </div>
              ) : viewMode === 'table' ? (
                /* Table View: Detailed page-by-page table */
                <div className="max-h-[560px] overflow-x-auto overflow-y-auto rounded-xl border border-slate-200 dark:border-slate-800">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-slate-50 text-slate-700 dark:bg-slate-800 dark:text-slate-300 sticky top-0 z-10">
                      <tr>
                        <th className="p-2.5 font-semibold">หน้า / บรรทัด</th>
                        <th className="p-2.5 font-semibold">คำผิดเดิม</th>
                        <th className="p-2.5 font-semibold">คำที่ถูกต้อง</th>
                        <th className="p-2.5 font-semibold">บริบท</th>
                        <th className="p-2.5 font-semibold text-right">ดำเนินการ</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
                      {filteredErrors.map((err, idx) => {
                        const isAccepted = err.status === 'accepted';
                        const isRejected = err.status === 'rejected';
                        const isCustom = err.status === 'custom';

                        return (
                          <tr
                            key={err.id || idx}
                            className={`hover:bg-slate-50/80 dark:hover:bg-slate-800/50 ${
                              isAccepted
                                ? 'bg-emerald-50/30 dark:bg-emerald-950/20'
                                : isRejected
                                ? 'bg-slate-50/60 opacity-60'
                                : ''
                            }`}
                          >
                            <td className="p-2.5 whitespace-nowrap">
                              <span className="font-bold text-indigo-700 dark:text-indigo-400">
                                หน้า {err.pageNumber || 1}
                              </span>
                              <div className="text-[10px] text-slate-400">
                                บรรทัด {err.lineIndex || 1}
                              </div>
                            </td>
                            <td className="p-2.5 whitespace-nowrap">
                              <span className="line-through text-rose-600 font-medium bg-rose-50 px-1 py-0.5 rounded dark:bg-rose-950/60">
                                {err.original}
                              </span>
                            </td>
                            <td className="p-2.5 whitespace-nowrap">
                              <span className="text-emerald-700 font-bold bg-emerald-50 px-1.5 py-0.5 rounded dark:bg-emerald-950/60 dark:text-emerald-300">
                                {isCustom && err.customFix ? err.customFix : err.suggestion}
                              </span>
                            </td>
                            <td className="p-2.5 max-w-[180px] truncate text-slate-600 dark:text-slate-400" title={err.context}>
                              {err.context || `[${err.original}]`}
                            </td>
                            <td className="p-2.5 whitespace-nowrap text-right">
                              <div className="flex items-center justify-end gap-1">
                                <button
                                  onClick={() => handleAcceptError(err.id)}
                                  title="ยืนยันแก้ไข"
                                  className={`p-1 rounded ${
                                    isAccepted
                                      ? 'bg-emerald-600 text-white'
                                      : 'text-slate-500 hover:bg-emerald-100 hover:text-emerald-700'
                                  }`}
                                >
                                  <Check className="h-3.5 w-3.5" />
                                </button>
                                <button
                                  onClick={() => handleRejectError(err.id)}
                                  title="คงคำเดิม"
                                  className={`p-1 rounded ${
                                    isRejected
                                      ? 'bg-slate-700 text-white'
                                      : 'text-slate-500 hover:bg-slate-100'
                                  }`}
                                >
                                  <XCircle className="h-3.5 w-3.5" />
                                </button>
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              ) : (
                /* Cards View: Detailed cards with page numbers and full context */
                <div className="space-y-3 max-h-[560px] overflow-y-auto pr-1">
                  {filteredErrors.map((err, idx) => {
                    const isAccepted = err.status === 'accepted';
                    const isRejected = err.status === 'rejected';
                    const isCustom = err.status === 'custom';

                    return (
                      <article
                        key={err.id || idx}
                        className={`rounded-xl border p-3.5 transition-all ${
                          isAccepted
                            ? 'border-emerald-200 bg-emerald-50/40 dark:border-emerald-900/60 dark:bg-emerald-950/20'
                            : isRejected
                            ? 'border-slate-200 bg-slate-50/50 opacity-60 dark:border-slate-800 dark:bg-slate-900/40'
                            : 'border-slate-200 bg-white hover:border-slate-300 dark:border-slate-800 dark:bg-slate-900'
                        }`}
                      >
                        {/* Top: Page number, location, and status */}
                        <div className="flex items-center justify-between text-[11px]">
                          <div className="flex items-center gap-1.5">
                            {/* Page Badge */}
                            <span className="rounded-md bg-indigo-100 px-1.5 py-0.5 font-bold text-indigo-700 dark:bg-indigo-950 dark:text-indigo-300">
                              หน้า {err.pageNumber || 1}
                            </span>
                            <span className="text-slate-400">
                              ย่อหน้า {err.paragraphIndex || 1}, บรรทัด {err.lineIndex || 1}
                            </span>
                          </div>

                          <div className="flex items-center gap-1.5">
                            <span
                              className={`font-semibold ${
                                err.category === 'spelling'
                                  ? 'text-rose-600'
                                  : err.category === 'grammar'
                                  ? 'text-amber-600'
                                  : 'text-indigo-600'
                              }`}
                            >
                              {err.category === 'spelling'
                                ? 'ตัวสะกด'
                                : err.category === 'grammar'
                                ? 'ไวยากรณ์'
                                : 'ทางการ'}
                            </span>

                            {isAccepted && (
                              <span className="flex items-center gap-0.5 text-emerald-600 font-semibold ml-1">
                                <CheckCircle className="h-3 w-3" />
                                <span>แก้ไขแล้ว</span>
                              </span>
                            )}
                            {isRejected && (
                              <span className="flex items-center gap-0.5 text-slate-500 ml-1">
                                <XCircle className="h-3 w-3" />
                                <span>คงเดิม</span>
                              </span>
                            )}
                            {isCustom && (
                              <span className="flex items-center gap-0.5 text-violet-600 font-semibold ml-1">
                                <CheckCircle className="h-3 w-3" />
                                <span>กำหนดเอง</span>
                              </span>
                            )}
                          </div>
                        </div>

                        {/* Word replacement preview */}
                        <div className="mt-2 flex items-center gap-2 text-sm">
                          <span className="line-through text-rose-600 font-medium bg-rose-50 px-1.5 py-0.5 rounded dark:bg-rose-950/60">
                            {err.original}
                          </span>
                          <ChevronRight className="h-3.5 w-3.5 text-slate-400" />
                          <span className="text-emerald-700 font-bold bg-emerald-50 px-2 py-0.5 rounded dark:bg-emerald-950/60 dark:text-emerald-300">
                            {isCustom && err.customFix ? err.customFix : err.suggestion}
                          </span>
                        </div>

                        {/* Context Snippet */}
                        {err.context && (
                          <div className="mt-2 rounded-md bg-slate-50 p-2 text-xs font-sans text-slate-700 dark:bg-slate-800/60 dark:text-slate-300 leading-relaxed border-l-2 border-indigo-400">
                            <span className="text-slate-400 text-[10px] block mb-0.5 font-mono">
                              บริบทในเอกสาร:
                            </span>
                            <span>{err.context}</span>
                          </div>
                        )}

                        {/* Linguistic reason */}
                        <p className="mt-1.5 text-xs text-slate-600 dark:text-slate-400 leading-relaxed">
                          {err.reason}
                        </p>

                        {/* Inline Custom Edit Input if active */}
                        {editingErrorId === err.id && (
                          <div className="mt-2 flex items-center gap-2">
                            <input
                              type="text"
                              value={customInputValue}
                              onChange={e => setCustomInputValue(e.target.value)}
                              placeholder="ระบุคำที่ต้องการแก้ไข..."
                              className="flex-1 rounded-md border border-slate-300 px-2.5 py-1 text-xs focus:border-indigo-500 focus:outline-hidden dark:border-slate-700 dark:bg-slate-800"
                            />
                            <button
                              onClick={() => handleSaveCustomFix(err.id)}
                              className="rounded-md bg-indigo-600 px-2.5 py-1 text-xs font-medium text-white hover:bg-indigo-700 cursor-pointer"
                            >
                              บันทึก
                            </button>
                            <button
                              onClick={() => setEditingErrorId(null)}
                              className="text-xs text-slate-500 hover:text-slate-800 cursor-pointer"
                            >
                              ยกเลิก
                            </button>
                          </div>
                        )}

                        {/* Action buttons */}
                        <div className="mt-3 flex items-center justify-between border-t border-slate-100 pt-2 text-xs dark:border-slate-800">
                          <div className="flex items-center gap-1.5">
                            <button
                              onClick={() => handleAcceptError(err.id)}
                              className={`flex items-center gap-1 rounded-md px-2.5 py-1 font-medium transition-colors cursor-pointer ${
                                isAccepted
                                  ? 'bg-emerald-600 text-white'
                                  : 'bg-slate-100 text-slate-700 hover:bg-emerald-100 hover:text-emerald-800 dark:bg-slate-800 dark:text-slate-300'
                              }`}
                            >
                              <Check className="h-3 w-3" />
                              <span>ยืนยันแก้ไข</span>
                            </button>

                            <button
                              onClick={() => handleRejectError(err.id)}
                              className={`flex items-center gap-1 rounded-md px-2 py-1 font-medium transition-colors cursor-pointer ${
                                isRejected
                                  ? 'bg-slate-700 text-white dark:bg-slate-600'
                                  : 'text-slate-500 hover:bg-slate-100 hover:text-slate-800 dark:hover:bg-slate-800'
                              }`}
                            >
                              <span>คงคำเดิม</span>
                            </button>
                          </div>

                          <div className="flex items-center gap-1">
                            <button
                              onClick={() => {
                                setEditingErrorId(err.id);
                                setCustomInputValue(err.suggestion);
                              }}
                              className="text-slate-500 hover:text-indigo-600 px-1.5 py-1 cursor-pointer"
                              title="พิมพ์แก้ไขคำนี้เอง"
                            >
                              แก้ไขเอง
                            </button>

                            <button
                              onClick={() => {
                                onAddCustomWord(err.original);
                                handleRejectError(err.id);
                              }}
                              title="เพิ่มคำนี้ลงคลังคำยกเว้นขององค์กร (จะไม่เตือนซ้ำอีก)"
                              className="flex items-center gap-1 text-slate-500 hover:text-slate-800 px-1.5 py-1 cursor-pointer"
                            >
                              <BookPlus className="h-3 w-3" />
                              <span className="hidden sm:inline">จำคำนี้</span>
                            </button>
                          </div>
                        </div>
                      </article>
                    );
                  })}
                </div>
              )}
            </div>
          </div>
        </section>
      </div>
    </div>
  );
};

/**
 * Highlighted Diff Renderer
 * Displays original text with red strikethrough for errors and green bold for replacements
 */
const HighlightedDiffText: React.FC<{
  originalText: string;
  errors: SpellError[];
}> = ({ originalText, errors }) => {
  if (!errors || errors.length === 0) {
    return <span>{originalText}</span>;
  }

  // Sort errors by startIndex
  const sorted = [...errors].sort((a, b) => a.startIndex - b.startIndex);
  const elements: React.ReactNode[] = [];
  let cursor = 0;

  sorted.forEach((err, idx) => {
    if (err.startIndex >= cursor && err.startIndex < originalText.length) {
      // Clean prefix text
      if (err.startIndex > cursor) {
        elements.push(
          <span key={`txt_${idx}`}>{originalText.substring(cursor, err.startIndex)}</span>
        );
      }

      const isAccepted = err.status === 'accepted';
      const isCustom = err.status === 'custom';
      const isRejected = err.status === 'rejected';

      if (isAccepted || isCustom) {
        elements.push(
          <span key={`diff_${err.id || idx}`} className="inline-flex items-center mx-1">
            <span className="line-through text-rose-500 bg-rose-50 px-1 rounded text-xs dark:bg-rose-950/60">
              {err.original}
            </span>
            <span className="text-emerald-700 bg-emerald-100 font-bold px-1.5 rounded ml-1 dark:bg-emerald-900/60 dark:text-emerald-200">
              {isCustom && err.customFix ? err.customFix : err.suggestion}
            </span>
          </span>
        );
      } else if (isRejected) {
        elements.push(
          <span
            key={`rej_${err.id || idx}`}
            className="text-slate-700 underline decoration-dotted font-medium"
            title="ผู้ใช้ปฏิเสธการแก้ไข คงคำเดิมไว้"
          >
            {err.original}
          </span>
        );
      } else {
        // Pending
        elements.push(
          <span
            key={`pen_${err.id || idx}`}
            className="bg-amber-100 text-amber-900 font-semibold px-1 rounded border-b-2 border-amber-500 dark:bg-amber-950/60 dark:text-amber-200"
            title={`${err.reason} -> แนะนำ: ${err.suggestion} (หน้า ${err.pageNumber || 1})`}
          >
            {err.original}
          </span>
        );
      }

      cursor = err.endIndex;
    }
  });

  if (cursor < originalText.length) {
    elements.push(<span key="tail">{originalText.substring(cursor)}</span>);
  }

  return <div>{elements}</div>;
};
