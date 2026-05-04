/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback, useMemo } from 'react';
import { FileText, Upload, CircleCheck as CheckCircle2, CircleAlert as AlertCircle, Loader as Loader2, Download, Play, Table as TableIcon, Archive, BookOpen, CalendarDays, Key, Eye, EyeOff, ChevronDown, ChevronUp } from 'lucide-react';
import { useDropzone } from 'react-dropzone';
import * as XLSX from 'xlsx';
import * as pdfjsLib from 'pdfjs-dist';
import { GoogleGenAI } from "@google/genai";
import { motion, AnimatePresence } from 'motion/react';

// Configure PDF.js worker
if (typeof window !== 'undefined' && 'Worker' in window) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjsLib.version}/build/pdf.worker.min.mjs`;
}

interface ArchiveRow {
  no: number;
  uraian: string;
  tahun: string;
  kodeKlasifikasi?: string;
  musnahPermanen?: string;
  dasarKlasifikasi?: string;
  processed: boolean;
  status: 'pending' | 'processing' | 'completed' | 'error';
}

interface ReferenceDoc {
  name: string;
  content: string;
  type: 'JK' | 'JRA';
  loaded: boolean;
}

const STORAGE_KEY = 'archive_gemini_api_key';

export default function App() {
  const [geminiStatus, setGeminiStatus] = useState<string>('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [rows, setRows] = useState<ArchiveRow[]>([]);
  const [refJK, setRefJK] = useState<ReferenceDoc>({ name: '', content: '', type: 'JK', loaded: false });
  const [refJRA, setRefJRA] = useState<ReferenceDoc>({ name: '', content: '', type: 'JRA', loaded: false });
  const [error, setError] = useState<string | null>(null);
  const [apiKey, setApiKey] = useState<string>(() => localStorage.getItem(STORAGE_KEY) || '');
  const [showApiKey, setShowApiKey] = useState(false);
  const [apiKeyInputValue, setApiKeyInputValue] = useState<string>(() => localStorage.getItem(STORAGE_KEY) || '');
  const [apiKeySaved, setApiKeySaved] = useState(false);
  const [configOpen, setConfigOpen] = useState(!localStorage.getItem(STORAGE_KEY));

  const saveApiKey = () => {
    const trimmed = apiKeyInputValue.trim();
    if (!trimmed) return;
    localStorage.setItem(STORAGE_KEY, trimmed);
    setApiKey(trimmed);
    setApiKeySaved(true);
    setTimeout(() => setApiKeySaved(false), 2000);
  };

  const clearApiKey = () => {
    localStorage.removeItem(STORAGE_KEY);
    setApiKey('');
    setApiKeyInputValue('');
  };

  const genAI = useMemo(() => {
    if (!apiKey) return null;
    return new GoogleGenAI({ apiKey });
  }, [apiKey]);

  const onDropExcel = useCallback(async (acceptedFiles: File[]) => {
    const file = acceptedFiles[0];
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 'A' }) as any[];

        let dataStartIdx = 0;
        let headerRowFound = false;
        for (let i = 0; i < jsonData.length; i++) {
          const rowValues = Object.values(jsonData[i]).join(' ').toLowerCase();
          if (rowValues.includes('uraian') || rowValues.includes('informasi')) {
            dataStartIdx = i + 1;
            headerRowFound = true;
            break;
          }
        }

        const effectiveStartIdx = headerRowFound ? dataStartIdx : 0;

        const formattedRows: ArchiveRow[] = jsonData.slice(effectiveStartIdx)
          .filter(row => row.D || row.C || row.A)
          .map((row, idx) => ({
            no: idx + 1,
            uraian: row.D || row.C || String(row.B || ''),
            tahun: String(row.E || row.D || ''),
            processed: false,
            status: 'pending'
          }));

        setRows(formattedRows);
        setError(null);
      } catch (err) {
        setError('Gagal membaca file Excel. Pastikan format benar.');
      }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const extractTextFromPdf = async (file: File) => {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    let fullText = '';
    const numPages = Math.min(pdf.numPages, 100);
    for (let i = 1; i <= numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      const strings = content.items.map((item: any) => item.str);
      fullText += strings.join(' ') + '\n';
    }
    return fullText;
  };

  const onDropJK = useCallback(async (acceptedFiles: File[]) => {
    const file = acceptedFiles[0];
    try {
      const text = await extractTextFromPdf(file);
      setRefJK({ name: file.name, content: text, type: 'JK', loaded: true });
    } catch (err) {
      setError('Gagal membaca file SK Klasifikasi PDF.');
    }
  }, []);

  const onDropJRA = useCallback(async (acceptedFiles: File[]) => {
    const file = acceptedFiles[0];
    try {
      const text = await extractTextFromPdf(file);
      setRefJRA({ name: file.name, content: text, type: 'JRA', loaded: true });
    } catch (err) {
      setError('Gagal membaca file SK JRA PDF.');
    }
  }, []);

  const { getRootProps: getExcelProps, getInputProps: getExcelInputProps, isDragActive: excelActive } = useDropzone({
    onDrop: onDropExcel,
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls']
    },
    multiple: false
  } as any);

  const { getRootProps: getJKProps, getInputProps: getJKInputProps, isDragActive: jkActive } = useDropzone({
    onDrop: onDropJK,
    accept: { 'application/pdf': ['.pdf'] },
    multiple: false
  } as any);

  const { getRootProps: getJRAProps, getInputProps: getJRAInputProps, isDragActive: jraActive } = useDropzone({
    onDrop: onDropJRA,
    accept: { 'application/pdf': ['.pdf'] },
    multiple: false
  } as any);

  const startClassification = async () => {
    if (!genAI || !refJK.loaded || !refJRA.loaded || rows.length === 0) return;

    setIsProcessing(true);
    setGeminiStatus('Menyiapkan model...');
    setError(null);

    try {
      const model = genAI.getGenerativeModel({
        model: "gemini-1.5-flash",
        systemInstruction: `Anda adalah pakar klasifikasi arsip dengan keahlian mendalam dalam sistem kearsipan berbasis SK Klasifikasi dan Jadwal Retensi Arsip (JRA).

DOKUMEN REFERENSI SK KLASIFIKASI:
---
${refJK.content.substring(0, 45000)}
---

DOKUMEN REFERENSI SK JRA:
---
${refJRA.content.substring(0, 45000)}
---

TUGAS UTAMA:
Untuk setiap arsip yang diberikan, tentukan:
1. Kode Klasifikasi yang paling spesifik dan tepat (level terdalam dari hierarki kode, misalnya "KP.01.03" bukan hanya "KP")
2. Nasib Akhir: HANYA boleh "Musnah" atau "Permanen" sesuai kolom nasib akhir di JRA
3. Dasar Klasifikasi: alasan singkat dan jelas mengapa kode tersebut dipilih

ATURAN KETAT:
- Selalu gunakan kode klasifikasi yang paling spesifik (paling dalam dalam hierarki)
- Nasib akhir WAJIB merujuk langsung ke JRA, jangan menebak
- Bedakan arsip mahasiswa vs arsip kepegawaian/staf
- Jika arsip menyangkut keuangan, gunakan kode keuangan yang sesuai
- Jika tidak ditemukan kode yang tepat, gunakan kode terdekat dan jelaskan di dasar

FORMAT RESPONS:
Balas HANYA dengan JSON murni tanpa markdown, tanpa komentar:
{"kode": "XX.XX.XX", "nasib": "Musnah" atau "Permanen", "dasar": "alasan singkat"}`
      });

      const updatedRows = [...rows];

      for (let i = 0; i < updatedRows.length; i++) {
        if (updatedRows[i].status === 'completed') continue;

        setGeminiStatus(`Memproses ${i + 1} dari ${updatedRows.length}...`);

        try {
          updatedRows[i].status = 'processing';
          setRows([...updatedRows]);

          const prompt = `Klasifikasikan arsip berikut berdasarkan SK Klasifikasi dan JRA yang telah diberikan:

Uraian Arsip: "${updatedRows[i].uraian}"
Tahun: ${updatedRows[i].tahun}

Tentukan kode klasifikasi yang paling spesifik, nasib akhir (Musnah/Permanen), dan dasar singkat pemilihan kode.
Balas HANYA dengan JSON: {"kode": "...", "nasib": "...", "dasar": "..."}`;

          const result = await model.generateContent(prompt);
          const response = await result.response;
          const text = response.text().trim();

          const jsonMatch = text.match(/\{[^{}]*"kode"[^{}]*"nasib"[^{}]*"dasar"[^{}]*\}/)
                         || text.match(/\{.*?\}/s);

          if (jsonMatch) {
            try {
              const data = JSON.parse(jsonMatch[0]);
              if (data.kode && data.nasib) {
                updatedRows[i].kodeKlasifikasi = data.kode;
                updatedRows[i].musnahPermanen = data.nasib;
                updatedRows[i].dasarKlasifikasi = data.dasar || '';
                updatedRows[i].status = 'completed';
              } else {
                updatedRows[i].status = 'error';
              }
            } catch {
              updatedRows[i].status = 'error';
            }
          } else {
            updatedRows[i].status = 'error';
          }
        } catch (err: any) {
          console.error('Row error:', err);
          if (err?.message?.includes('API_KEY') || err?.message?.includes('401') || err?.message?.includes('403')) {
            setError('API key tidak valid atau tidak memiliki akses. Periksa kembali API key Anda.');
            setIsProcessing(false);
            setGeminiStatus('');
            return;
          }
          updatedRows[i].status = 'error';
        }
        setRows([...updatedRows]);

        if (i < updatedRows.length - 1) {
          await new Promise(r => setTimeout(r, 300));
        }
      }

      setIsProcessing(false);
      setGeminiStatus('Proses selesai');
      setTimeout(() => setGeminiStatus(''), 3000);
    } catch (err: any) {
      console.error('Global error:', err);
      setError('Terjadi kesalahan: ' + (err?.message || 'Unknown error'));
      setIsProcessing(false);
      setGeminiStatus('');
    }
  };

  const exportToExcel = () => {
    const exportData = rows.map(r => ({
      'No': r.no,
      'Kode Klasifikasi': r.kodeKlasifikasi || '',
      'Uraian Informasi Arsip': r.uraian,
      'Tahun': r.tahun,
      'Dasar Klasifikasi': r.dasarKlasifikasi || '',
      'Nasib Akhir': r.musnahPermanen || ''
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Hasil Klasifikasi");
    XLSX.writeFile(wb, "Hasil_Klasifikasi_Arsip.xlsx");
  };

  const stats = useMemo(() => {
    const total = rows.length;
    const completed = rows.filter(r => r.status === 'completed').length;
    const errored = rows.filter(r => r.status === 'error').length;
    const permanent = rows.filter(r => r.musnahPermanen === 'Permanen').length;
    const destroyed = rows.filter(r => r.musnahPermanen === 'Musnah').length;
    return { total, completed, errored, permanent, destroyed };
  }, [rows]);

  const canProcess = genAI && refJK.loaded && refJRA.loaded && rows.length > 0 && !isProcessing;

  return (
    <div className="min-h-screen pb-24 bg-gray-50">
      {/* Header */}
      <header className="bg-gray-900 py-12 px-8 text-white shadow-2xl">
        <div className="max-w-6xl mx-auto flex flex-col items-center text-center">
          <motion.div
            initial={{ scale: 0.8, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="w-16 h-16 bg-amber-400 rounded-2xl flex items-center justify-center shadow-xl mb-5"
          >
            <Archive className="w-8 h-8 text-gray-900" />
          </motion.div>
          <motion.h1
            initial={{ opacity: 0, y: 10 }}
            animate={{ opacity: 1, y: 0 }}
            className="text-4xl font-black tracking-tight sm:text-5xl mb-3"
          >
            Archive <span className="text-amber-400">Classifier</span>
          </motion.h1>
          <motion.p
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            transition={{ delay: 0.15 }}
            className="text-gray-400 text-base max-w-xl"
          >
            Klasifikasi dan penentuan retensi arsip otomatis menggunakan Gemini AI berdasarkan SK Klasifikasi dan Jadwal Retensi Arsip (JRA).
          </motion.p>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-6 mt-8 grid grid-cols-1 lg:grid-cols-12 gap-8">
        {/* Sidebar */}
        <aside className="lg:col-span-4 space-y-5">

          {/* API Key Panel */}
          <section className="bg-white border border-gray-200 rounded-2xl shadow-sm overflow-hidden">
            <button
              onClick={() => setConfigOpen(v => !v)}
              className="w-full flex items-center justify-between px-5 py-4 hover:bg-gray-50 transition-colors"
            >
              <div className="flex items-center gap-3">
                <div className={`p-2 rounded-lg ${apiKey ? 'bg-green-100 text-green-600' : 'bg-amber-50 text-amber-600'}`}>
                  <Key className="w-4 h-4" />
                </div>
                <div className="text-left">
                  <p className="text-sm font-bold text-gray-800">Gemini API Key</p>
                  <p className="text-xs text-gray-400">
                    {apiKey ? 'Terkonfigurasi' : 'Belum dikonfigurasi'}
                  </p>
                </div>
              </div>
              {configOpen ? <ChevronUp className="w-4 h-4 text-gray-400" /> : <ChevronDown className="w-4 h-4 text-gray-400" />}
            </button>

            <AnimatePresence>
              {configOpen && (
                <motion.div
                  initial={{ height: 0, opacity: 0 }}
                  animate={{ height: 'auto', opacity: 1 }}
                  exit={{ height: 0, opacity: 0 }}
                  transition={{ duration: 0.2 }}
                  className="overflow-hidden border-t border-gray-100"
                >
                  <div className="p-5 space-y-3">
                    <p className="text-xs text-gray-500">
                      Masukkan Google Gemini API key Anda. Key disimpan di browser dan tidak dikirim ke server manapun.
                    </p>
                    <div className="relative">
                      <input
                        type={showApiKey ? 'text' : 'password'}
                        value={apiKeyInputValue}
                        onChange={e => setApiKeyInputValue(e.target.value)}
                        onKeyDown={e => e.key === 'Enter' && saveApiKey()}
                        placeholder="AIza..."
                        className="w-full pr-10 pl-3 py-2.5 border border-gray-200 rounded-xl text-sm font-mono focus:outline-none focus:ring-2 focus:ring-gray-300 bg-gray-50"
                      />
                      <button
                        onClick={() => setShowApiKey(v => !v)}
                        className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600"
                      >
                        {showApiKey ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                      </button>
                    </div>
                    <div className="flex gap-2">
                      <button
                        onClick={saveApiKey}
                        disabled={!apiKeyInputValue.trim()}
                        className="flex-1 bg-gray-900 hover:bg-gray-700 disabled:bg-gray-100 disabled:text-gray-400 text-white text-xs font-bold py-2.5 rounded-xl transition-colors"
                      >
                        {apiKeySaved ? 'Tersimpan!' : 'Simpan Key'}
                      </button>
                      {apiKey && (
                        <button
                          onClick={clearApiKey}
                          className="px-4 text-xs font-bold text-red-500 hover:bg-red-50 border border-red-200 rounded-xl transition-colors"
                        >
                          Hapus
                        </button>
                      )}
                    </div>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>
          </section>

          {/* File Upload Panel */}
          <section className="bg-white border border-gray-200 rounded-2xl shadow-sm p-5">
            <h2 className="text-xs font-bold text-gray-500 mb-4 flex items-center gap-2 uppercase tracking-wider">
              <Upload className="w-3.5 h-3.5" /> Dokumen Referensi & Data
            </h2>

            <div className="space-y-3">
              <div {...getJKProps()} className={`p-4 border-2 border-dashed rounded-xl cursor-pointer transition-all flex items-center gap-3 ${jkActive ? 'border-gray-500 bg-gray-50' : 'border-gray-200 hover:border-gray-400'}`}>
                <input {...getJKInputProps()} />
                <div className={`p-2.5 rounded-lg shrink-0 ${refJK.loaded ? 'bg-green-100 text-green-600' : 'bg-gray-100 text-gray-400'}`}>
                  <FileText className="w-5 h-5" />
                </div>
                <div className="flex-1 overflow-hidden">
                  <p className="text-sm font-semibold text-gray-700">SK Klasifikasi</p>
                  <p className="text-xs text-gray-400 truncate">
                    {refJK.loaded ? refJK.name : 'Klik atau seret PDF di sini'}
                  </p>
                </div>
                {refJK.loaded && <CheckCircle2 className="w-4 h-4 text-green-500 shrink-0" />}
              </div>

              <div {...getJRAProps()} className={`p-4 border-2 border-dashed rounded-xl cursor-pointer transition-all flex items-center gap-3 ${jraActive ? 'border-gray-500 bg-gray-50' : 'border-gray-200 hover:border-gray-400'}`}>
                <input {...getJRAInputProps()} />
                <div className={`p-2.5 rounded-lg shrink-0 ${refJRA.loaded ? 'bg-green-100 text-green-600' : 'bg-gray-100 text-gray-400'}`}>
                  <BookOpen className="w-5 h-5" />
                </div>
                <div className="flex-1 overflow-hidden">
                  <p className="text-sm font-semibold text-gray-700">SK JRA</p>
                  <p className="text-xs text-gray-400 truncate">
                    {refJRA.loaded ? refJRA.name : 'Klik atau seret PDF di sini'}
                  </p>
                </div>
                {refJRA.loaded && <CheckCircle2 className="w-4 h-4 text-green-500 shrink-0" />}
              </div>

              <div {...getExcelProps()} className={`p-4 border-2 border-dashed rounded-xl cursor-pointer transition-all flex items-center gap-3 ${excelActive ? 'border-gray-500 bg-gray-50' : 'border-gray-200 hover:border-gray-400'}`}>
                <input {...getExcelInputProps()} />
                <div className={`p-2.5 rounded-lg shrink-0 ${rows.length > 0 ? 'bg-green-100 text-green-600' : 'bg-gray-100 text-gray-400'}`}>
                  <TableIcon className="w-5 h-5" />
                </div>
                <div className="flex-1 overflow-hidden">
                  <p className="text-sm font-semibold text-gray-700">Daftar Arsip</p>
                  <p className="text-xs text-gray-400 truncate">
                    {rows.length > 0 ? `${rows.length} baris terdeteksi` : 'Pilih file Excel (.xlsx)'}
                  </p>
                </div>
                {rows.length > 0 && <CheckCircle2 className="w-4 h-4 text-green-500 shrink-0" />}
              </div>
            </div>

            {/* Readiness checklist */}
            <div className="mt-4 p-3 bg-gray-50 rounded-xl space-y-1.5">
              {[
                { label: 'API Key', ready: !!apiKey },
                { label: 'SK Klasifikasi', ready: refJK.loaded },
                { label: 'SK JRA', ready: refJRA.loaded },
                { label: 'Data Arsip', ready: rows.length > 0 },
              ].map(item => (
                <div key={item.label} className="flex items-center gap-2">
                  <div className={`w-1.5 h-1.5 rounded-full ${item.ready ? 'bg-green-500' : 'bg-gray-300'}`} />
                  <span className={`text-xs ${item.ready ? 'text-gray-700 font-medium' : 'text-gray-400'}`}>{item.label}</span>
                </div>
              ))}
            </div>

            {error && (
              <motion.div
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                className="mt-4 p-3 bg-red-50 text-red-700 rounded-xl flex items-start gap-2.5 text-xs border border-red-100"
              >
                <AlertCircle className="w-4 h-4 shrink-0 mt-0.5" />
                <p className="font-medium">{error}</p>
              </motion.div>
            )}

            <button
              onClick={startClassification}
              disabled={!canProcess}
              className="w-full mt-5 bg-gray-900 hover:bg-gray-700 disabled:bg-gray-100 disabled:text-gray-400 text-white font-bold py-3.5 rounded-xl transition-all shadow-md flex items-center justify-center gap-2 active:scale-95"
            >
              {isProcessing ? <Loader2 className="w-5 h-5 animate-spin" /> : <Play className="w-5 h-5 fill-current" />}
              {isProcessing ? 'Memproses...' : 'Mulai Klasifikasi'}
            </button>
          </section>

          {/* Stats */}
          {rows.length > 0 && (
            <section className="bg-white border border-gray-200 rounded-2xl shadow-sm p-5">
              <h3 className="text-xs font-bold text-gray-400 uppercase tracking-wider mb-4">Laporan</h3>

              <div className="w-full h-2 bg-gray-100 rounded-full overflow-hidden mb-4 flex">
                <div
                  className="h-full bg-green-500 transition-all duration-500"
                  style={{ width: `${(stats.completed / stats.total) * 100}%` }}
                />
                <div
                  className="h-full bg-amber-400 transition-all duration-300"
                  style={{ width: `${(rows.filter(r => r.status === 'processing').length / stats.total) * 100}%` }}
                />
                <div
                  className="h-full bg-red-400 transition-all duration-300"
                  style={{ width: `${(stats.errored / stats.total) * 100}%` }}
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div className="p-3 bg-gray-50 rounded-xl">
                  <p className="text-[10px] font-bold text-gray-400 uppercase">Total</p>
                  <p className="text-xl font-black text-gray-800">{stats.total}</p>
                </div>
                <div className="p-3 bg-green-50 rounded-xl">
                  <p className="text-[10px] font-bold text-green-600 uppercase">Selesai</p>
                  <p className="text-xl font-black text-green-700">{stats.completed}</p>
                </div>
                {stats.completed > 0 && (
                  <>
                    <div className="p-3 bg-blue-50 rounded-xl">
                      <p className="text-[10px] font-bold text-blue-500 uppercase">Permanen</p>
                      <p className="text-xl font-black text-blue-700">{stats.permanent}</p>
                    </div>
                    <div className="p-3 bg-orange-50 rounded-xl">
                      <p className="text-[10px] font-bold text-orange-500 uppercase">Musnah</p>
                      <p className="text-xl font-black text-orange-700">{stats.destroyed}</p>
                    </div>
                  </>
                )}
                {stats.errored > 0 && (
                  <div className="col-span-2 p-3 bg-red-50 rounded-xl">
                    <p className="text-[10px] font-bold text-red-500 uppercase">Gagal</p>
                    <p className="text-xl font-black text-red-700">{stats.errored}</p>
                  </div>
                )}
              </div>
            </section>
          )}

          {stats.completed > 0 && (
            <motion.button
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              onClick={exportToExcel}
              className="w-full bg-green-700 hover:bg-green-800 text-white font-bold py-3.5 rounded-xl transition-all shadow-md flex items-center justify-center gap-2"
            >
              <Download className="w-5 h-5" /> Unduh Hasil (.xlsx)
            </motion.button>
          )}
        </aside>

        {/* Main Table */}
        <section className="lg:col-span-8 space-y-5">
          <div className="flex items-center justify-between">
            <h2 className="text-xl font-black text-gray-800">
              Daftar Kerja Arsip
            </h2>
            <AnimatePresence>
              {geminiStatus && (
                <motion.div
                  initial={{ opacity: 0, x: 10 }}
                  animate={{ opacity: 1, x: 0 }}
                  exit={{ opacity: 0, x: 10 }}
                  className="flex items-center gap-2 px-4 py-2 bg-gray-900 text-amber-400 rounded-full text-xs font-bold"
                >
                  <Loader2 className="w-3 h-3 animate-spin"/>
                  {geminiStatus}
                </motion.div>
              )}
            </AnimatePresence>
          </div>

          <div className="bg-white border border-gray-200 rounded-2xl shadow-sm overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-gray-50 border-b border-gray-100">
                    <th className="px-5 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-wider w-14">#</th>
                    <th className="px-5 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-wider">Informasi Arsip</th>
                    <th className="px-5 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-wider w-44 text-center">Hasil Klasifikasi</th>
                    <th className="px-5 py-4 text-[10px] font-bold text-gray-400 uppercase tracking-wider w-20 text-right">Status</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-50">
                  {rows.length === 0 ? (
                    <tr>
                      <td colSpan={4} className="py-24 text-center">
                        <div className="flex flex-col items-center gap-4 opacity-30">
                          <div className="p-6 bg-gray-100 rounded-2xl">
                            <Archive className="w-12 h-12 text-gray-400" />
                          </div>
                          <p className="font-semibold text-gray-500 text-sm max-w-xs">
                            Unggah file Excel untuk memulai analisis
                          </p>
                        </div>
                      </td>
                    </tr>
                  ) : (
                    rows.map((row) => (
                      <motion.tr
                        key={row.no}
                        layout
                        className={`transition-colors ${row.status === 'processing' ? 'bg-amber-50/60' : 'hover:bg-gray-50/80'}`}
                      >
                        <td className="px-5 py-4">
                          <div className={`w-7 h-7 rounded-lg flex items-center justify-center font-bold text-xs ${
                            row.status === 'completed' ? 'bg-green-100 text-green-700' :
                            row.status === 'error' ? 'bg-red-100 text-red-600' :
                            'bg-gray-100 text-gray-400'
                          }`}>
                            {row.no}
                          </div>
                        </td>
                        <td className="px-5 py-4">
                          <div className="max-w-md">
                            <p className="text-sm font-semibold text-gray-800 leading-snug mb-1.5">
                              {row.uraian}
                            </p>
                            <div className="flex items-center flex-wrap gap-2">
                              <span className="flex items-center gap-1 text-[10px] font-semibold text-gray-400">
                                <CalendarDays className="w-3 h-3"/> {row.tahun || '-'}
                              </span>
                              {row.dasarKlasifikasi && (
                                <span className="text-[10px] text-gray-500 bg-gray-100 px-2 py-0.5 rounded-md italic max-w-xs truncate">
                                  {row.dasarKlasifikasi}
                                </span>
                              )}
                            </div>
                          </div>
                        </td>
                        <td className="px-5 py-4 text-center">
                          <div className="flex flex-col items-center gap-1.5">
                            <span className={`font-mono text-xs font-bold px-3 py-1 rounded-lg ${
                              row.kodeKlasifikasi ? 'bg-gray-900 text-white' : 'bg-gray-100 text-gray-300'
                            }`}>
                              {row.kodeKlasifikasi || '---'}
                            </span>
                            {row.musnahPermanen && (
                              <span className={`text-[9px] font-bold px-2 py-0.5 rounded border ${
                                row.musnahPermanen === 'Permanen'
                                  ? 'border-blue-500 text-blue-600 bg-blue-50'
                                  : 'border-orange-400 text-orange-600 bg-orange-50'
                              }`}>
                                {row.musnahPermanen}
                              </span>
                            )}
                          </div>
                        </td>
                        <td className="px-5 py-4 text-right">
                          {row.status === 'completed' ? (
                            <CheckCircle2 className="w-5 h-5 text-green-500 ml-auto" />
                          ) : row.status === 'processing' ? (
                            <Loader2 className="w-5 h-5 text-amber-500 animate-spin ml-auto" />
                          ) : row.status === 'error' ? (
                            <AlertCircle className="w-5 h-5 text-red-400 ml-auto" />
                          ) : (
                            <div className="w-5 h-5 rounded-full border-2 border-gray-200 ml-auto" />
                          )}
                        </td>
                      </motion.tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </section>
      </main>

      {/* Footer */}
      <footer className="fixed bottom-0 left-0 right-0 bg-white/90 backdrop-blur-md border-t border-gray-100 py-3 px-8 flex justify-between items-center z-50">
        <div className="flex items-center gap-3">
          <div className="p-1.5 bg-gray-900 rounded-lg">
            <Archive className="w-4 h-4 text-amber-400" />
          </div>
          <span className="text-xs font-bold text-gray-500 uppercase tracking-wider">
            Archive Classifier
          </span>
        </div>
        <div className="flex items-center gap-4">
          <span className="text-[10px] text-gray-400">Powered by Google Gemini AI</span>
          <span className="text-[10px] font-bold bg-gray-100 text-gray-500 px-2.5 py-1 rounded-full">v2.0</span>
        </div>
      </footer>
    </div>
  );
}
