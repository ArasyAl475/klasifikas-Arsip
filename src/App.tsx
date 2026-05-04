/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback, useMemo } from 'react';
import { 
  FileText, 
  Upload, 
  CheckCircle2, 
  AlertCircle, 
  Loader2, 
  Download, 
  Play, 
  Table as TableIcon,
  Archive,
  BookOpen,
  CalendarDays,
  ExternalLink
} from 'lucide-react';
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

export default function App() {
  const [geminiStatus, setGeminiStatus] = useState<string>('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [rows, setRows] = useState<ArchiveRow[]>([]);
  const [refJK, setRefJK] = useState<ReferenceDoc>({ name: '', content: '', type: 'JK', loaded: false });
  const [refJRA, setRefJRA] = useState<ReferenceDoc>({ name: '', content: '', type: 'JRA', loaded: false });
  const [error, setError] = useState<string | null>(null);

  const genAI = useMemo(() => {
    // Vite handles process.env.GEMINI_API_KEY replacement via define config
    const apiKey = typeof process !== 'undefined' ? process.env.GEMINI_API_KEY : undefined;
    if (!apiKey || apiKey === 'undefined') return null;
    return new GoogleGenAI({ apiKey });
  }, []);

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
          .filter(row => row.D || row.C || row.A) // Basic safety check for data presence
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
      setError('Gagal membaca SK Klasifikasi PDF.');
    }
  }, []);

  const onDropJRA = useCallback(async (acceptedFiles: File[]) => {
    const file = acceptedFiles[0];
    try {
      const text = await extractTextFromPdf(file);
      setRefJRA({ name: file.name, content: text, type: 'JRA', loaded: true });
    } catch (err) {
      setError('Gagal membaca SK JRA PDF.');
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

    const model = genAI.getGenerativeModel({ 
      model: "gemini-1.5-flash",
      systemInstruction: `Anda adalah pakar klasifikasi arsip Universitas Diponegoro (Undip). 
      
      REFERENSI KLASIFIKASI:
      ${refJK.content.substring(0, 50000)} // Limiting size to avoid context window issues in flash if too large
      
      REFERENSI JRA:
      ${refJRA.content.substring(0, 50000)}
      
      TUGAS:
      Tentukan Kode Klasifikasi dan Nasib Akhir (Musnah/Permanen) untuk setiap uraian arsip.
      
      ATURAN PENTING:
      1. Kode Klasifikasi: Cari yang paling spesifik (level terdalam, misal AK.02.01).
      2. Nasib Akhir: Harus 'Musnah' atau 'Permanen'. Lihat "nasib akhir" di JRA.
      3. Dasar Klasifikasi: Alasan singkat pengambilan kode tersebut.
      4. Bedakan Mahasiswa vs Staf/Dosen.
      
      Balas HANYA dengan JSON: {"kode": "...", "nasib": "...", "dasar": "..."}`
    });

    const updatedRows = [...rows];
    
    for (let i = 0; i < updatedRows.length; i++) {
      if (updatedRows[i].status === 'completed') continue;
      
      try {
        updatedRows[i].status = 'processing';
        setRows([...updatedRows]);
        
        const prompt = `Analisis arsip ini:
        Uraian: ${updatedRows[i].uraian}
        Tahun: ${updatedRows[i].tahun}
        
        Kembalikan JSON.`;

        const result = await model.generateContent(prompt);
        const response = await result.response;
        const text = response.text();
        
        const jsonMatch = text.match(/\{.*\}/s);
        if (jsonMatch) {
          const data = JSON.parse(jsonMatch[0]);
          updatedRows[i].kodeKlasifikasi = data.kode;
          updatedRows[i].musnahPermanen = data.nasib;
          updatedRows[i].dasarKlasifikasi = data.dasar;
          updatedRows[i].status = 'completed';
        } else {
          updatedRows[i].status = 'error';
        }
      } catch (err) {
        console.error(err);
        updatedRows[i].status = 'error';
      }
      setRows([...updatedRows]);
      setGeminiStatus(`Memproses baris ${i + 1} dari ${updatedRows.length}...`);
    }

    setIsProcessing(false);
    setGeminiStatus('Proses selesai');
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
    XLSX.writeFile(wb, "Hasil_Klasifikasi_Arsip_Undip.xlsx");
  };

  const stats = useMemo(() => {
    const total = rows.length;
    const completed = rows.filter(r => r.status === 'completed').length;
    const permanent = rows.filter(r => r.musnahPermanen === 'Permanen').length;
    const destroyed = rows.filter(r => r.musnahPermanen === 'Musnah').length;
    return { total, completed, permanent, destroyed };
  }, [rows]);

  return (
    <div className="min-h-screen pb-24 bg-slate-50">
      {/* Header Section */}
      <header className="undip-gradient py-16 px-8 text-white relative overflow-hidden shadow-2xl">
        <div className="absolute top-0 left-0 w-full h-full opacity-10 pointer-events-none">
          <div className="absolute inset-0 bg-[radial-gradient(circle_at_center,_var(--color-undip-gold)_0%,_transparent_70%)]"></div>
        </div>
        
        <div className="max-w-6xl mx-auto flex flex-col items-center text-center relative z-10">
          <motion.div 
            initial={{ scale: 0.8, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="w-20 h-20 bg-undip-gold rounded-3xl flex items-center justify-center shadow-2xl mb-6 transform -rotate-6"
          >
            <Archive className="w-10 h-10 text-undip-blue" />
          </motion.div>
          <motion.h1 
            initial={{ opacity: 0, y: 10 }}
            animate={{ opacity: 1, y: 0 }}
            className="text-5xl font-black tracking-tighter sm:text-6xl mb-4"
          >
            UNDIP ARCHIVE <span className="text-undip-gold">ASSISTANT</span>
          </motion.h1>
          <motion.p 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            transition={{ delay: 0.2 }}
            className="text-blue-100 text-lg max-w-2xl font-medium"
          >
            Solusi cerdas klasifikasi arsip universitas dengan teknologi Generative AI untuk akurasi dan efisiensi pengarsipan.
          </motion.p>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-6 -mt-12 grid grid-cols-1 lg:grid-cols-12 gap-8 relative z-20">
        {/* Configuration Sidebar */}
        <aside className="lg:col-span-4 space-y-6">
          <section className="glass-card p-6 rounded-3xl shadow-xl shadow-slate-200/50">
            <h2 className="text-xl font-black text-undip-blue mb-6 flex items-center gap-2">
              <Upload className="w-6 h-6" /> DATA MASUKAN
            </h2>
            
            <div className="space-y-4">
              <div {...getJKProps()} className={`p-5 border-2 border-dashed rounded-2xl cursor-pointer transition-all flex items-center gap-4 ${jkActive ? 'border-undip-blue bg-blue-50' : 'border-slate-200 hover:border-undip-blue/40'}`}>
                <input {...getJKInputProps()} />
                <div className={`p-3 rounded-xl ${refJK.loaded ? 'bg-green-100 text-green-600' : 'bg-slate-100 text-slate-400'}`}>
                  <FileText className="w-6 h-6" />
                </div>
                <div className="flex-1 overflow-hidden">
                  <p className="text-sm font-bold text-slate-700">SK Klasifikasi</p>
                  <p className="text-xs text-slate-400 truncate font-mono uppercase tracking-tighter">
                    {refJK.name || 'Klik / Tarik PDF di sini'}
                  </p>
                </div>
              </div>

              <div {...getJRAProps()} className={`p-5 border-2 border-dashed rounded-2xl cursor-pointer transition-all flex items-center gap-4 ${jraActive ? 'border-undip-blue bg-blue-50' : 'border-slate-200 hover:border-undip-blue/40'}`}>
                <input {...getJRAInputProps()} />
                <div className={`p-3 rounded-xl ${refJRA.loaded ? 'bg-green-100 text-green-600' : 'bg-slate-100 text-slate-400'}`}>
                  <BookOpen className="w-6 h-6" />
                </div>
                <div className="flex-1 overflow-hidden">
                  <p className="text-sm font-bold text-slate-700">SK JRA</p>
                  <p className="text-xs text-slate-400 truncate font-mono uppercase tracking-tighter">
                    {refJRA.name || 'Klik / Tarik PDF di sini'}
                  </p>
                </div>
              </div>

              <div {...getExcelProps()} className={`p-5 border-2 border-dashed rounded-2xl cursor-pointer transition-all flex items-center gap-4 ${excelActive ? 'border-undip-blue bg-blue-50' : 'border-slate-200 hover:border-undip-blue/40'}`}>
                <input {...getExcelInputProps()} />
                <div className={`p-3 rounded-xl ${rows.length > 0 ? 'bg-green-100 text-green-600' : 'bg-slate-100 text-slate-400'}`}>
                  <TableIcon className="w-6 h-6" />
                </div>
                <div className="flex-1 overflow-hidden">
                  <p className="text-sm font-bold text-slate-700">Daftar Arsip</p>
                  <p className="text-xs text-slate-400 truncate font-mono uppercase tracking-tighter">
                    {rows.length > 0 ? `${rows.length} Baris Terdeteksi` : 'Pilih file Excel (.xlsx)'}
                  </p>
                </div>
              </div>
            </div>

            {error && (
              <motion.div 
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                className="mt-6 p-4 bg-red-50 text-red-700 rounded-2xl flex items-start gap-3 text-sm border border-red-100"
              >
                <AlertCircle className="w-5 h-5 shrink-0 mt-0.5" />
                <p className="font-medium">{error}</p>
              </motion.div>
            )}

            <button 
              onClick={startClassification}
              disabled={isProcessing || !refJK.loaded || !refJRA.loaded || rows.length === 0}
              className="w-full mt-8 bg-undip-blue hover:bg-undip-blue-light disabled:bg-slate-100 disabled:text-slate-400 text-white font-black py-4 rounded-2xl transition-all shadow-xl shadow-undip-blue/20 hover:shadow-undip-blue/30 flex items-center justify-center gap-3 active:scale-95"
            >
              {isProcessing ? <Loader2 className="w-6 h-6 animate-spin" /> : <Play className="w-6 h-6 fill-current" />}
              PROSES SEKARANG
            </button>
          </section>

          {/* Stats Bar */}
          {rows.length > 0 && (
            <section className="glass-card p-6 rounded-3xl grid grid-cols-2 gap-4">
              <div className="col-span-2">
                <h3 className="text-xs font-black text-slate-400 uppercase tracking-widest mb-4">Laporan Pemrosesan</h3>
                <div className="w-full h-4 bg-slate-100 rounded-full overflow-hidden mb-6 flex">
                  <div 
                    className="h-full bg-green-500 transition-all duration-500"
                    style={{ width: `${(stats.completed / stats.total) * 100}%` }}
                  ></div>
                  <div 
                    className="h-full bg-undip-blue animate-pulse transition-all duration-300"
                    style={{ width: `${(rows.filter(r => r.status === 'processing').length / stats.total) * 100}%` }}
                  ></div>
                </div>
              </div>
              <div className="p-4 bg-slate-50 rounded-2xl">
                <p className="text-[10px] font-black text-slate-400 uppercase">Selesai</p>
                <p className="text-2xl font-black text-slate-800">{stats.completed}</p>
              </div>
              <div className="p-4 bg-slate-50 rounded-2xl text-right">
                <p className="text-[10px] font-black text-slate-400 uppercase">Input</p>
                <p className="text-2xl font-black text-slate-800">{stats.total}</p>
              </div>
            </section>
          )}

          {stats.completed > 0 && (
            <motion.button 
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              onClick={exportToExcel}
              className="w-full bg-green-600 hover:bg-green-700 text-white font-black py-4 rounded-2xl transition-all shadow-xl shadow-green-600/20 flex items-center justify-center gap-3"
            >
              <Download className="w-5 h-5" /> UNDUH HASIL
            </motion.button>
          )}
        </aside>

        {/* Preview Content */}
        <section className="lg:col-span-8 space-y-6">
          <div className="flex items-center justify-between">
            <h2 className="text-2xl font-black text-undip-blue uppercase tracking-tighter underline decoration-undip-gold decoration-4 underline-offset-8">
              DAFTAR KERJA ARSIP
            </h2>
            {geminiStatus && (
              <div className="flex items-center gap-3 px-4 py-2 bg-undip-blue-dark text-undip-gold rounded-full text-[10px] font-black uppercase tracking-widest animate-fade-in shadow-lg">
                <Loader2 className="w-3 h-3 animate-spin"/>
                {geminiStatus}
              </div>
            )}
          </div>

          <div className="glass-card rounded-[2rem] overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-slate-50/50 border-b border-slate-100">
                    <th className="p-6 text-[10px] font-black text-slate-400 uppercase tracking-widest w-16">#</th>
                    <th className="p-6 text-[10px] font-black text-slate-400 uppercase tracking-widest">Informasi Arsip</th>
                    <th className="p-6 text-[10px] font-black text-slate-400 uppercase tracking-widest w-40 text-center">Hasil AI</th>
                    <th className="p-6 text-[10px] font-black text-slate-400 uppercase tracking-widest w-24 text-right">Aksi</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 font-sans">
                  {rows.length === 0 ? (
                    <tr>
                      <td colSpan={4} className="p-24 text-center ">
                        <div className="flex flex-col items-center gap-6 opacity-30">
                          <div className="p-8 bg-slate-100 rounded-full">
                            <Archive className="w-16 h-16 text-slate-400" />
                          </div>
                          <p className="font-bold text-slate-500 max-w-xs uppercase text-sm tracking-widest leading-loose">
                            Menunggu unggahan berkas untuk memulai analisis
                          </p>
                        </div>
                      </td>
                    </tr>
                  ) : (
                    rows.map((row) => (
                      <motion.tr 
                        key={row.no}
                        layout
                        className={`group transition-all ${row.status === 'processing' ? 'bg-blue-50/40 relative z-10' : 'hover:bg-slate-50/50'}`}
                      >
                        <td className="p-6">
                           <div className={`w-8 h-8 rounded-lg flex items-center justify-center font-black text-xs ${
                             row.status === 'completed' ? 'bg-green-100 text-green-700' : 'bg-slate-100 text-slate-400'
                           }`}>
                             {row.no}
                           </div>
                        </td>
                        <td className="p-6">
                           <div className="max-w-md">
                              <p className="text-sm font-bold text-slate-800 leading-tight mb-2">
                                {row.uraian}
                              </p>
                              <div className="flex items-center gap-4">
                                <span className="flex items-center gap-1 text-[10px] font-black text-slate-400 uppercase">
                                  <CalendarDays className="w-3 h-3"/> {row.tahun}
                                </span>
                                {row.dasarKlasifikasi && (
                                  <span className="text-[9px] font-bold text-undip-blue bg-blue-50 px-2 py-0.5 rounded-md italic">
                                    "{row.dasarKlasifikasi}"
                                  </span>
                                )}
                              </div>
                           </div>
                        </td>
                        <td className="p-6 text-center">
                           <div className="flex flex-col items-center gap-2">
                              <span className={`font-mono text-xs font-black px-3 py-1 rounded-lg ${
                                row.kodeKlasifikasi ? 'bg-undip-blue text-white' : 'bg-slate-100 text-slate-300'
                              }`}>
                                {row.kodeKlasifikasi || '---'}
                              </span>
                              {row.musnahPermanen && (
                                <span className={`text-[8px] font-black px-2 py-0.5 rounded-md uppercase border-2 ${
                                  row.musnahPermanen === 'Permanen' 
                                    ? 'border-indigo-600 text-indigo-600' 
                                    : 'border-orange-500 text-orange-500'
                                }`}>
                                  {row.musnahPermanen}
                                </span>
                              )}
                           </div>
                        </td>
                        <td className="p-6 text-right">
                           {row.status === 'completed' ? (
                             <CheckCircle2 className="w-6 h-6 text-green-500 ml-auto" />
                           ) : row.status === 'processing' ? (
                             <Loader2 className="w-6 h-6 text-undip-blue animate-spin ml-auto" />
                           ) : row.status === 'error' ? (
                             <AlertCircle className="w-6 h-6 text-red-500 ml-auto" />
                           ) : (
                             <div className="w-6 h-6 rounded-full border-2 border-slate-100 ml-auto group-hover:border-slate-200" />
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

      <footer className="fixed bottom-0 left-0 right-0 bg-white/80 backdrop-blur-xl border-t border-slate-100 py-4 px-10 flex justify-between items-center z-50">
        <div className="flex items-center gap-6">
           <div className="flex items-center gap-2">
              <div className="w-6 h-6 bg-undip-blue rounded flex items-center justify-center">
                 <span className="text-white text-[10px] font-black">U</span>
              </div>
              <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest">
                Universitas Diponegoro
              </span>
           </div>
           <div className="h-4 w-px bg-slate-200"></div>
           <p className="text-[10px] font-bold text-slate-400 uppercase tracking-[0.2em]">
             Pusat Arsip dan Dokumentasi &copy; 2025
           </p>
        </div>
        <div className="flex items-center gap-8">
           <div className="flex items-center gap-2 px-3 py-1.5 bg-slate-50 rounded-full border border-slate-100 group cursor-help">
              <span className="text-[9px] font-black text-slate-400 uppercase group-hover:text-undip-blue transition-colors">v1.2.0 Stable</span>
           </div>
           <div className="flex items-center gap-2 text-undip-blue hover:text-undip-blue-light transition-colors cursor-pointer group">
              <span className="text-[10px] font-black uppercase tracking-widest underline decoration-2 underline-offset-4 decoration-undip-gold/40 group-hover:decoration-undip-gold">Panduan Teknis</span>
              <ExternalLink className="w-3 h-3" />
           </div>
        </div>
      </footer>
    </div>
  );
}
