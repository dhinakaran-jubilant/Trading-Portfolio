import Layout from './Layout';
import React, { useState, useRef } from 'react';
import axios from 'axios';
import { DotLottiePlayer } from '@dotlottie/react-player';
import config from './config';

function Upload({ user, onLogout }) {
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [status, setStatus] = useState(null);
  const [dragActive, setDragActive] = useState(false);

  const [showSuccessModal, setShowSuccessModal] = useState(false);
  const [successMessage, setSuccessMessage] = useState('');
  const [processingTime, setProcessingTime] = useState(0);

  const [selectedName, setSelectedName] = useState('Gautham');
  const [selectedBank, setSelectedBank] = useState('HDFC');

  const fileInputRef = useRef(null);

  const isValidExcel = (f) => {
    return f.name.endsWith('.xlsx') || f.name.endsWith('.xls') || f.name.endsWith('.csv') ||
      [
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "application/vnd.ms-excel",
        "text/csv",
        "application/csv"
      ].includes(f.type);
  };

  const handleFiles = (files) => {
    const fileList = Array.from(files);
    if (fileList.length > 0 && isValidExcel(fileList[0])) {
      setFile(fileList[0]);
    }
    setStatus(null);
  };

  const handleDrag = (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  const handleDragOver = (e) => {
    e.preventDefault();
    e.stopPropagation();
  };

  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFiles(e.dataTransfer.files);
    }
  };

  const handleFileChange = (e) => {
    if (e.target.files && e.target.files.length > 0) {
      handleFiles(e.target.files);
    }
  };

  const handleUploadClick = () => {
    fileInputRef.current.click();
  };

  const removeFile = () => {
    setFile(null);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const clearAll = () => {
    setFile(null);
    setStatus(null);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const handleSubmit = async () => {
    if (!file) {
      setStatus({ type: 'error', message: 'Please select an Excel or CSV file first.' });
      return;
    }

    setLoading(true);
    setStatus(null);
    const startTime = performance.now();

    const formData = new FormData();
    formData.append('file', file);
    formData.append('name', selectedName);
    formData.append('bank', selectedBank);

    try {
      const response = await axios.post(`${config.API_BASE_URL}/process-excel`, formData, {
        responseType: 'blob',
        headers: {
          'Content-Type': 'multipart/form-data',
        },
      });

      const endTime = performance.now();
      const duration = ((endTime - startTime) / 1000).toFixed(2);
      setProcessingTime(duration);

      const dateStr = new Date().toLocaleDateString('en-GB').replace(/\//g, '.');
      const downloadName = `${selectedName}_${selectedBank}_${dateStr}.xlsx`;

      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', downloadName);
      document.body.appendChild(link);
      link.click();

      link.parentNode.removeChild(link);
      window.URL.revokeObjectURL(url);

      setSuccessMessage('Your Portfolio Summary Report has been processed and is ready for use.');
      setShowSuccessModal(true);
      clearAll();
    } catch (err) {
      console.error(err);
      let errorMsg = 'Failed to process files. Ensure backend is running.';

      if (err.response && err.response.data && err.response.data.type === 'application/json') {
        try {
          const text = await new Response(err.response.data).text();
          const json = JSON.parse(text);
          if (json.error) {
            errorMsg = json.error;
          }
        } catch (e) { }
      }

      setStatus({ type: 'error', message: errorMsg });
    } finally {
      setLoading(false);
    }
  };

  const isProcessing = loading;
  const errorMessage = status?.type === 'error' ? status.message : null;
  const dismissError = () => setStatus(null);

  const selectedFiles = file ? [{ name: file.name, size: file.size, type: file.name.endsWith('.csv') ? 'CSV' : 'EXCEL' }] : [];

  return (
    <Layout user={user} onLogout={onLogout} activeTab="portfolio" breadcrumbs={['Summary']}>
      <main className="flex-1 bg-slate-50 dark:bg-[#101822] px-10 pb-8 transition-colors duration-300">
        <div className="flex flex-col">
          {/* Header Section */}
          <div className="flex flex-col items-start text-left gap-2 mb-10">
            <h1 className="text-4xl font-black tracking-tight text-slate-900 dark:text-white transition-colors duration-300">Portfolio Summary Processor</h1>
            <p className="text-slate-600 dark:text-slate-400 text-sm leading-relaxed transition-colors duration-300">Upload your portfolio equity summary Excel or CSV file. The system automatically maps columns and calculates Day P&L to generate a clean, formatted report in the HDFC style.</p>
          </div>

          {/* Main Upload Card Container */}
          <div className="bg-white dark:bg-[#0f172b] rounded-3xl shadow-2xl border border-slate-200 dark:border-slate-800 overflow-hidden transition-colors duration-300 p-8">

            {/* Selection Section */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8">
              <div className="flex flex-col gap-3">
                <label className="text-[11px] font-black uppercase tracking-[0.2em] text-slate-500 ml-1">Report For</label>
                <div className="relative">
                  <select
                    value={selectedName}
                    onChange={(e) => {
                      const val = e.target.value;
                      setSelectedName(val);
                      if (val === 'Sudhakar') setSelectedBank('ICICI');
                    }}
                    className="w-full appearance-none px-6 py-4 bg-slate-50 dark:bg-[#111926] border border-slate-200 dark:border-slate-800 rounded-2xl font-bold text-slate-900 dark:text-white focus:ring-2 focus:ring-blue-500/50 transition-all outline-none"
                  >
                    <option value="Sudhakar">Sudhakar</option>
                    <option value="Gautham">Gautham</option>
                  </select>
                  <span className="material-symbols-outlined absolute right-5 top-1/2 -translate-y-1/2 pointer-events-none text-slate-400">expand_more</span>
                </div>
              </div>
              <div className="flex flex-col gap-3">
                <label className="text-[11px] font-black uppercase tracking-[0.2em] text-slate-500 ml-1">Select Bank</label>
                <div className="relative">
                  <select
                    value={selectedBank}
                    onChange={(e) => setSelectedBank(e.target.value)}
                    className="w-full appearance-none px-6 py-4 bg-slate-50 dark:bg-[#111926] border border-slate-200 dark:border-slate-800 rounded-2xl font-bold text-slate-900 dark:text-white focus:ring-2 focus:ring-blue-500/50 transition-all outline-none"
                  >
                    <option value="ICICI">ICICI</option>
                    {selectedName === 'Gautham' && <option value="HDFC">HDFC</option>}
                  </select>
                  <span className="material-symbols-outlined absolute right-5 top-1/2 -translate-y-1/2 pointer-events-none text-slate-400">expand_more</span>
                </div>
              </div>
            </div>

            {/* Unified Drop Zone */}
            <div
              className={`flex flex-col items-center justify-center gap-8 rounded-2xl border-2 border-dashed transition-all duration-300 p-10 cursor-pointer group ${dragActive ? 'border-blue-500 bg-blue-50 dark:bg-blue-500/5' : 'border-slate-300 dark:border-slate-800 bg-slate-50 dark:bg-[#16223b] hover:border-slate-400 dark:hover:border-slate-700 hover:bg-slate-100 dark:hover:bg-[#1a2846]'}`}
              onClick={handleUploadClick}
              onDragEnter={handleDrag}
              onDragLeave={handleDrag}
              onDragOver={handleDragOver}
              onDrop={handleDrop}
            >
              <input
                type="file"
                ref={fileInputRef}
                className="hidden"
                accept=".xlsx, .xls, .csv"
                multiple
                onChange={handleFileChange}
              />
              <div className="flex flex-col items-center gap-5">
                <div className="w-15 h-15 bg-white dark:bg-[#111926] border border-slate-200 dark:border-slate-800 rounded-full flex items-center justify-center shadow-xl group-hover:scale-110 transition-transform">
                  <span className="material-symbols-outlined text-5xl text-blue-500">upload_file</span>
                </div>
                <div className="text-center">
                  <h4 className="text-2xl font-bold text-slate-900 dark:text-white mb-2 transition-colors duration-300">Drag and drop files here</h4>
                  <p className="text-slate-500 text-sm font-medium">Supports Excel and CSV files.</p>
                </div>
              </div>
              <button className="flex items-center gap-3 px-10 py-4 bg-blue-600 text-white rounded-xl font-bold shadow-xl shadow-blue-600/20 hover:bg-blue-700 transition-all active:scale-95 cursor-pointer">
                <span className="material-symbols-outlined text-xl">add_circle</span>
                Browse Files
              </button>
            </div>

            {/* Inline Error Message */}
            {errorMessage && (
              <div className="mt-6 bg-red-500/10 border border-red-500/20 text-red-500 px-6 py-4 rounded-xl flex items-center gap-4 animate-in fade-in slide-in-from-top-2">
                <span className="material-symbols-outlined text-2xl">error_outline</span>
                <span className="text-[12px] font-bold flex-1 tracking-widest">{errorMessage}</span>
                <button
                  onClick={dismissError}
                  className="flex items-center justify-center h-8 w-8 rounded-lg hover:bg-red-500/20 transition-colors"
                >
                  <span className="material-symbols-outlined text-xl">close</span>
                </button>
              </div>
            )}

            {/* Selected Files List */}
            {selectedFiles.length > 0 && (
              <div className="mt-10 flex flex-col gap-4">
                <h5 className="font-black text-[11px] uppercase tracking-[0.2em] text-slate-500 ml-1">Selected Files:</h5>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {selectedFiles.map((f, index) => (
                    <div key={index} className="flex items-center justify-between p-5 rounded-2xl border border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-[#111926]/50 transition-colors duration-300 animate-in zoom-in-95 duration-300">
                      <div className="flex items-center gap-4 overflow-hidden">
                        <div className="w-10 h-10 rounded-lg flex items-center justify-center bg-blue-100 text-blue-600">
                          <span className="material-symbols-outlined">description</span>
                        </div>
                        <div className="flex flex-col min-w-0">
                          <div className="flex items-center gap-2">
                            <span className="text-[10px] font-black uppercase tracking-widest text-slate-500">Portfolio Report</span>
                          </div>
                          <span className="text-sm font-bold text-slate-900 dark:text-white truncate transition-colors duration-300" title={f.name}>{f.name}</span>
                          <span className="text-[10px] font-black uppercase tracking-widest text-slate-400">
                            {(f.size / 1024 / 1024).toFixed(2)} MB
                          </span>
                        </div>
                      </div>
                      <button
                        onClick={() => removeFile()}
                        className="flex items-center justify-center rounded-xl h-10 w-10 text-slate-400 hover:text-red-500 hover:bg-red-50 dark:hover:bg-red-500/10 transition-all cursor-pointer"
                      >
                        <span className="material-symbols-outlined">delete</span>
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Submit Button Section */}
            <div className="mt-10 flex justify-end border-t border-slate-100 dark:border-slate-800 pt-8">
              <button
                onClick={handleSubmit}
                disabled={!file || isProcessing}
                className={`flex items-center gap-4 px-12 py-5 bg-blue-600 text-white rounded-2xl font-black uppercase tracking-[0.2em] text-[12px] shadow-2xl shadow-blue-600/30 transition-all ${isProcessing || !file ? 'opacity-50 cursor-not-allowed scale-95' : 'hover:bg-blue-700 hover:-translate-y-1 active:scale-95 cursor-pointer'}`}
              >
                {isProcessing ? (
                  <>
                    <div className="w-5 h-5 border-2 border-white/30 border-t-white rounded-full animate-spin" />
                    <span>Processing...</span>
                  </>
                ) : (
                  <>
                    <span className="material-symbols-outlined">auto_awesome</span>
                    <span>Generate Report</span>
                  </>
                )}
              </button>
            </div>
          </div>
        </div>

        {/* Footer */}
        <footer className="mt-20 pt-10 border-t border-slate-200 dark:border-slate-800/50 text-center text-[12px] font-bold tracking-[0.1em] text-slate-500">
          <p>
            All rights reserved &copy; 2026 @ Ardent Capital. Designed and Developed by{' '}
            <a href="mailto:dhinakaran.s@jubilantenterprises.in" className="text-blue-500 hover:underline">
              Dhinakaran Sekar
            </a>
          </p>
        </footer>
      </main>

      {/* Processing Modal Overlay */}
      {isProcessing && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-slate-900/40 dark:bg-slate-900/60 backdrop-blur-sm animate-in fade-in duration-300">
          <div className="bg-white dark:bg-[#0a0f18] rounded-3xl shadow-2xl p-12 flex flex-col items-center gap-8 max-w-sm w-full mx-4 border border-slate-200 dark:border-slate-800 animate-in zoom-in-95 duration-300">
            <div className="size-32">
              <DotLottiePlayer
                src="https://lottie.host/26c29f4c-00d1-4c13-a161-b608499b7ccb/5uUeWZYlLT.lottie"
                autoplay
                loop
                className="size-full"
              />
            </div>
            <div className="text-center">
              <h3 className="text-2xl font-black tracking-tight text-slate-900 dark:text-white mb-2 transition-colors duration-300">Processing Reports</h3>
              <p className="text-slate-600 dark:text-slate-400 text-sm transition-colors duration-300">Please wait, this may take a few moments...</p>
            </div>
          </div>
        </div>
      )}

      {/* Success Modal Overlay */}
      {showSuccessModal && (
        <div className="fixed inset-0 z-[110] flex items-center justify-center bg-slate-900/40 dark:bg-black/60 backdrop-blur-sm p-4 animate-in fade-in duration-200">
          <div className="bg-white dark:bg-[#0a0f18] rounded-3xl shadow-2xl p-12 flex flex-col items-center max-w-sm w-full mx-4 border border-slate-200 dark:border-slate-800 animate-in zoom-in-95 duration-200 text-center">
            <div className="w-15 h-15 rounded-full bg-emerald-50 dark:bg-emerald-600/10 text-emerald-500 flex items-center justify-center mb-8 border border-emerald-200 dark:border-emerald-500/20">
              <span className="material-symbols-outlined text-[30px]">check_circle</span>
            </div>
            <h3 className="text-2xl font-black tracking-tight text-slate-900 dark:text-white mb-2 transition-colors duration-300">Process Complete!</h3>
            <p className="text-slate-600 dark:text-slate-400 text-sm mb-4 leading-relaxed transition-colors duration-300">
              {successMessage || 'Your report has been generated and downloaded successfully.'}
            </p>
            <div className="bg-emerald-500/10 dark:bg-emerald-500/5 px-6 py-3 rounded-2xl border border-emerald-500/20 mb-10 flex flex-col items-center gap-1 shadow-inner">
              <span className="text-[10px] font-black uppercase tracking-[0.2em] text-emerald-600/60 dark:text-emerald-400/40">Efficiency Score</span>
              <div className="flex items-center gap-2">
                <span className="material-symbols-outlined text-emerald-500 text-sm">bolt</span>
                <span className="text-lg font-black text-emerald-600 dark:text-emerald-400 leading-none">{processingTime}s</span>
                <span className="text-[10px] font-bold text-emerald-500/60 mt-1 uppercase tracking-widest">Processing Time</span>
              </div>
            </div>
            <button
              onClick={() => setShowSuccessModal(false)}
              className="w-full py-5 bg-blue-600 hover:bg-blue-700 text-white rounded-2xl font-black uppercase tracking-[0.2em] text-[12px] transition-all cursor-pointer shadow-2xl shadow-blue-600/30 active:scale-95"
            >
              Close and Continue
            </button>
          </div>
        </div>
      )}
    </Layout>
  );
}

export default Upload;
