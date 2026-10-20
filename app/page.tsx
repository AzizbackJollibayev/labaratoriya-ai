"use client";

import React, { useEffect, useRef, useState } from "react";

const MAX_MB = 15;

function formatSize(bytes: number) {
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

export default function LaborantPage() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [file, setFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [downloadName, setDownloadName] = useState("");
  const [error, setError] = useState("");
  const [isDragging, setIsDragging] = useState(false);

  // Eski blob URL xotirada qolib ketmasligi uchun tozalaymiz
  useEffect(() => {
    return () => {
      if (downloadUrl) window.URL.revokeObjectURL(downloadUrl);
    };
  }, [downloadUrl]);

  const acceptFile = (f: File | undefined | null) => {
    if (!f) return;
    if (!f.name.toLowerCase().endsWith(".docx")) {
      setError("Faqat .docx formatidagi Word fayllarini yuklashingiz mumkin!");
      return;
    }
    if (f.size > MAX_MB * 1024 * 1024) {
      setError(`Fayl hajmi ${MAX_MB} MB dan oshmasligi kerak.`);
      return;
    }
    setFile(f);
    setError("");
    setDownloadUrl(null);
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    acceptFile(e.target.files?.[0]);
    e.target.value = ""; // bir xil faylni qayta tanlash mumkin bo'lsin
  };

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    // Ichki elementlar ustiga o'tganda "miltillamasligi" uchun
    if (!e.currentTarget.contains(e.relatedTarget as Node)) setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    acceptFile(e.dataTransfer.files?.[0]);
  };

  const removeFile = () => {
    setFile(null);
    setDownloadUrl(null);
    setError("");
  };

  const handleProcess = async () => {
    if (!file) {
      setError("Iltimos, avval laboratoriya faylini yuklang!");
      return;
    }

    setLoading(true);
    setError("");
    setDownloadUrl(null);

    const formData = new FormData();
    formData.append("file", file);

    try {
      const response = await fetch("/api/process-lab", { method: "POST", body: formData });

      if (!response.ok) {
        const errData = await response.json().catch(() => ({}));
        throw new Error(errData.error || "Tahlil qilishda xatolik yuz berdi");
      }

      const blob = await response.blob();
      setDownloadUrl(window.URL.createObjectURL(blob));
      setDownloadName(`TAHLIL_${file.name}`);
    } catch (err: any) {
      setError(err?.message || "Noma'lum xatolik yuz berdi");
    } finally {
      setLoading(false);
    }
  };

  return (
    <main className="lab-wrap">
      <header className="lab-header">
        <div className="lab-logo" aria-hidden="true">
          <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M9 3h6M10 3v6l-5 9a2 2 0 0 0 1.8 3h10.4a2 2 0 0 0 1.8-3l-5-9V3" />
            <path d="M7.5 15h9" />
          </svg>
        </div>
        <h1>Laboratoriya AI Tahlil Tizimi</h1>
        <p>Faylni yuklang — AI tahlil izohi qo&apos;shilgan tayyor hujjatni bemorga yuboring</p>
      </header>

      <ol className="lab-steps" aria-label="Ish tartibi">
        <li className={file ? "done" : "active"}><span>1</span> Yuklash</li>
        <li className={downloadUrl ? "done" : file ? "active" : ""}><span>2</span> AI tahlil</li>
        <li className={downloadUrl ? "active" : ""}><span>3</span> Yuklab olish</li>
      </ol>

      <section className="lab-card">
        <div
          className={`dropzone${isDragging ? " drag" : ""}`}
          role="button"
          tabIndex={0}
          aria-label="Word faylini tanlash yoki shu yerga tashlash"
          onClick={() => inputRef.current?.click()}
          onKeyDown={(e) => {
            if (e.key === "Enter" || e.key === " ") {
              e.preventDefault();
              inputRef.current?.click();
            }
          }}
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
        >
          <input ref={inputRef} type="file" accept=".docx" onChange={handleFileChange} hidden />
          <svg className="dz-icon" width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
            <path d="M12 16V4M7 9l5-5 5 5" />
            <path d="M4 16v3a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-3" />
          </svg>
          <strong>Faylni tanlash (.docx)</strong>
          <span className="dz-sub">yoki faylni shu yerga sudrab kelib tashlang · maks. {MAX_MB} MB</span>
        </div>

        {file && (
          <div className="file-chip">
            <div className="file-ico" aria-hidden="true">W</div>
            <div className="file-meta">
              <span className="file-name" title={file.name}>{file.name}</span>
              <span className="file-size">{formatSize(file.size)}</span>
            </div>
            <button type="button" className="chip-x" onClick={removeFile} disabled={loading} aria-label="Faylni olib tashlash">
              ✕
            </button>
          </div>
        )}

        <button className="btn-primary" onClick={handleProcess} disabled={loading || !file}>
          {loading ? (
            <>
              <span className="spinner" aria-hidden="true" /> AI tahlil qilmoqda...
            </>
          ) : (
            "Tahlil qilish va faylni tayyorlash"
          )}
        </button>
      </section>

      {error && (
        <div className="alert-error" role="alert">
          {error}
        </div>
      )}

      {downloadUrl && (
        <div className="result" role="status">
          <div className="result-ico" aria-hidden="true">✓</div>
          <p>Fayl tayyor! AI xulosasi fayl oxiriga biriktirildi.</p>
          <a className="btn-success" href={downloadUrl} download={downloadName}>
            Tayyor faylni yuklab olish (bemor uchun)
          </a>
        </div>
      )}

      <p className="lab-foot">AI xulosasi tashxis emas — yakuniy xulosani shifokor beradi.</p>
    </main>
  );
}
