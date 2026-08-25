"use client";

import React, { useState } from "react";

export default function LaborantPage() {
  const [file, setFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [downloadName, setDownloadName] = useState("");
  const [error, setError] = useState("");
  const [isDragging, setIsDragging] = useState(false);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setFile(e.target.files[0]);
      setError("");
      setDownloadUrl(null);
    }
  };

  // Drag & Drop (Faylni sudrab kelib tashlash) hodisalari
  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const droppedFile = e.dataTransfer.files[0];
      if (droppedFile.name.endsWith(".docx")) {
        setFile(droppedFile);
        setError("");
        setDownloadUrl(null);
      } else {
        setError("Faqat .docx formatidagi Word fayllarini yuklashingiz mumkin!");
      }
    }
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
      const response = await fetch("/api/process-lab", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const errData = await response.json();
        throw new Error(errData.error || "Tahlil qilishda xatolik yuz berdi");
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      setDownloadUrl(url);
      setDownloadName(`TAHLIL_${file.name}`);
    } catch (err: any) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{ maxWidth: "650px", margin: "60px auto", fontFamily: "sans-serif", padding: "0 20px" }}>
      <div style={{ textAlign: "center", marginBottom: "30px" }}>
        <h1 style={{ color: "#0f172a", marginBottom: "8px" }}>🩺 Laboratoriya AI Tahlil Tizimi</h1>
        <p style={{ color: "#64748b", margin: 0 }}>Faylni yuklang, AI tahlil qo'shilgan tayyor hujjatni bemorga yuboring</p>
      </div>

      <div
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        style={{
          background: isDragging ? "#eff6ff" : "#ffffff",
          border: isDragging ? "2px dashed #2563eb" : "2px dashed #cbd5e1",
          borderRadius: "16px",
          padding: "40px 20px",
          textAlign: "center",
          transition: "all 0.2s ease",
        }}
      >
        <input
          type="file"
          accept=".docx"
          id="fileInput"
          onChange={handleFileChange}
          style={{ display: "none" }}
        />
        <label
          htmlFor="fileInput"
          style={{
            display: "inline-block",
            padding: "12px 24px",
            background: "#f1f5f9",
            borderRadius: "8px",
            cursor: "pointer",
            fontWeight: "600",
            color: "#334155",
            marginBottom: "8px",
          }}
        >
          📁 Faylni tanlash (.docx)
        </label>
        
        <p style={{ color: "#94a3b8", fontSize: "14px", margin: "8px 0" }}>
          yoki faylni shu yerga sudrab kelib tashlang
        </p>

        {file && <p style={{ color: "#2563eb", fontWeight: "600", marginTop: "12px" }}>Tanlandi: {file.name}</p>}

        <div style={{ marginTop: "20px" }}>
          <button
            onClick={handleProcess}
            disabled={loading || !file}
            style={{
              width: "100%",
              padding: "14px",
              background: loading ? "#94a3b8" : "#2563eb",
              color: "#fff",
              border: "none",
              borderRadius: "10px",
              fontSize: "16px",
              fontWeight: "bold",
              cursor: loading || !file ? "not-allowed" : "pointer",
            }}
          >
            {loading ? "Gemini AI tahlil qilmoqda..." : "🚀 Tahlil qilish va Faylni Tayyorlash"}
          </button>
        </div>
      </div>

      {error && (
        <div style={{ background: "#fef2f2", color: "#dc2626", padding: "14px", borderRadius: "10px", marginTop: "20px" }}>
          ⚠️ {error}
        </div>
      )}

      {downloadUrl && (
        <div style={{ background: "#f0fdf4", border: "1px solid #bbf7d0", padding: "20px", borderRadius: "12px", marginTop: "24px", textAlign: "center" }}>
          <p style={{ color: "#166534", fontWeight: "bold", marginBottom: "12px" }}>
            ✅ Fayl tayyorlandi! Gemini AI xulosasi fayl oxiriga biriktirildi.
          </p>
          <a
            href={downloadUrl}
            download={downloadName}
            style={{
              display: "inline-block",
              padding: "12px 28px",
              background: "#16a34a",
              color: "#fff",
              textDecoration: "none",
              borderRadius: "8px",
              fontWeight: "bold",
            }}
          >
            📥 Tayyor Faylni Yuklab Olish (Bemor uchun)
          </a>
        </div>
      )}
    </div>
  );
}
