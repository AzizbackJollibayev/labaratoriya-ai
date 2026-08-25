"client"; // agar page boshida client directive bo'lsa saqlang, yo'qsa shundoq tashlang

import { useState } from "react";

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string>("");

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setFile(e.target.files[0]);
      setDownloadUrl(null);
    }
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!file) return;

    setLoading(true);
    setDownloadUrl(null);

    const formData = new FormData();
    formData.append("file", file);

    try {
      const response = await fetch("/api/analyze", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const errData = await response.json();
        alert(errData.error || "Xatolik yuz berdi");
        setLoading(false);
        return;
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      setDownloadUrl(url);
      setFileName(`TAHLIL_${file.name}`);
    } catch (err) {
      alert("Tarmoqda xatolik yuz berdi");
    } finally {
      setLoading(false);
    }
  };

  return (
    <main style={{ minHeight: "100vh", backgroundColor: "#f3f4f6", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: "20px", fontFamily: "sans-serif" }}>
      <div style={{ backgroundColor: "white", padding: "40px", borderRadius: "12px", boxShadow: "0 4px 6px rgba(0,0,0,0.1)", width: "100%", maxWidth: "600px", textAlign: "center" }}>
        
        <h1 style={{ fontSize: "24px", fontWeight: "bold", color: "#111827", marginBottom: "8px" }}>
          Laboratoriya AI Tahlil Tizimi
        </h1>
        <p style={{ color: "#6b7280", fontSize: "14px", marginBottom: "30px" }}>
          Faylni yuklang, AI tahlil qo'shilgan tayyor hujjatni bemorga yuboring
        </p>

        <form onSubmit={handleSubmit}>
          <div style={{ border: "2px dashed #d1d5db", borderRadius: "8px", padding: "30px", marginBottom: "20px", backgroundColor: "#f9fafb" }}>
            <label style={{ display: "inline-block", padding: "10px 20px", backgroundColor: "#e5e7eb", color: "#374151", borderRadius: "6px", cursor: "pointer", fontWeight: "600", fontSize: "14px" }}>
              Faylni tanlash (.docx)
              <input type="file" accept=".docx" onChange={handleFileChange} style={{ display: "none" }} />
            </label>
            <p style={{ color: "#9ca3af", fontSize: "12px", marginTop: "10px" }}>
              yoki faylni shu yerga sudrab kelib tashlang
            </p>
            {file && (
              <p style={{ color: "#2563eb", fontSize: "14px", fontWeight: "600", marginTop: "10px" }}>
                Tanlandi: {file.name}
              </p>
            )}
          </div>

          <button
            type="submit"
            disabled={!file || loading}
            style={{
              width: "100%",
              padding: "12px",
              backgroundColor: !file || loading ? "#9ca3af" : "#2563eb",
              color: "white",
              border: "none",
              borderRadius: "6px",
              fontSize: "16px",
              fontWeight: "600",
              cursor: !file || loading ? "not-allowed" : "pointer",
            }}
          >
            {loading ? "Tahlil qilinmoqda va tayyorlanmoqda..." : "Tahlil qilish va Faylni Tayyorlash"}
          </button>
        </form>

        {downloadUrl && (
          <div style={{ marginTop: "30px", padding: "20px", backgroundColor: "#ecfdf5", border: "1px solid #a7f3d0", borderRadius: "8px" }}>
            <p style={{ color: "#065f46", fontSize: "14px", fontWeight: "600", marginBottom: "15px" }}>
              Fayl tayyorlandi!
            </p>
            <a
              href={downloadUrl}
              download={fileName}
              style={{
                display: "inline-block",
                padding: "10px 20px",
                backgroundColor: "#059669",
                color: "white",
                borderRadius: "6px",
                textDecoration: "none",
                fontWeight: "600",
                fontSize: "14px",
              }}
            >
              Tayyor! Faylni Yuklab Olish
            </a>
          </div>
        )}

      </div>
    </main>
  );
}
