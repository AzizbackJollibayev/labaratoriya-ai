"use client";

import React, { useCallback, useEffect, useRef, useState } from "react";

type Phase = "idle" | "dragging" | "processing" | "done" | "error";

export default function LaborantPage() {
  const [phase, setPhase] = useState<Phase>("idle");
  const [fileName, setFileName] = useState("");
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [downloadName, setDownloadName] = useState("");
  const [errorMsg, setErrorMsg] = useState("");
  const [today, setToday] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    setToday(
      new Date().toLocaleDateString("uz-UZ", { day: "2-digit", month: "2-digit", year: "numeric" })
    );
  }, []);

  const processFile = useCallback(async (file: File) => {
    if (!file.name.toLowerCase().endsWith(".docx")) {
      setPhase("error");
      setErrorMsg("Bu docx emas. Faqat .docx fayllar qabul qilinadi.");
      return;
    }

    setFileName(file.name);
    setPhase("processing");
    setErrorMsg("");

    const formData = new FormData();
    formData.append("file", file);

    try {
      const res = await fetch("/api/process-lab", { method: "POST", body: formData });

      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        throw new Error(data.error || "Faylni qayta ishlashda xatolik yuz berdi.");
      }

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      setDownloadUrl(url);
      setDownloadName(`TAHLIL_${file.name}`);
      setPhase("done");
    } catch (err: any) {
      setPhase("error");
      setErrorMsg(err?.message || "Kutilmagan xatolik yuz berdi.");
    }
  }, []);

  const onDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      const file = e.dataTransfer.files?.[0];
      if (file) processFile(file);
    },
    [processFile]
  );

  const onSelect = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (file) processFile(file);
    },
    [processFile]
  );

  const reset = () => {
    setPhase("idle");
    setFileName("");
    setDownloadUrl(null);
    setErrorMsg("");
    if (inputRef.current) inputRef.current.value = "";
  };

  const openPicker = () => {
    if (phase !== "processing") inputRef.current?.click();
  };

  return (
    <main className="page">
      <div className="tag-row mono">
        <span className="tag-eyebrow">Laboratoriya · AI izoh</span>
        <span className="tag-date">{today}</span>
      </div>

      <h1 className="headline">Tahlil faylini qo&apos;ying.</h1>
      <p className="sub">
        Natijalar normativ ko&apos;rsatkichlar bilan solishtiriladi, izoh biriktirilib qaytariladi.
      </p>

      <section
        className={`dropzone phase-${phase}`}
        onDragOver={(e) => {
          e.preventDefault();
          if (phase === "idle") setPhase("dragging");
        }}
        onDragLeave={() => {
          if (phase === "dragging") setPhase("idle");
        }}
        onDrop={onDrop}
        onClick={openPicker}
        role="button"
        tabIndex={0}
        aria-label="Tahlil faylini tashlang yoki tanlang"
        onKeyDown={(e) => {
          if (e.key === "Enter" || e.key === " ") {
            e.preventDefault();
            openPicker();
          }
        }}
      >
        <input
          ref={inputRef}
          type="file"
          accept=".docx"
          onChange={onSelect}
          style={{ display: "none" }}
        />

        <div className="prompt">
          {phase === "idle" && (
            <>
              <p className="prompt-main">Faylni shu yerga tashlang</p>
              <p className="prompt-sub mono">yoki bosib tanlang · .docx</p>
            </>
          )}
          {phase === "dragging" && <p className="prompt-main">Qo&apos;yib yuboring</p>}
          {phase === "processing" && (
            <>
              <p className="prompt-main">Solishtirilmoqda</p>
              <p className="prompt-sub mono">{fileName}</p>
            </>
          )}
          {phase === "done" && <p className="prompt-main">Tayyor</p>}
          {phase === "error" && <p className="prompt-main flag">{errorMsg}</p>}
        </div>

        <div className="track">
          <div className="zone" />
          <div className={`marker marker-${phase}`} />
        </div>
        <div className="track-labels mono">
          <span>past</span>
          <span>me&apos;yor oralig&apos;i</span>
          <span>yuqori</span>
        </div>
      </section>

      {phase === "done" && downloadUrl && (
        <div className="stub">
          <div className="stub-perf" aria-hidden="true" />
          <div className="stub-row">
            <div className="stub-info">
              <span className="stub-file mono">{downloadName}</span>
              <span className="stub-pill">tayyor</span>
            </div>
            <a className="stub-btn" href={downloadUrl} download={downloadName}>
              Faylni yuklab olish
            </a>
          </div>
          <button className="stub-reset" onClick={reset}>
            Yana bir fayl
          </button>
        </div>
      )}

      {phase === "error" && (
        <button className="retry" onClick={reset}>
          Qayta urinish
        </button>
      )}

      <style jsx>{`
        .page {
          max-width: 640px;
          margin: 0 auto;
          padding: 72px 24px 56px;
        }
        .tag-row {
          display: flex;
          justify-content: space-between;
          align-items: baseline;
          margin-bottom: 28px;
        }
        .tag-eyebrow {
          font-size: 12px;
          letter-spacing: 0.06em;
          text-transform: uppercase;
          color: var(--teal);
        }
        .tag-date {
          font-size: 12px;
          color: var(--ink-faint);
        }
        .headline {
          font-size: 32px;
          font-weight: 600;
          line-height: 1.2;
          color: var(--ink);
          margin-bottom: 10px;
        }
        .sub {
          font-size: 15px;
          color: var(--ink-muted);
          line-height: 1.6;
          margin: 0 0 36px;
          max-width: 46ch;
        }
        .dropzone {
          background: var(--card);
          border: 1px dashed var(--line);
          border-radius: 4px;
          padding: 36px 28px 28px;
          cursor: pointer;
          transition: border-color 0.15s ease, background 0.15s ease;
        }
        .dropzone.phase-dragging {
          border-color: var(--teal);
          border-style: solid;
          background: var(--teal-tint);
        }
        .dropzone.phase-processing {
          cursor: default;
          border-style: solid;
          border-color: var(--line);
        }
        .dropzone.phase-done {
          border-style: solid;
          border-color: var(--teal);
        }
        .dropzone.phase-error {
          border-style: solid;
          border-color: var(--flag);
        }
        .prompt {
          text-align: center;
          margin-bottom: 24px;
          min-height: 44px;
        }
        .prompt-main {
          font-family: "Space Grotesk", sans-serif;
          font-size: 17px;
          font-weight: 500;
          margin: 0 0 6px;
          color: var(--ink);
        }
        .prompt-main.flag {
          color: var(--flag);
          font-size: 14px;
          font-weight: 400;
          font-family: "Inter", sans-serif;
          max-width: 40ch;
          margin: 0 auto;
        }
        .prompt-sub {
          font-size: 12.5px;
          color: var(--ink-faint);
          margin: 0;
        }
        .track {
          position: relative;
          height: 34px;
          margin: 0 6px;
          background-image: repeating-linear-gradient(
            to right,
            var(--line) 0,
            var(--line) 1px,
            transparent 1px,
            transparent 8.33%
          );
          background-position: center;
          background-size: 100% 1px;
          background-repeat: no-repeat;
        }
        .zone {
          position: absolute;
          left: 33%;
          width: 34%;
          top: 0;
          bottom: 0;
          background: var(--teal-tint);
        }
        .marker {
          position: absolute;
          top: 50%;
          width: 12px;
          height: 12px;
          border-radius: 50%;
          background: var(--teal);
          transform: translate(-50%, -50%);
          opacity: 0;
          left: 6%;
        }
        .marker-dragging {
          opacity: 1;
          animation: pulse 1s ease-in-out infinite;
        }
        .marker-processing {
          opacity: 1;
          animation: sweep 1.3s cubic-bezier(0.65, 0, 0.35, 1) infinite;
        }
        .marker-done {
          opacity: 1;
          left: 78%;
          transition: left 0.6s ease;
        }
        .marker-error {
          opacity: 1;
          left: 6%;
          background: var(--flag);
        }
        @keyframes pulse {
          0%,
          100% {
            transform: translate(-50%, -50%) scale(1);
          }
          50% {
            transform: translate(-50%, -50%) scale(1.4);
          }
        }
        @keyframes sweep {
          0% {
            left: 6%;
          }
          50% {
            left: 82%;
          }
          100% {
            left: 6%;
          }
        }
        .track-labels {
          display: flex;
          justify-content: space-between;
          margin: 8px 6px 0;
          font-size: 11px;
          color: var(--ink-faint);
        }
        .stub {
          margin-top: 20px;
          background: var(--card);
          border: 1px solid var(--line);
          border-radius: 4px;
          padding: 4px 20px 18px;
        }
        .stub-perf {
          height: 1px;
          margin: 0 -20px 16px;
          background-image: repeating-linear-gradient(
            to right,
            var(--line) 0,
            var(--line) 4px,
            transparent 4px,
            transparent 9px
          );
        }
        .stub-row {
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 16px;
          flex-wrap: wrap;
        }
        .stub-info {
          display: flex;
          align-items: center;
          gap: 10px;
          min-width: 0;
        }
        .stub-file {
          font-size: 13px;
          color: var(--ink);
          overflow: hidden;
          text-overflow: ellipsis;
          white-space: nowrap;
        }
        .stub-pill {
          flex-shrink: 0;
          font-size: 11px;
          color: var(--teal-deep);
          background: var(--teal-tint);
          padding: 3px 9px;
          border-radius: 3px;
        }
        .stub-btn {
          flex-shrink: 0;
          background: var(--ink);
          color: var(--paper);
          text-decoration: none;
          font-size: 13.5px;
          font-weight: 500;
          padding: 10px 18px;
          border-radius: 4px;
        }
        .stub-reset {
          display: block;
          margin: 14px 0 0;
          background: none;
          border: none;
          color: var(--ink-faint);
          font-size: 12.5px;
          text-decoration: underline;
          cursor: pointer;
          padding: 0;
        }
        .retry {
          margin-top: 16px;
          background: none;
          border: 1px solid var(--flag);
          color: var(--flag);
          font-size: 13px;
          padding: 8px 16px;
          border-radius: 4px;
          cursor: pointer;
        }
      `}</style>
    </main>
  );
}
