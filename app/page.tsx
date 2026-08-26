"use client";

import React, { useCallback, useRef, useState } from "react";

type Phase = "idle" | "dragging" | "processing" | "done" | "error";

export default function LaborantPage() {
  const [phase, setPhase] = useState<Phase>("idle");
  const [fileName, setFileName] = useState("");
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [downloadName, setDownloadName] = useState("");
  const [errorMsg, setErrorMsg] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);

  const processFile = useCallback(async (file: File) => {
    if (!file.name.toLowerCase().endsWith(".docx")) {
      setPhase("error");
      setErrorMsg("faqat .docx");
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
        throw new Error(data.error || "xatolik");
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      setDownloadUrl(url);
      setDownloadName(`TAHLIL_${file.name}`);
      setPhase("done");
    } catch (err: any) {
      setPhase("error");
      setErrorMsg(err?.message || "xatolik");
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

  const stateWord =
    phase === "idle"
      ? "tashlang"
      : phase === "dragging"
      ? "qo'ying"
      : phase === "processing"
      ? "skanerlanmoqda"
      : phase === "done"
      ? "tayyor"
      : "xato";

  return (
    <main className="page">
      <span className="tag">lab · ai</span>

      <section
        className={`frame phase-${phase}`}
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

        <span className="corner tl" />
        <span className="corner tr" />
        <span className="corner bl" />
        <span className="corner br" />

        {phase === "processing" && <span className="beam" />}

        <div className="center">
          <span className="word">{stateWord}</span>
          {phase === "idle" && <span className="hint">.docx</span>}
          {phase === "processing" && <span className="hint">{fileName}</span>}
          {phase === "error" && <span className="hint flag">{errorMsg}</span>}
        </div>
      </section>

      {phase === "done" && downloadUrl && (
        <div className="stub">
          <span className="stub-name">{downloadName}</span>
          <div className="stub-actions">
            <a className="stub-dl" href={downloadUrl} download={downloadName}>
              yuklab olish
            </a>
            <button className="stub-reset" onClick={reset} aria-label="Yana bir fayl">
              ↺
            </button>
          </div>
        </div>
      )}

      {phase === "error" && (
        <button className="retry" onClick={reset}>
          qayta
        </button>
      )}

      <style jsx>{`
        .page {
          max-width: 480px;
          margin: 0 auto;
          padding: 88px 24px 40px;
          display: flex;
          flex-direction: column;
          align-items: center;
        }
        .tag {
          align-self: flex-start;
          font-size: 11px;
          letter-spacing: 0.08em;
          color: var(--text-muted);
          margin-bottom: 20px;
        }
        .frame {
          position: relative;
          width: 100%;
          aspect-ratio: 1 / 0.82;
          background: var(--panel);
          border: 1px solid var(--line);
          border-radius: 2px;
          cursor: pointer;
          overflow: hidden;
          transition: border-color 0.15s ease, background 0.15s ease;
        }
        .frame.phase-dragging {
          background: var(--blue-soft);
        }
        .frame.phase-processing {
          cursor: default;
        }
        .frame.phase-error {
          border-color: var(--flag);
        }
        .corner {
          position: absolute;
          width: 22px;
          height: 22px;
          border-color: var(--blue);
          transition: border-color 0.15s ease;
        }
        .phase-error .corner {
          border-color: var(--flag);
        }
        .phase-dragging .corner,
        .phase-processing .corner {
          animation: glow 1.4s ease-in-out infinite;
        }
        .tl {
          top: 14px;
          left: 14px;
          border-top: 2px solid;
          border-left: 2px solid;
        }
        .tr {
          top: 14px;
          right: 14px;
          border-top: 2px solid;
          border-right: 2px solid;
        }
        .bl {
          bottom: 14px;
          left: 14px;
          border-bottom: 2px solid;
          border-left: 2px solid;
        }
        .br {
          bottom: 14px;
          right: 14px;
          border-bottom: 2px solid;
          border-right: 2px solid;
        }
        @keyframes glow {
          0%,
          100% {
            opacity: 0.55;
          }
          50% {
            opacity: 1;
          }
        }
        .beam {
          position: absolute;
          left: 14px;
          right: 14px;
          height: 2px;
          background: var(--blue);
          box-shadow: 0 0 12px 2px var(--blue);
          animation: sweep 1.6s ease-in-out infinite;
        }
        @keyframes sweep {
          0% {
            top: 14px;
          }
          50% {
            top: calc(100% - 16px);
          }
          100% {
            top: 14px;
          }
        }
        .center {
          position: absolute;
          inset: 0;
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          gap: 8px;
        }
        .word {
          font-size: 19px;
          font-weight: 600;
          color: var(--text);
        }
        .hint {
          font-size: 12px;
          color: var(--text-muted);
        }
        .hint.flag {
          color: var(--flag);
        }
        .stub {
          width: 100%;
          margin-top: 16px;
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 12px;
          padding: 12px 4px;
          border-top: 1px solid var(--line);
        }
        .stub-name {
          font-size: 12px;
          color: var(--text-muted);
          overflow: hidden;
          text-overflow: ellipsis;
          white-space: nowrap;
        }
        .stub-actions {
          display: flex;
          align-items: center;
          gap: 14px;
          flex-shrink: 0;
        }
        .stub-dl {
          color: var(--blue);
          text-decoration: none;
          font-size: 13px;
          font-weight: 600;
        }
        .stub-reset {
          background: none;
          border: none;
          color: var(--text-muted);
          font-size: 15px;
          cursor: pointer;
          line-height: 1;
          padding: 2px;
        }
        .retry {
          margin-top: 14px;
          background: none;
          border: 1px solid var(--flag);
          color: var(--flag);
          font-size: 12px;
          padding: 7px 14px;
          border-radius: 2px;
          cursor: pointer;
        }
      `}</style>
    </main>
  );
}
