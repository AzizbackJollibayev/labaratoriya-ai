import JSZip from "jszip";

// ---------- Yordamchi: raqamni ajratib olish ----------
function toFloat(raw: string): number | null {
  const s = raw.trim().replace(",", ".").replace(/\s/g, "");
  const v = parseFloat(s);
  return isNaN(v) ? null : v;
}

// ---------- Norma matnini tahlil qilish ----------
export type Norm =
  | { type: "range"; lo: number | null; hi: number | null }
  | { type: "complex"; raw: string };

export function parseNorm(rawText: string): Norm | null {
  const text = rawText.trim();
  if (!text) return null;

  // Jins (Э/А) yoki yoshga bog'liq murakkab norma — hozircha avtomatik hisoblanmaydi
  if (text.includes("Э:") || text.includes("А:") || text.includes("ёш")) {
    return { type: "complex", raw: text };
  }

  // Oddiy diapazon: "3,2-6,1"
  const rangeMatch = text.match(/(-?\d+[.,]?\d*)\s*[-–]\s*(-?\d+[.,]?\d*)/);
  if (rangeMatch) {
    const lo = toFloat(rangeMatch[1]);
    const hi = toFloat(rangeMatch[2]);
    if (lo !== null && hi !== null) return { type: "range", lo, hi };
  }

  // Faqat yuqori chegara: "5,1 гача" yoki "<40"
  const upperMatch = text.match(/(?:<|гача)\s*(-?\d+[.,]?\d*)|(-?\d+[.,]?\d*)\s*гача/);
  if (upperMatch) {
    const v = toFloat(upperMatch[1] || upperMatch[2]);
    if (v !== null) return { type: "range", lo: null, hi: v };
  }

  return { type: "complex", raw: text };
}

// ---------- document.xml dagi jadval katakchalarini o'qish ----------

function stripTagsGetText(cellXml: string): string {
  // Har bir <w:p>...</w:p> — alohida qator (\n bilan ajratiladi, python-docx kabi)
  const paragraphs = cellXml.split(/<\/w:p>/);
  const lines = paragraphs.map((p) => {
    const matches = [...p.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)];
    return matches
      .map((m) =>
        m[1]
          .replace(/&amp;/g, "&")
          .replace(/&lt;/g, "<")
          .replace(/&gt;/g, ">")
          .replace(/&quot;/g, '"')
          .replace(/&apos;/g, "'")
      )
      .join("");
  });
  return lines.filter((l) => l.length > 0).join("\n");
}

function splitTopLevel(xml: string, tag: string): string[] {
  // <w:tag ...>...</w:tag> bloklarini (ichma-ich bo'lmagan) ajratib oladi
  const results: string[] = [];
  const openRe = new RegExp(`<w:${tag}(?:\\s[^>]*)?>`, "g");
  let m: RegExpExecArray | null;
  while ((m = openRe.exec(xml)) !== null) {
    const start = m.index;
    // shu tagning yopilishini topamiz (nested bir xil tag hisobga olinmaydi — bu hujjatda kerak emas)
    const closeTag = `</w:${tag}>`;
    const end = xml.indexOf(closeTag, start);
    if (end === -1) break;
    results.push(xml.slice(start, end + closeTag.length));
    openRe.lastIndex = end + closeTag.length;
  }
  return results;
}

export interface Finding {
  test: string;
  resultRaw: string;
  resultVal: number;
  normRaw: string;
  norm: Norm | null;
}

export function extractFilledResults(documentXml: string): {
  findings: Finding[];
  patientLine: string;
} {
  const findings: Finding[] = [];
  let patientLine = "";

  const tables = splitTopLevel(documentXml, "tbl");
  for (const tableXml of tables) {
    const rows = splitTopLevel(tableXml, "tr");
    if (rows.length === 0) continue;

    const rowCells: string[][] = rows.map((rowXml) =>
      splitTopLevel(rowXml, "tc").map(stripTagsGetText)
    );

    // Bemor F.I.O ni qidiramiz (har qanday jadvalda bo'lishi mumkin)
    for (const cells of rowCells) {
      for (let i = 0; i < cells.length; i++) {
        if (cells[i].includes("Ф.И.О") && cells[i + 1]) {
          patientLine = `Bemor: ${cells[i + 1]}`;
        }
      }
    }

    // Sarlavha qatorini topamiz: "Натижа" va "Норма" so'zlari bor ustunlar
    let headerIdx = -1;
    let natijaCol = -1;
    let normaCol = -1;
    for (let r = 0; r < rowCells.length; r++) {
      for (let c = 0; c < rowCells[r].length; c++) {
        if (rowCells[r][c].includes("Натижа") && natijaCol === -1) natijaCol = c;
        if (rowCells[r][c].includes("Норма") && normaCol === -1) normaCol = c;
      }
      if (natijaCol !== -1 && normaCol !== -1) {
        headerIdx = r;
        break;
      }
    }
    if (headerIdx === -1) continue; // bu jadval mos emas

    for (let r = headerIdx + 1; r < rowCells.length; r++) {
      const cells = rowCells[r];
      if (natijaCol >= cells.length || normaCol >= cells.length) continue;
      const testName = (cells[0] || "").replace(/\n/g, " ").trim();
      const resultRaw = (cells[natijaCol] || "").trim();
      const normRaw = (cells[normaCol] || "").trim();
      if (!testName || !resultRaw) continue;
      const resultVal = toFloat(resultRaw);
      if (resultVal === null) continue;
      findings.push({
        test: testName,
        resultRaw,
        resultVal,
        normRaw,
        norm: parseNorm(normRaw),
      });
    }
  }

  return { findings, patientLine };
}

// ---------- Statusni aniqlash ----------
export type Status = "normal" | "high" | "low" | "unknown";

export function classify(finding: Finding): Status {
  if (!finding.norm || finding.norm.type !== "range") return "unknown";
  const { lo, hi } = finding.norm;
  if (lo !== null && finding.resultVal < lo) return "low";
  if (hi !== null && finding.resultVal > hi) return "high";
  return "normal";
}

// ---------- Izoh yozish (shablon yoki AI) ----------
function templateComment(f: Finding, status: Status): string {
  const dir = status === "high" ? "yuqori" : "past";
  return `${f.test} darajasi belgilangan normadan ${dir} chiqdi (${f.resultRaw}, norma: ${f.normRaw}). Bu ko'rsatkich turli sabablarga bog'liq bo'lishi mumkin — aniq sababni faqat shifokor aniqlay oladi.`;
}

export async function generateComment(f: Finding, status: Status): Promise<string> {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) return templateComment(f, status);

  try {
    const Anthropic = (await import("@anthropic-ai/sdk")).default;
    const client = new Anthropic({ apiKey });
    const dir = status === "high" ? "yuqori" : "past";
    const resp = await client.messages.create({
      model: "claude-sonnet-4-6",
      max_tokens: 200,
      messages: [
        {
          role: "user",
          content: `Bemor tahlil natijasi: ${f.test} = ${f.resultRaw}, norma: ${f.normRaw}. Bu ko'rsatkich normadan ${dir}. O'zbek tilida, 2-3 gapli, tushunarli va tibbiy jihatdan ehtiyotkor izoh yoz. Tashxis qo'yma, faqat mumkin bo'lgan umumiy izoh ber va shifokorga murojaat qilishni tavsiya qil.`,
        },
      ],
    });
    const block = resp.content[0];
    if (block.type === "text") return block.text.trim();
    return templateComment(f, status);
  } catch {
    return templateComment(f, status);
  }
}

// ---------- Word bo'limini quruvchi (XML) ----------
function esc(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function p(
  text: string,
  opts: { bold?: boolean; italic?: boolean; size?: number; color?: string } = {}
): string {
  const { bold, italic, size = 22, color } = opts;
  let rpr = "";
  if (bold) rpr += "<w:b/>";
  if (italic) rpr += "<w:i/>";
  rpr += `<w:sz w:val="${size}"/><w:szCs w:val="${size}"/>`;
  if (color) rpr += `<w:color w:val="${color}"/>`;
  return `<w:p><w:pPr><w:rPr>${rpr}</w:rPr></w:pPr><w:r><w:rPr>${rpr}</w:rPr><w:t xml:space="preserve">${esc(
    text
  )}</w:t></w:r></w:p>`;
}

export function buildSectionXml(
  patientLine: string,
  results: { finding: Finding; status: Status; comment: string }[]
): string {
  let xml = "";
  xml += p("", { size: 20 });
  xml += p("———————————————————————————", { size: 22 });
  xml += p("TAHLIL NATIJALARI BO\u02BBYICHA AI IZOHI", { bold: true, size: 26 });
  xml += p("", { size: 16 });
  if (patientLine) {
    xml += p(patientLine, { bold: true, size: 22 });
    xml += p("", { size: 16 });
  }

  const normalCount = results.filter((r) => r.status === "normal").length;
  const flagged = results.filter((r) => r.status === "high" || r.status === "low");
  const unknown = results.filter((r) => r.status === "unknown");

  xml += p(
    `Tekshirilgan ko'rsatkichlar: ${results.length} ta (normal: ${normalCount}, e'tibor talab qiladi: ${flagged.length})`,
    { size: 22 }
  );
  xml += p("", { size: 16 });

  if (flagged.length > 0) {
    xml += p("E'tibor talab qiladigan ko'rsatkichlar:", { bold: true, size: 22 });
    for (const { finding: f, status, comment } of flagged) {
      const label = status === "high" ? "YUQORI" : "PAST";
      xml += p(`\u2022 ${f.test}: ${f.resultRaw}  (norma: ${f.normRaw}) \u2014 ${label}`, {
        color: "C00000",
        size: 22,
      });
      xml += p(`  ${comment}`, { size: 20, italic: true });
    }
    xml += p("", { size: 16 });
  }

  if (unknown.length > 0) {
    xml += p(
      "Avtomatik solishtirib bo'lmagan (jins/yoshga bog'liq murakkab norma — laborant tomonidan qo'lda tekshirilishi kerak):",
      { bold: true, size: 20 }
    );
    for (const { finding: f } of unknown) {
      xml += p(`\u2022 ${f.test}: ${f.resultRaw}  (norma: ${f.normRaw})`, { size: 20 });
    }
    xml += p("", { size: 16 });
  }

  xml += p(
    "\u26A0 MUHIM ESLATMA: Ushbu izoh sun'iy intellekt (AI) tomonidan avtomatik tuzilgan. Bu tibbiy tashxis emas. Yakuniy tashxis va davolash tavsiyasi uchun albatta shifokoringiz bilan maslahatlashing.",
    { italic: true, size: 20, color: "595959" }
  );
  xml += p("", { size: 16 });
  return xml;
}

// ---------- Asosiy oqim: .docx buferini qayta ishlash ----------
export async function processLabDocx(inputBuffer: Buffer): Promise<Buffer> {
  const zip = await JSZip.loadAsync(inputBuffer);
  const docXmlFile = zip.file("word/document.xml");
  if (!docXmlFile) throw new Error("Noto'g'ri .docx fayl: word/document.xml topilmadi");

  const documentXml = await docXmlFile.async("string");
  const { findings, patientLine } = extractFilledResults(documentXml);

  if (findings.length === 0) {
    // Hech narsa topilmasa, asl faylni qaytaramiz
    return inputBuffer;
  }

  const results = await Promise.all(
    findings.map(async (f) => {
      const status = classify(f);
      const comment = status === "high" || status === "low" ? await generateComment(f, status) : "";
      return { finding: f, status, comment };
    })
  );

  const sectionXml = buildSectionXml(patientLine, results);
  const sectPrIdx = documentXml.indexOf("<w:sectPr");
  const newXml =
    sectPrIdx === -1
      ? documentXml + sectionXml
      : documentXml.slice(0, sectPrIdx) + sectionXml + documentXml.slice(sectPrIdx);

  zip.file("word/document.xml", newXml);
  const outBuffer = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  return outBuffer;
}
