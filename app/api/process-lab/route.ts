import { NextRequest, NextResponse } from "next/server";
import { GoogleGenerativeAI } from "@google/generative-ai";
import JSZip from "jszip";

// Matnni xavfsiz XML formatiga o'tkazish
function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// Bitta <w:r> (run) yaratadi
function makeRun(text: string, bold: boolean): string {
  return `
    <w:r>
      <w:rPr>
        ${bold ? "<w:b/>" : ""}
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:sz w:val="28"/>
        <w:szCs w:val="28"/>
        <w:lang w:val="en-US"/>
      </w:rPr>
      <w:t xml:space="preserve">${escapeXml(text)}</w:t>
    </w:r>`;
}

// AI qaytargan matnni qatorlarga bo'lib, HAR BIR qatorni alohida Word paragraf qiladi.
// Har qatorda ":" bo'lsa — ":" gacha bo'lgan qism (ko'rsatkich nomi) QALIN, qolgani oddiy yoziladi.
// "- ", "* ", "1. " kabi belgi bilan boshlangan qatorlar bullet (•) va otstup bilan ajratiladi.
function textToWordParagraphs(text: string): string {
  const lines = text
    .split("\n")
    .map((l) => l.trim())
    .filter((l) => l.length > 0);

  return lines
    .map((line) => {
      const isBullet = /^([-*•]|\d+[.)])\s+/.test(line);
      let cleanLine = line.replace(/^([-*•]|\d+[.)])\s+/, "").replace(/\*\*/g, "");

      const indent = isBullet ? `<w:ind w:left="360"/>` : "";
      const bulletPrefix = isBullet ? "•  " : "";

      // ":" bo'yicha bold label / oddiy matn qismlarga ajratish
      const colonIdx = cleanLine.indexOf(":");
      let runs: string;
      if (colonIdx > -1 && colonIdx < 60) {
        const label = cleanLine.slice(0, colonIdx + 1); // ":" bilan birga
        const rest = cleanLine.slice(colonIdx + 1);
        runs = makeRun(bulletPrefix + label, true) + makeRun(rest, false);
      } else {
        runs = makeRun(bulletPrefix + cleanLine, false);
      }

      return `
      <w:p>
        <w:pPr>
          <w:jc w:val="left"/>
          ${indent}
        </w:pPr>
        ${runs}
      </w:p>`;
    })
    .join("");
}

// ---------- Asosiy model band bo'lsa, avtomatik zaxira modelga o'tish ----------

const PRIMARY_MODEL = "gemini-3.6-flash";
const FALLBACK_MODEL = "gemini-2.5-flash-lite";

function isOverloadError(err: any): boolean {
  const status = err?.status ?? err?.response?.status;
  const message = String(err?.message || "").toLowerCase();
  return (
    status === 429 ||
    status === 503 ||
    status === 404 ||
    message.includes("overloaded") ||
    message.includes("resource_exhausted") ||
    message.includes("unavailable") ||
    message.includes("service is currently unavailable") ||
    message.includes("too many requests") ||
    message.includes("quota") ||
    message.includes("not found") ||
    message.includes("not_found") ||
    message.includes("429") ||
    message.includes("503") ||
    message.includes("404")
  );
}

async function generateWithFallback(
  genAI: GoogleGenerativeAI,
  systemInstruction: string,
  prompt: string
): Promise<{ text: string; modelUsed: string }> {
  try {
    const primary = genAI.getGenerativeModel({
      model: PRIMARY_MODEL,
      systemInstruction,
    });
    const result = await primary.generateContent(prompt);
    return { text: result.response.text().trim(), modelUsed: PRIMARY_MODEL };
  } catch (err: any) {
    if (!isOverloadError(err)) {
      throw err;
    }

    console.warn(
      `[AI] ${PRIMARY_MODEL} band/ishlamayapti (${err?.message}), ${FALLBACK_MODEL} ga o'tildi.`
    );

    const fallback = genAI.getGenerativeModel({
      model: FALLBACK_MODEL,
      systemInstruction,
    });
    const result = await fallback.generateContent(prompt);
    return { text: result.response.text().trim(), modelUsed: FALLBACK_MODEL };
  }
}

// TUZATILDI: avval barcha <w:t> matnlar bitta qatorga bo'shliq bilan qo'shib
// yuborilardi — jadval qator/ustun chegaralari yo'qolib, AI natija va norma
// qiymatlarini noto'g'ri bog'lashi yoki ba'zi ko'rsatkichlarni butunlay
// o'tkazib yuborishi mumkin edi. Endi har bir jadval qatori (<w:tr>) alohida
// qatorga chiqariladi va katakchalar " | " bilan ajratiladi — shunda
// "Ko'rsatkich | Natija | Norma | birlik" tuzilishi saqlanadi.
function extractTextFromDocumentXml(docXml: string): string {
  const withoutDeletedRuns = docXml.replace(/<w:del\b[^>]*>[\s\S]*?<\/w:del>/g, "");

  const rows = withoutDeletedRuns.match(/<w:tr\b[\s\S]*?<\/w:tr>/g);
  if (rows && rows.length > 0) {
    const lines = rows.map((row) => {
      const cells = row.match(/<w:tc\b[\s\S]*?<\/w:tc>/g) || [];
      const cellTexts = cells.map((cell) => {
        const tMatches = cell.match(/<w:t[^>]*>(.*?)<\/w:t>/g) || [];
        return tMatches
          .map((t) => t.replace(/<[^>]+>/g, ""))
          .join("")
          .trim();
      });
      return cellTexts.filter((c) => c.length > 0).join(" | ");
    });
    const tableText = lines.filter((l) => l.length > 0).join("\n");

    // Jadvaldan tashqaridagi matn (sarlavha, F.I.O, sana kabi) ham kerak bo'lishi mumkin
    const withoutTables = withoutDeletedRuns.replace(/<w:tbl\b[\s\S]*?<\/w:tbl>/g, "\n");
    const outsideMatches = withoutTables.match(/<w:t[^>]*>(.*?)<\/w:t>/g) || [];
    const outsideText = outsideMatches.map((t) => t.replace(/<[^>]+>/g, "")).join(" ").trim();

    return `${outsideText}\n\n${tableText}`.trim();
  }

  const textMatches = withoutDeletedRuns.match(/<w:t[^>]*>(.*?)<\/w:t>/g) || [];
  return textMatches.map((t) => t.replace(/<[^>]+>/g, "")).join(" ");
}

const MAX_FILE_SIZE_MB = 15;

export async function POST(req: NextRequest) {
  try {
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      return NextResponse.json(
        { error: ".env.local faylida GEMINI_API_KEY topilmadi." },
        { status: 500 }
      );
    }

    const genAI = new GoogleGenerativeAI(apiKey);

    const formData = await req.formData();
    const file = formData.get("file");

    if (!file || !(file instanceof File)) {
      return NextResponse.json({ error: "Fayl tanlanmadi" }, { status: 400 });
    }

    if (!file.name.toLowerCase().endsWith(".docx")) {
      return NextResponse.json(
        { error: "Faqat .docx formatidagi Word fayllarini yuklashingiz mumkin" },
        { status: 400 }
      );
    }

    if (file.size > MAX_FILE_SIZE_MB * 1024 * 1024) {
      return NextResponse.json(
        { error: `Fayl hajmi ${MAX_FILE_SIZE_MB} MB dan oshmasligi kerak` },
        { status: 400 }
      );
    }

    const arrayBuffer = await file.arrayBuffer();

    let zip: JSZip;
    let docXml: string | undefined;
    try {
      zip = await JSZip.loadAsync(arrayBuffer);
      docXml = await zip.file("word/document.xml")?.async("string");
    } catch {
      return NextResponse.json(
        { error: "Fayl buzilgan yoki noto'g'ri .docx formatda" },
        { status: 400 }
      );
    }

    if (!docXml) {
      return NextResponse.json({ error: "DOCX faylni o'qishda xatolik yuz berdi" }, { status: 400 });
    }

    const extractedText = extractTextFromDocumentXml(docXml);

    if (!extractedText.trim()) {
      return NextResponse.json(
        { error: "Fayl ichida tahlil qilinadigan matn topilmadi" },
        { status: 400 }
      );
    }

    // TUZATILDI: endi "faqat me'yordan chetlashganlarni yoz" emas, balki
    // "natijasi mavjud BARCHA ko'rsatkichni" tahlil qiladi — bemor topshirgan
    // har bir tahlil to'liq ko'rib chiqiladi. Shu bilan birga har bir qator
    // qat'iy qisqa (bir qatorlik) tutiladi — to'liq, lekin cho'zilmagan.
    const systemInstruction =
      "Siz tajribali laboratoriya shifokorisiz. Sizga jadval ko'rinishidagi laboratoriya tahlili " +
      "berilgan; har bir qator 'katakcha1 | katakcha2 | ... ' formatida, odatda " +
      "'Ko'rsatkich nomi | Natija | Norma | O'lchov birligi' tartibida keladi.\n\n" +
      "QOIDALAR:\n" +
      "1. Faqat NATIJA QIYMATI ko'rsatilgan (bo'sh bo'lmagan) ko'rsatkichlarni tahlil qiling. " +
      "Natija katakchasi bo'sh bo'lsa (o'lchov o'tkazilmagan), o'sha ko'rsatkichni butunlay o'tkazib yuboring.\n" +
      "2. Bemor topshirgan HAR BIR natijasi mavjud ko'rsatkich uchun ALOHIDA qator yozing — " +
      "birortasini ham tashlab ketmang, birortasini ham birlashtirmang. Maqsad — bemorga barcha " +
      "topshirilgan tahlillarni TO'LIQ tushuntirish, faqat chetlashganlarini emas.\n" +
      "3. Har bir qatorda: ko'rsatkich nomi, natija qiymati (birligi bilan), me'yor bilan solishtirilgan " +
      "holati (me'yorda / me'yordan yuqori / me'yordan past). Me'yorda bo'lsa — shunchaki 'me'yorda' deb " +
      "yozing, qo'shimcha izohsiz. Chetlashgan bo'lsa — 5-10 so'zli juda qisqa tibbiy izoh qo'shing " +
      "(nima anglatishi mumkinligi haqida, tashxis qo'ymasdan).\n" +
      "4. Tilni sodda va tushunarli tuting — bemor ham o'qib tushunadigan darajada, tibbiy jargonni " +
      "kamaytiring, gaplarni cho'zmang. Har bir qator BIR QATORDAN OSHMASIN.\n" +
      "5. Faqat oddiy matn (markdown, **, # belgilarisiz) bilan javob bering.\n\n" +
      "JAVOB FORMATI (qat'iy shu tartibda):\n\n" +
      "1-qator: 'Qon: ' bilan boshlanadigan bitta gaplik juda qisqa umumiy xulosa " +
      "(masalan: 'Qon: N ta ko'rsatkich tekshirildi, ulardan M tasi me'yorda, K tasida chetlashish bor.').\n\n" +
      "Keyingi qatorlar: natijasi mavjud HAR BIR ko'rsatkich uchun bittadan qator, " +
      "'- Ko'rsatkich nomi: natija (birlik) — holati, [chetlashsa qisqa izoh]' formatida.\n\n" +
      "Oxirgi qator: 'Tavsiya: ' bilan boshlanadigan bir gaplik amaliy tavsiya " +
      "(qaysi shifokorga murojaat qilish kerakligi, chetlashishlar asosida; hech narsa chetlashmagan " +
      "bo'lsa — profilaktik ko'rikni davom ettirish tavsiyasi).\n\n" +
      "Ko'rsatkich nomlarini aniq va tibbiy jihatdan to'g'ri yozing, taxmin qilmang — faqat berilgan " +
      "ma'lumotlarga tayaning.";

    const prompt = `Quyidagi laboratoriya tahlil jadvalini o'rganib chiqib, yuqoridagi qoida va formatga qat'iy rioya qilgan holda, bemor topshirgan HAR BIR natijasi mavjud ko'rsatkichni to'liq, ammo qisqa tushuntiring:\n\n${extractedText}`;

    let analysisResult: string;
    let modelUsed: string;
    try {
      const result = await generateWithFallback(genAI, systemInstruction, prompt);
      analysisResult = result.text;
      modelUsed = result.modelUsed;
    } catch (aiErr: any) {
      console.error("[AI] Ikkala model ham ishlamadi:", aiErr?.message);
      return NextResponse.json(
        {
          error:
            "AI xizmati hozircha band yoki kunlik limit tugagan. Bir necha daqiqadan so'ng qayta urinib ko'ring.",
        },
        { status: 503 }
      );
    }

    const bodyParagraphs = textToWordParagraphs(analysisResult);

    // Shablon bilan bir xil formatda: markazlashtirilgan qalin sarlavha +
    // sz=28 li paragraflar, DIQQAT bloki matn oxirida (avvalgidek, footersiz)
    const xmlToInsert = `
      <w:p/>
      <w:p/>
      <w:p>
        <w:pPr>
          <w:jc w:val="center"/>
        </w:pPr>
        <w:r>
          <w:rPr>
            <w:b/>
            <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
            <w:sz w:val="28"/>
            <w:szCs w:val="28"/>
            <w:lang w:val="en-US"/>
          </w:rPr>
          <w:t>Analiz xulosasi</w:t>
        </w:r>
      </w:p>
      <w:p/>
      ${bodyParagraphs}
      <w:p/>
      <w:p/>
      <w:p>
        <w:pPr>
          <w:jc w:val="center"/>
        </w:pPr>
        <w:r>
          <w:rPr>
            <w:b/>
            <w:color w:val="FF5555"/>
            <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
            <w:sz w:val="22"/>
          </w:rPr>
          <w:t>DIQQAT / OGOHLANTIRISH:</w:t>
        </w:r>
      </w:p>
      <w:p>
        <w:pPr>
          <w:jc w:val="center"/>
        </w:pPr>
        <w:r>
          <w:rPr>
            <w:color w:val="FF5555"/>
            <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
            <w:sz w:val="20"/>
          </w:rPr>
          <w:t>Ushbu tahlil xulosasi Sun'iy Intellekt (AI) tizimi tomonidan avtomatik ravishda tayyorlangan. Davolanish va aniq tashxis uchun mutaxassis shifokorga murojaat qiling.</w:t>
        </w:r>
      </w:p>
    `;

    // MUHIM: <w:sectPr> w:body ichidagi ENG OXIRGI element bo'lishi SHART
    // (OOXML standarti). Shuning uchun kontentni sectPr'dan OLDIN joylashtiramiz.
    const sectPrRegex = /(<w:sectPr\b[^>]*\/>|<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>)(\s*<\/w:body>)/;
    let updatedXml: string;
    if (sectPrRegex.test(docXml)) {
      updatedXml = docXml.replace(sectPrRegex, `${xmlToInsert}$1$2`);
    } else {
      updatedXml = docXml.replace("</w:body>", `${xmlToInsert}</w:body>`);
    }

    zip.file("word/document.xml", updatedXml);
    const modifiedBuffer = await zip.generateAsync({ type: "uint8array" });

    const encodedFilename = encodeURIComponent(`TAHLIL_${file.name}`);

    return new NextResponse(modifiedBuffer as unknown as BodyInit, {
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": `attachment; filename="${encodedFilename}"; filename*=UTF-8''${encodedFilename}`,
        "X-AI-Model-Used": modelUsed,
      },
    });
  } catch (err: any) {
    console.error("[process-lab] Kutilmagan xato:", err);
    return NextResponse.json(
      { error: "Faylni qayta ishlashda kutilmagan xatolik yuz berdi" },
      { status: 500 }
    );
  }
}
