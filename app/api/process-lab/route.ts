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

// Bitta <w:r> (run) yaratadi. `color` berilsa (masalan "C00000"), matn shu rangda chiqadi —
// bu AI natijasi original hujjat bilan mos kelmagan qatorlarni belgilash uchun ishlatiladi.
function makeRun(text: string, bold: boolean, color?: string): string {
  return `
    <w:r>
      <w:rPr>
        ${bold ? "<w:b/>" : ""}
        ${color ? `<w:color w:val="${color}"/>` : ""}
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:sz w:val="28"/>
        <w:szCs w:val="28"/>
        <w:lang w:val="en-US"/>
      </w:rPr>
      <w:t xml:space="preserve">${escapeXml(text)}</w:t>
    </w:r>`;
}

// ---------- Raqamlarni solishtirish qatlami ----------
// Original jadval matnidan barcha raqamlarni (natija va norma qiymatlari) yig'ib olamiz.
// AI qaytargan har bir qatordagi "natija" raqami shu to'plamda bor-yo'qligini tekshiramiz.
// Eslatma: bu "oddiy" tekshiruv — norma oralig'idagi raqam ham to'plamda bo'lgani uchun,
// agar AI natija o'rniga normani yozib qo'ysa, bu holat har doim ham ushlanmasligi mumkin.
// Lekin AI butunlay o'zidan raqam o'ylab topsa (hallyutsinatsiya), bu albatta ushlanadi.

function extractNumbersFromText(text: string): Set<number> {
  const matches = text.match(/-?\d+(?:[.,]\d+)?/g) || [];
  const nums = new Set<number>();
  for (const m of matches) {
    const val = parseFloat(m.replace(",", "."));
    if (!isNaN(val)) {
      nums.add(Math.round(val * 1000) / 1000);
    }
  }
  return nums;
}

function numberExistsIn(value: number, set: Set<number>, epsilon = 0.01): boolean {
  const rounded = Math.round(value * 1000) / 1000;
  if (set.has(rounded)) return true;
  for (const n of set) {
    if (Math.abs(n - rounded) < epsilon) return true;
  }
  return false;
}

// AI qaytargan matnni qatorlarga bo'lib, HAR BIR qatorni alohida Word paragraf qiladi.
// Har qatorda ":" bo'lsa — ":" gacha bo'lgan qism (ko'rsatkich nomi) QALIN, qolgani oddiy yoziladi.
// "- ", "* ", "1. " kabi belgi bilan boshlangan qatorlar bullet (•) va otstup bilan ajratiladi.
// originalNumbers berilsa, har bir qatordagi birinchi raqam original hujjatdagi raqamlar
// to'plamida borligi tekshiriladi; mos kelmasa qator qizil rangda va ogohlantirish belgisi
// bilan chiqadi, mismatchCount oshiriladi.
function textToWordParagraphs(
  text: string,
  originalNumbers: Set<number>
): { xml: string; mismatchCount: number } {
  const lines = text
    .split("\n")
    .map((l) => l.trim())
    .filter((l) => l.length > 0);

  let mismatchCount = 0;

  const xml = lines
    .map((line) => {
      const isBullet = /^([-*•]|\d+[.)])\s+/.test(line);
      let cleanLine = line.replace(/^([-*•]|\d+[.)])\s+/, "").replace(/\*\*/g, "");

      const indent = isBullet ? `<w:ind w:left="360"/>` : "";
      const bulletPrefix = isBullet ? "•  " : "";

      const colonIdx = cleanLine.indexOf(":");

      // Ushbu qatordagi "natija" raqamini topib, original hujjat bilan solishtiramiz
      let isFlagged = false;
      const numMatch = cleanLine.match(/:\s*(-?\d+(?:[.,]\d+)?)/);
      if (numMatch) {
        const val = parseFloat(numMatch[1].replace(",", "."));
        if (!isNaN(val) && !numberExistsIn(val, originalNumbers)) {
          isFlagged = true;
          mismatchCount++;
        }
      }

      const warnColor = isFlagged ? "C00000" : undefined;
      const warnPrefix = isFlagged ? "⚠ [TEKSHIRING] " : "";

      let runs: string;
      if (colonIdx > -1 && colonIdx < 60) {
        const label = cleanLine.slice(0, colonIdx + 1);
        const rest = cleanLine.slice(colonIdx + 1);
        runs =
          makeRun(warnPrefix + bulletPrefix + label, true, warnColor) +
          makeRun(rest, false, warnColor);
      } else {
        runs = makeRun(warnPrefix + bulletPrefix + cleanLine, false, warnColor);
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

  return { xml, mismatchCount };
}

// ---------- Asosiy model band bo'lsa, avtomatik zaxira modelga o'tish ----------

const PRIMARY_MODEL = "gemini-3.6-flash";
// gemini-2.5-flash-lite 2026-yil 16-oktabrda to'xtatiladi, shuning uchun yangiroq
// 3.x seriyadagi flash-lite modelga o'tkazildi.
const FALLBACK_MODEL = "gemini-3.5-flash-lite";

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

// Gemini chaqiruvini timeout bilan o'raymiz — server abadiy "osilib qolmasligi" uchun
const AI_TIMEOUT_MS = 30000;

function withTimeout<T>(promise: Promise<T>, ms: number, label: string): Promise<T> {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(`${label} vaqt limiti tugadi (${ms}ms)`)), ms);
    promise.then(
      (val) => {
        clearTimeout(timer);
        resolve(val);
      },
      (err) => {
        clearTimeout(timer);
        reject(err);
      }
    );
  });
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
    const result = await withTimeout(
      primary.generateContent(prompt),
      AI_TIMEOUT_MS,
      PRIMARY_MODEL
    );
    const text = result.response.text().trim();
    if (!text) {
      throw new Error(`${PRIMARY_MODEL} bo'sh javob qaytardi (safety block bo'lishi mumkin)`);
    }
    return { text, modelUsed: PRIMARY_MODEL };
  } catch (err: any) {
    if (!isOverloadError(err) && !String(err?.message).includes("vaqt limiti") && !String(err?.message).includes("bo'sh javob")) {
      throw err;
    }

    console.warn(
      `[AI] ${PRIMARY_MODEL} band/ishlamayapti yoki vaqt tugadi (${err?.message}), ${FALLBACK_MODEL} ga o'tildi.`
    );

    const fallback = genAI.getGenerativeModel({
      model: FALLBACK_MODEL,
      systemInstruction,
    });
    const result = await withTimeout(
      fallback.generateContent(prompt),
      AI_TIMEOUT_MS,
      FALLBACK_MODEL
    );
    const text = result.response.text().trim();
    if (!text) {
      throw new Error(`${FALLBACK_MODEL} ham bo'sh javob qaytardi`);
    }
    return { text, modelUsed: FALLBACK_MODEL };
  }
}

// Jadval qatorlarini alohida saqlab qoladigan matn ajratish — "Ko'rsatkich | Natija |
// Norma | birlik" tuzilishi yo'qolmasin deb, AI natija va normani noto'g'ri bog'lamasin.
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

    const withoutTables = withoutDeletedRuns.replace(/<w:tbl\b[\s\S]*?<\/w:tbl>/g, "\n");
    const outsideMatches = withoutTables.match(/<w:t[^>]*>(.*?)<\/w:t>/g) || [];
    const outsideText = outsideMatches.map((t) => t.replace(/<[^>]+>/g, "")).join(" ").trim();

    return `${outsideText}\n\n${tableText}`.trim();
  }

  const textMatches = withoutDeletedRuns.match(/<w:t[^>]*>(.*?)<\/w:t>/g) || [];
  return textMatches.map((t) => t.replace(/<[^>]+>/g, "")).join(" ");
}

// Fayl haqiqatan laboratoriya tahlili ekanligini tekshiruvchi oddiy filtr —
// tasodifiy/mos kelmaydigan hujjat AI'ga yuborilib, noto'g'ri "tahlil" chiqmasligi uchun
const LAB_KEYWORDS = [
  "натижа",
  "natija",
  "норма",
  "norma",
  "текширилувчи",
  "tekshiriluvchi",
  "курсаткич",
  "ko'rsatkich",
  "ммоль",
  "mmol",
  "мкмоль",
];

function looksLikeLabReport(text: string): boolean {
  const lower = text.toLowerCase();
  return LAB_KEYWORDS.some((kw) => lower.includes(kw.toLowerCase()));
}

const MAX_FILE_SIZE_MB = 15;

export async function POST(req: NextRequest) {
  try {
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      return NextResponse.json(
        { error: "GEMINI_API_KEY topilmadi. .env.local (local) yoki hosting muhitidagi environment variables (production) ni tekshiring." },
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

    if (!looksLikeLabReport(extractedText)) {
      return NextResponse.json(
        {
          error:
            "Fayl laboratoriya tahlil natijasiga o'xshamayapti. Iltimos, to'g'ri hujjatni yuklaganingizni tekshiring.",
        },
        { status: 400 }
      );
    }

    // AI'dan: (1) faqat aniq/berilgan ma'lumotga tayangan, (2) natijasi mavjud
    // HAR BIR ko'rsatkichni to'liq, (3) oddiy odam tushunadigan sodda tilda,
    // (4) tashxis emasligini ochiq aytadigan tahlil so'raladi.
    const systemInstruction =
      "Siz tajribali laboratoriya shifokorisiz. Sizga jadval ko'rinishidagi laboratoriya tahlili " +
      "berilgan; har bir qator 'katakcha1 | katakcha2 | ...' formatida, odatda " +
      "'Ko'rsatkich nomi | Natija | Norma | O'lchov birligi' tartibida keladi.\n\n" +
      "ANIQLIK BO'YICHA QAT'IY QOIDALAR:\n" +
      "- Faqat berilgan matnda aniq ko'rsatilgan sonlar va nomlarga tayaning. Hech narsani " +
      "taxmin qilmang, o'ylab topmang yoki umumlashtirmang.\n" +
      "- Natija qiymatini tegishli norma oralig'i bilan diqqat bilan solishtiring — qaysi qatorga " +
      "tegishli ekanini chalkashtirmang (masalan 'Umumiy', 'Bog'langan', 'Erkin' bilirubin — bularning " +
      "har biri alohida natija va alohida norma oralig'iga ega, ularni aralashtirmang).\n" +
      "- Erkak (Э/А, эркак/аёл) uchun norma alohida ko'rsatilgan bo'lsa va bemor jinsi hujjatda aniq " +
      "bo'lmasa, ikkala normani ham eslatib o'ting, aniq bittasini tanlab olmang.\n\n" +
      "TARKIB BO'YICHA QOIDALAR:\n" +
      "1. Faqat NATIJA QIYMATI ko'rsatilgan (bo'sh bo'lmagan) ko'rsatkichlarni tahlil qiling. Natija " +
      "bo'sh bo'lsa (o'lchov o'tkazilmagan), o'sha ko'rsatkichni butunlay o'tkazib yuboring.\n" +
      "2. Bemor topshirgan HAR BIR natijasi mavjud ko'rsatkich uchun ALOHIDA qator yozing — " +
      "birortasini ham tashlab ketmang, birortasini ham birlashtirmang.\n" +
      "3. Har bir qatorda: ko'rsatkich nomi, natija qiymati (birligi bilan), me'yor bilan solishtirilgan " +
      "holati (me'yorda / me'yordan yuqori / me'yordan past). Me'yorda bo'lsa — shunchaki 'me'yorda' " +
      "deb yozing. Chetlashgan bo'lsa — bu ko'rsatkich odatda nimani anglatishi haqida oddiy, sodda " +
      "tilda 1 qisqa jumla qo'shing (tashxis qo'ymasdan, faqat umumiy tushuntirish sifatida).\n" +
      "4. Tilni ODDIY ODAM tushunadigan darajada sodda tuting — tibbiy atamalarni ishlatgan bo'lsangiz, " +
      "qavs ichida bir og'iz so'z bilan izohlang (masalan 'kreatinin (buyrak faoliyati ko'rsatkichi)'). " +
      "Gaplarni cho'zmang, har bir qator bitta qisqa qatordan oshmasin.\n" +
      "5. Bu XULOSA TASHXIS EMASLIGINI hech qachon unutmang — faqat natijalarni tushuntirasiz, " +
      "kasallik nomini aytmaysiz yoki davolash tayinlamaysiz.\n" +
      "6. Faqat oddiy matn (markdown, **, # belgilarisiz) bilan javob bering.\n\n" +
      "JAVOB FORMATI (qat'iy shu tartibda):\n\n" +
      "1-qator: 'Qon: ' bilan boshlanadigan bitta gaplik umumiy xulosa, sodda tilda " +
      "(masalan: 'Qon: N ta ko'rsatkich tekshirilgan, ulardan M tasi me'yorda, K tasida me'yordan " +
      "chetlashish bor.').\n\n" +
      "Keyingi qatorlar: natijasi mavjud HAR BIR ko'rsatkich uchun bittadan qator, " +
      "'- Ko'rsatkich nomi: natija (birlik) — holati, [chetlashsa qisqa sodda izoh]' formatida.\n\n" +
      "Oxirgi qator: 'Tavsiya: ' bilan boshlanadigan bir gaplik amaliy tavsiya — bu tashxis emasligini " +
      "va aniq baho uchun shifokorga murojaat qilish kerakligini eslatib, qaysi yo'nalishdagi shifokorga " +
      "(masalan terapevt, gastroenterolog, nefrolog) borish maqsadga muvofiqligini ayting. Hech narsa " +
      "chetlashmagan bo'lsa — shunchaki profilaktik ko'rikni davom ettirish tavsiyasini bering.";

    const prompt =
      `Quyidagi laboratoriya tahlil jadvalini diqqat bilan, aniqlikka rioya qilgan holda o'rganib chiqing. ` +
      `Bemor topshirgan HAR BIR natijasi mavjud ko'rsatkichni to'liq, aniq va oddiy odam tushunadigan ` +
      `sodda tilda tushuntiring. Bu tashxis emasligini yodda tuting:\n\n${extractedText}`;

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

    // AI natijasidagi raqamlarni original jadval matnidagi raqamlar bilan solishtiramiz
    const originalNumbers = extractNumbersFromText(extractedText);
    const { xml: bodyParagraphs, mismatchCount } = textToWordParagraphs(
      analysisResult,
      originalNumbers
    );

    const mismatchWarningXml =
      mismatchCount > 0
        ? `
      <w:p/>
      <w:p>
        <w:pPr>
          <w:jc w:val="left"/>
        </w:pPr>
        <w:r>
          <w:rPr>
            <w:b/>
            <w:color w:val="C00000"/>
            <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
            <w:sz w:val="22"/>
          </w:rPr>
          <w:t xml:space="preserve">⚠ Avtomatik tekshiruv: ${mismatchCount} ta qatordagi raqam original hujjatdagi qiymatlar bilan aniq mos kelmadi (yuqorida qizil rangda belgilangan). Iltimos, ularni qo'lda tekshiring.</w:t>
        </w:r>
      </w:p>`
        : "";

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
      ${mismatchWarningXml}
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
          <w:t>Ushbu xulosa Sun'iy Intellekt (AI) tomonidan tayyorlangan va TASHXIS HISOBLANMAYDI — u faqat tahlil natijalarini tushunishga yordam beradi. Aniq baho va davolanish uchun albatta mutaxassis shifokorga murojaat qiling.</w:t>
        </w:r>
      </w:p>
    `;

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
        "X-AI-Mismatch-Count": String(mismatchCount),
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
