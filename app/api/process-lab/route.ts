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
// TUZATILDI: eski "gemini-1.5-flash" allaqachon butunlay o'chirilgan (har doim 404 qaytaradi),
// shuning uchun zaxira sifatida ishlamas edi. O'rniga hozirgi barqaror va BEPUL (free-tier)
// "gemini-2.5-flash-lite" modeli qo'yildi.
const FALLBACK_MODEL = "gemini-2.5-flash-lite";

// TUZATILDI: avval faqat 429/503 (band bo'lish) holatlarini "overload" deb hisoblardi.
// Agar asosiy model nomi noto'g'ri bo'lib qolsa yoki o'chirilgan bo'lsa, API odatda 404/
// "not found" xatosi qaytaradi — bu ilgari umuman ushlanmas va fallback ishga tushmas edi.
// Endi bunday holatlar ham fallbackni ishga tushiradi.
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
      throw err; // boshqa turdagi xato — fallback yordam bermaydi, to'g'ridan-to'g'ri uloqtiramiz
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

// TUZATILDI: DOCX dan matn ajratib olishda endi "o'chirilgan" (Track Changes -> w:del)
// qatorlar hisobga olinmaydi. Avval oddiy regex <w:del> ichidagi matnni ham ushlab olar
// va bemor tahliliga aloqasi bo'lmagan/o'chirilgan matn AI'ga yuborilar edi.
function extractTextFromDocumentXml(docXml: string): string {
  const withoutDeletedRuns = docXml.replace(/<w:del\b[^>]*>[\s\S]*?<\/w:del>/g, "");
  const textMatches = withoutDeletedRuns.match(/<w:t[^>]*>(.*?)<\/w:t>/g) || [];
  return textMatches.map((t) => t.replace(/<[^>]+>/g, "")).join(" ");
}

// Ruxsat etilgan maksimal fayl hajmi (MB). Serverga haddan tashqari katta yoki
// soxta fayl yuklanishining oldini olish uchun.
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

    // TUZATILDI: avval "file" maydoni to'g'ridan-to'g'ri File deb hisoblanardi.
    // Agar maydon umuman kelmasa yoki boshqa turda bo'lsa, keyinroq tushunarsiz
    // xato chiqar edi. Endi aniq va tushunarli tekshiruv qo'yildi.
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

    const systemInstruction =
      "Siz tajribali laboratoriya shifokorisisiz. Berilgan laboratoriya tahlili natijalarini o'rganib, " +
      "quyidagi QAT'IY formatda, faqat oddiy matn ko'rinishida (markdown, **, # belgilarisiz) javob bering:\n\n" +
      "1-qator: Umumiy xulosa — 'Qon: ' bilan boshlanadigan bitta-ikkita gapli umumiy baho " +
      "(masalan: 'Qon: Qonning biokimyoviy tahlili natijalariga ko'ra, asosiy ko'rsatkichlar me'yorda bo'lib, faqat X ko'rsatkichida chetlashish aniqlandi.').\n\n" +
      "Keyingi qatorlar: me'yordan chetlashgan yoki alohida e'tibor talab qiladigan HAR BIR ko'rsatkich uchun " +
      "ALOHIDA QATORDAN boshlab, '- Ko'rsatkich nomi: natija qiymati, me'yor bilan solishtirilgan holda, qisqa tibbiy izoh' " +
      "formatida yozing. Nechta ko'rsatkich chetlashgan yoki muhim bo'lsa, shunchasini alohida qatorda bering — " +
      "hech birini birlashtirmang.\n\n" +
      "Oxirgi qator: 'Tavsiya: ' bilan boshlanadigan qisqa amaliy tavsiya (qaysi shifokorga murojaat qilish kerakligi).\n\n" +
      "Har bir ko'rsatkich nomini aniq va tibbiy jihatdan to'g'ri yozing, taxmin qilmang — faqat berilgan ma'lumotlarga tayaning.";

    const prompt = `Quyidagi laboratoriya tahlil natijasini o'rganib chiqib, yuqorida ko'rsatilgan formatga qat'iy rioya qilgan holda xulosa yozing:\n\n${extractedText}`;

    let analysisResult: string;
    let modelUsed: string;
    try {
      const result = await generateWithFallback(genAI, systemInstruction, prompt);
      analysisResult = result.text;
      modelUsed = result.modelUsed;
    } catch (aiErr: any) {
      // Ikkala model (asosiy va zaxira) ham javob berolmadi — server logida to'liq
      // xatoni saqlaymiz, lekin laborantga tushunarli, qisqa xabar chiqaramiz.
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

    // Shablon (Analiz_xulosasi_shablon.docx) bilan bir xil formatda:
    // markazlashtirilgan qalin sarlavha + sz=28 li paragraflar, DIQQAT bloki eng pastda
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

    // MUHIM: <w:sectPr> (sahifa/bo'lim sozlamalari) w:body ichidagi ENG OXIRGI element
    // bo'lishi SHART (OOXML standarti). Agar yangi kontentni shundan keyin qo'shsak,
    // fayl tuzilishi buziladi va Word uni tartibsiz/xato ko'rsatadi yoki "repair" qiladi.
    // Shuning uchun kontentni sectPr'dan OLDIN joylashtiramiz.
    const sectPrRegex = /(<w:sectPr\b[^>]*\/>|<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>)(\s*<\/w:body>)/;
    let updatedXml: string;
    if (sectPrRegex.test(docXml)) {
      updatedXml = docXml.replace(sectPrRegex, `${xmlToInsert}$1$2`);
    } else {
      // sectPr topilmasa (kamdan-kam holat), eski usulda body oxiriga qo'shamiz
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
    // TUZATILDI: avval err.message to'g'ridan-to'g'ri klientga qaytarilardi — bu server
    // ichki tuzilishi haqida keraksiz ma'lumot oshkor qilishi mumkin edi. Endi to'liq xato
    // faqat serverda loglanadi, klientga esa umumiy va xavfsiz xabar boriladi.
    console.error("[process-lab] Kutilmagan xato:", err);
    return NextResponse.json(
      { error: "Faylni qayta ishlashda kutilmagan xatolik yuz berdi" },
      { status: 500 }
    );
  }
}
