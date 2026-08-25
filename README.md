# Lab AI — brauzerda ishlaydigan tahlil izohlovchi

Laborant `.docx` tahlil faylini brauzerga tashlaydi → tizim natijalarni normativ
ko'rsatkichlar bilan avtomatik solishtiradi → me'yordan chetlangan ko'rsatkichlar
uchun izoh yozadi → tayyor faylni yuklab olish tugmasi chiqadi. Boshqa hech qanday
qadam yo'q.

## Ishga tushirish (lokal, o'z kompyuteringizda)

```bash
npm install
cp .env.local.example .env.local   # ixtiyoriy — pastga qarang
npm run dev
```

Brauzerda oching: **http://localhost:3000**

## AI izoh haqida

`.env.local` faylida `ANTHROPIC_API_KEY` bo'sh bo'lsa, tizim ishonchli, oldindan
yozilgan shablon izohlardan foydalanadi (offline ham ishlaydi). Haqiqiy Claude AI
yozgan, tabiiyroq izoh olish uchun [console.anthropic.com](https://console.anthropic.com)
dan kalit oling va shu faylga yozing:

```
ANTHROPIC_API_KEY=sk-ant-...
```

## Qanday ishlaydi (arxitektura)

```
Brauzer (app/page.tsx)
   │  fayl tashlanadi (drag & drop)
   ▼
POST /api/process-lab  (app/api/process-lab/route.ts)
   │  fayl serverda qabul qilinadi
   ▼
lib/processLab.ts
   1. .docx ichidagi jadvalni o'qiydi (jszip)
   2. "Натижа" ustunini "Норма" bilan solishtiradi
   3. Chetlangan har bir ko'rsatkich uchun AI (yoki shablon) izoh yozadi
   4. Asl hujjat oxiriga yangi bo'lim qo'shib, tayyor faylni qaytaradi
   ▼
Brauzerda "Yuklab olish" tugmasi chiqadi
```

`ANTHROPIC_API_KEY` faqat serverda (`.env.local`) saqlanadi — brauzerga hech qachon
yuborilmaydi, shuning uchun xavfsiz.

## Real klinikaga joriy qilishda e'tiborga olinadigan narsalar

- **Shifokor tasdiqlashi** — hozirgi holatda natija to'g'ridan-to'g'ri bemorga emas,
  avval laborant/shifokor ko'rib chiqishi uchun mo'ljallangan. Bemorga avtomatik
  yuborishdan oldin tasdiqlash bosqichini qo'shish tavsiya etiladi.
- **Jins/yoshga bog'liq murakkab normalar** (masalan "Э: <40, А: <31") hozircha
  avtomatik hisoblanmaydi — "qo'lda tekshirish kerak" deb alohida ko'rsatiladi,
  chunki bemor kartasida jins/yosh ma'lumoti kerak bo'ladi. Buni `lib/processLab.ts`
  ichidagi `parseNorm` funksiyasini kengaytirib qo'shish mumkin.
- **Joylashtirish (deploy)** — Vercel'da bemalol ishlaydi (`vercel deploy`), yoki
  boshqa Node.js serverida (`npm run build && npm start`).

## Loyihaning tarkibi

```
app/
  page.tsx                  — fayl tashlash sahifasi
  layout.tsx                — asosiy HTML qobiq
  globals.css                — dizayn
  api/process-lab/route.ts  — fayl qabul qiluvchi backend
lib/
  processLab.ts              — asosiy mantiq (parsing, solishtirish, AI izoh)
```
