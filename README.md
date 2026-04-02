# 🎓 Coursera Certificate Verifier Pro

## 🇺🇿 O'zbekcha tavsif

Coursera Certificate Verifier Pro — bu Excel fayldagi Coursera sertifikat havolalarini avtomatik tekshiruvchi tizim.

U quyidagi imkoniyatlarga ega:
- ✅ Sertifikat mavjudligini tekshiradi
- 📅 Sertifikat olingan sanani aniqlaydi
- 🔁 Takrorlanuvchi (duplicate) sertifikatlarni topadi
- 📊 Excel hisobot yaratadi
- 📂 Excel fayldagi barcha sheetlarni avtomatik tekshiradi
- ⚡ Parallel (tezkor) tekshiruv

---

## 🚀 Qanday ishlaydi

1. Excel fayl yuklanadi
2. Har bir sheet ichidagi Coursera linklar topiladi
3. Har bir link tekshiriladi:
   - mavjud / mavjud emas / xato
4. Natijalar Excel hisobotga yoziladi
5. Har bir sheet uchun alohida natija chiqariladi

---

## 📁 Kiruvchi fayl talablari

- Excel (.xlsx) yoki CSV
- Sertifikat linklari quyidagi formatda bo‘lishi mumkin:
  - `https://coursera.org/share/...`
  - `https://coursera.org/verify/...`
  - `https://coursera.org/account/accomplishments/...`

---

## 📤 Natija (Output)

Yangi Excel fayl:
- `JAMI_HISOBOT`
- original sheet nomlari bilan alohida listlar

Har bir satr:
- F.I.SH
- Kurs yo‘nalishi
- Holati
- Natija
- Havola
- Sertifikat kodi
- Sertifikat olingan sana

---

## ⚙️ O‘rnatish (Local)

```bash
pip install -r requirements.txt
streamlit run coursera_pro.py
