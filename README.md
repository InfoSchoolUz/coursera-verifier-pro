# 🎓 Coursera Certificate Verifier Pro

> **Maktab o'quvchilarining Coursera sertifikatlarini avtomatik tekshiruvchi va hisobot tuzuvchi tizim**

🌐 **Onlayn demo:** [coursera-verifier.streamlit.app](https://coursera-verifier.streamlit.app/)

---

## 📌 Loyiha nima qiladi?

Bu dastur maktab o'quvchilari topshirgan **Coursera sertifikat havolalarini** Excel fayldan o'qib, har birini internetda **avtomatik tekshiradi** va natijani yangi Excel faylga yozib chiqaradi.

**Muammo:** 500 ta o'quvchining sertifikatini qo'lda tekshirish — soatlab vaqt.  
**Yechim:** Dastur bu ishni bir necha daqiqada avtomatik bajaradi.

---

## 🗂 Loyiha tuzilishi

```
coursera-verifier-pro-main/
│
├── coursera_pro.py          ← Asosiy dastur (Streamlit ilovasi)
├── run_app.py               ← Dasturni ishga tushirish skripti
├── requirements.txt         ← Kerakli kutubxonalar ro'yxati
├── coursera_template.xlsx   ← Namuna Excel fayl (shablon)
├── .devcontainer/
│   └── devcontainer.json    ← GitHub Codespaces sozlamalari
├── .gitignore
└── LICENSE
```

---

## ⚙️ Texnologiyalar

| Kutubxona | Vazifasi |
|-----------|----------|
| **Streamlit** | Veb-interfeys (UI) yaratish |
| **Pandas** | Excel/CSV fayllarni o'qish va qayta ishlash |
| **Requests** | Internet orqali Coursera sahifalariga ulanish |
| **BeautifulSoup4** | HTML sahifasidan sertifikat sanasini ajratib olish |
| **OpenPyXL** | Natija Excel faylini yaratish |

---

## 🚀 Ishga tushirish (Lokal)

### 1. Talablar

- Python 3.8 yoki undan yuqori
- `pip` o'rnatilgan bo'lishi kerak

### 2. Kutubxonalarni o'rnatish

> ⚠️ `requirements.txt` da kichik xato bor: `openpyxl` va `beautifulsoup4` o'rtasida bo'sh qator yo'q. Quyidagi buyruqni ishlating:

```bash
pip install streamlit pandas requests beautifulsoup4 openpyxl
```

### 3. Dasturni ishga tushirish

```bash
streamlit run coursera_pro.py
```

Brauzer avtomatik ochiladi: `http://localhost:8501`

---

## 📊 Dastur qanday ishlaydi — Qadam-baqadam

```
1. Foydalanuvchi Excel fayl yuklaydi
        ↓
2. Dastur Excel'dagi barcha listlarni o'qiydi
        ↓
3. Foydalanuvchi qaysi listni tekshirishni tanlaydi
        ↓
4. Dastur Coursera havolalarini topadi
        ↓
5. Har bir havola parallel (bir vaqtda) tekshiriladi
        ↓
6. Har bir sertifikat uchun natija yoziladi
        ↓
7. Foydalanuvchi _Verify.xlsx faylini yuklab oladi
```

---

## 📥 Excel fayl talablari

Yuklanadigan fayl quyidagi **9 ta ustun**ga ega bo'lishi shart:

| # | Ustun nomi | Izoh |
|---|-----------|------|
| 1 | № | Tartib raqami |
| 2 | Туман шаҳар номи | Tuman/shahar |
| 3 | Мактаб рақами | Maktab raqami |
| 4 | Синфи | O'quvchi sinfi |
| 5 | Ўқувчининг ФИШ | Familiya Ism Sharif |
| 6 | Гувоҳнома серия ва рақами | Guvohnoma |
| 7 | Туғилган куни | Tug'ilgan sana |
| 8 | Сертификат ҳаволаси | ⭐ Coursera linki (asosiy) |
| 9 | Электрон почтаси | Email manzil |

> **Muhim:** Dastur ustun nomlarini avtomatik taniydi. Nomlarni o'zgartirmang!

---

## 📋 Natija fayli nimalardan iborat?

Tekshiruv tugagach, `fayl_nomi_Verify.xlsx` fayliga quyidagi ma'lumotlar yoziladi:

| Ustun | Mazmuni |
|-------|---------|
| F.I.SH | O'quvchi ismi |
| Kurs yo'nalishi | Qaysi kurs ustunidan olingan |
| Holati | `MAVJUD` / `XATO` / `MAVJUD EMAS` |
| Natija | Batafsil xabar yoki `TAKRORLANUVCHI 🔄` |
| Havola | Asl Coursera linki |
| Sertifikat kodi | Linkdan ajratilgan unikal kod |
| Sertifikat olingan sana | Sahifadan avtomatik topilgan sana |

### Holat turlari:

| Holat | Rang | Ma'nosi |
|-------|------|---------|
| ✅ `MAVJUD` | Yashil | Sertifikat haqiqiy va faol |
| ❌ `XATO` | Qizil | Link ishlamaydi yoki login so'raydi |
| 🟡 `MAVJUD EMAS` | Sariq | Coursera sahifasi emas |
| 🔵 `TAKRORLANUVCHI` | Ko'k | Bir xil sertifikat bir necha marta ishlatilgan |

---

## 🧠 Dasturning texnik ishlash tartibi

### Sertifikat kodi qanday aniqlanadi?

Dastur Coursera linkidan unikal kodni ajratib oladi:

```
https://coursera.org/share/abc123xyz   →  kod: abc123xyz
https://coursera.org/verify/def456     →  kod: def456
```

Bu kod yordamida **takrorlanuvchi sertifikatlar** aniqlanadi.

### Parallel tekshiruv nima?

Bir vaqtda bir nechta havolani tekshirish imkonini beradi. Masalan:

- **25 parallel kanal** → 500 linkni ~2-3 daqiqada tekshiradi
- **5 parallel kanal** → sekinroq, lekin server blok qilishi ehtimoli kam

### Sana qanday topiladi?

Dastur Coursera sahifasining HTML matnidan sanani qidiradi:

```
"January 15, 2024"  →  qaytaradi
"2024-01-15"        →  qaytaradi
```

---

## ⚡ Ishlash tavsiyalari

| Linklar soni | Parallel tekshiruvlar | Kutish vaqti |
|-------------|----------------------|-------------|
| 100–300 | 25–30 | 15 sek |
| 300–500 | 10–20 | 15 sek |
| 500–1000 | 5–10 | 20 sek |
| 1000+ | 3–5 | 25–30 sek |

---

## ⚠️ Cheklovlar

- Coursera ko'p so'rov kelsa, vaqtincha **blok qilishi mumkin** → parallel sonini kamayting
- Katta fayllar (1000+ o'quvchi) ko'proq vaqt oladi
- Internet tezligi natijaga ta'sir qiladi
- Agar `MAVJUD EMAS` ko'p chiqsa → parallel sonini **3–5** ga tushiring

---

## 🐛 Bilinom xatolar

| Xato | Sababi | Yechim |
|------|--------|--------|
| `requirements.txt` o'rnatilmaydi | `openpyxl` va `bs4` birikib yozilgan | Alohida `pip install` qiling |
| Ko'p `MAVJUD EMAS` | Coursera blok qilgan | Parallel sonini kamaytiring |
| Sana topilmadi | Coursera sahifasi yopiq | Normal holat, sertifikat amal qilishi mumkin |

---

## 👨‍💻 Muallif

**Azamat Madrimov**  
Informatika va AT o'qituvchisi  
📧 azamat3533141@gmail.com  
📱 [@futurex_azamat](https://t.me/futurex_azamat)

---

## 📄 Litsenziya

MIT License — erkin foydalanish, o'zgartirish va tarqatish mumkin.

---

*© 2026 Azamat Madrimov — Ta'lim muassasalari uchun ishlab chiqilgan*
