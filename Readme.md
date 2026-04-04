# 🎓 Coursera Certificate Verifier Pro

Maktab o'quvchilarining Coursera sertifikatlarini **avtomatik tekshirish** tizimi.  
Excel faylni yuklang — tizim har bir sertifikat havolasini tekshirib, ism mosligini tahlil qilib, natijani yangi Excel faylda qaytaradi.

🔗 **Demo:** [coursera-verifier.streamlit.app](https://coursera-verifier.streamlit.app)

---

## ⚙️ Imkoniyatlar

- ✅ Coursera sertifikat havolalarini real vaqtda tekshirish
- 👤 Sertifikatdagi ism va Excel dagi F.I.SH ni avtomatik solishtirish
- 🔄 Takrorlanuvchi sertifikatlarni aniqlash
- 📅 Sertifikat olingan sanani ajratib olish
- ⚡ Parallel tekshiruv (bir vaqtda 5–50 ta havola)
- 📊 Natijani rangli jadval va yuklab olinadigan Excel shaklida taqdim etish
- 🌐 Kirill va lotin imlo farqlarini avtomatik hisobga olish (fuzzy matching)

---

## 🖥️ Ishlatish

### 1. Reponi klonlash

```bash
git clone https://github.com/InfoSchoolUz/coursera-verifier.git
cd coursera-verifier
```

### 2. Kutubxonalarni o'rnatish

```bash
pip install -r requirements.txt
```

### 3. Ilovani ishga tushirish

```bash
streamlit run coursera_pro.py
```

---

## 📁 Fayl tuzilmasi

```
coursera-verifier/
├── coursera_pro.py       # Asosiy ilova
├── run_app.py            # PyInstaller uchun ishga tushiruvchi
├── requirements.txt      # Kerakli kutubxonalar
├── coursera_template.xlsx # Namuna Excel fayl
└── README.md
```

---

## 📋 Excel fayl formati

Yuklanadigan Excel faylda quyidagi ustunlar bo'lishi lozim:

| № | Tuman/Shahar | Maktab raqami | Sinf | F.I.SH | Guvohnoma seriyasi | Tug'ilgan sana | Sertifikat havolasi | Elektron pochta |
|---|---|---|---|---|---|---|---|---|

> ⚠️ Ustun nomlarini o'zgartirmang — tizim ularni avtomatik taniydi.

Namuna faylni ilova ichidan yuklab olishingiz mumkin.

---

## 📊 Natija ustunlari

| Ustun | Tavsif |
|---|---|
| `Holati` | MAVJUD / MAVJUD EMAS / XATO |
| `Natija` | Tasdiqlandi, Redirect, Takrorlanuvchi va h.k. |
| `Ism Moslik` | MOS ✅ / QISMAN MOS ⚠️ / MOS EMAS ❌ |
| `Moslik Tafsiloti` | Farq yoki moslik sababi |
| `Sertifikatdagi Ism` | Sahifadan olingan ism |
| `Sertifikat olingan sana` | Avtomatik ajratilgan sana |
| `Sertifikat kodi` | URL dan ajratilgan unikal kod |

---

## 🛠️ Texnologiyalar

| Texnologiya | Maqsad |
|---|---|
| [Streamlit](https://streamlit.io) | Web interfeys |
| [Pandas](https://pandas.pydata.org) | Excel/CSV bilan ishlash |
| [BeautifulSoup4](https://beautiful-soup-4.readthedocs.io) | HTML tahlil qilish |
| [Requests](https://docs.python-requests.org) | HTTP so'rovlar |
| [difflib](https://docs.python.org/3/library/difflib.html) | Fuzzy ism solishtirish |
| [ThreadPoolExecutor](https://docs.python.org/3/library/concurrent.futures.html) | Parallel tekshiruv |

---

## 👤 Muallif

**Azamat Madrimov**  
📧 [azamat3533141@gmail.com](mailto:azamat3533141@gmail.com)  
💬 Telegram: [@futurex_azamat](https://t.me/futurex_azamat)  
🏫 [InfoSchoolUz](https://github.com/InfoSchoolUz)

---

## 📄 Litsenziya

MIT License — erkin foydalanishingiz mumkin.
