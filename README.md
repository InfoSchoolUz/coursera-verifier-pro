# 🎓 Coursera Sertifikat Tekshiruvchi Pro

Excel fayllar orqali Coursera sertifikatlarini avtomatik tekshiruvchi va hisobot shakllantiruvchi tizim.

---

## 🌐 Onlayn demo

👉 https://coursera-verifier.streamlit.app/

---

## 🚀 Asosiy imkoniyatlar

* ✅ Coursera sertifikat linklarini avtomatik tekshirish
* 📊 Excel formatida professional hisobot chiqarish
* 📅 Sertifikat olingan sanani aniqlash
* 🔁 Takrorlanuvchi sertifikatlarni aniqlash
* 📂 Excel ichidagi bir nechta listlarni qo‘llab-quvvatlash
* 🎯 Faqat bitta listni tanlab tekshirish imkoniyati
* ⚡ Parallel (tezkor) tekshiruv tizimi
* 📥 Natijani avtomatik yuklab olish (`_Verify.xlsx`)

---

## 📥 Qanday ishlaydi

### 1. Fayl yuklash

* Excel (.xlsx) yoki CSV fayl yuklanadi
* Faylda bir nechta list bo‘lishi mumkin

### 2. List tanlash

* Tekshiriladigan list tanlanadi

### 3. Tekshirishni boshlash

* **"TEKSHIRISHNI BOSHLASH"** tugmasi bosiladi

### 4. Natijani yuklab olish

* Hisobot quyidagi nom bilan yuklanadi:

```
fayl_nomi_Verify.xlsx
```

---

## 📌 Excel fayl talablari (MAJBURIY)

Yuklanadigan Excel fayl quyidagi ustunlarga ega bo‘lishi shart:

| № | Ustun nomi                |
| - | ------------------------- |
| 1 | №                         |
| 2 | Туман шаҳар номи          |
| 3 | Мактаб рақами             |
| 4 | Синфи                     |
| 5 | Ўқувчининг ФИШ            |
| 6 | Гувоҳнома серия ва рақами |
| 7 | Туғилган куни             |
| 8 | Сертификат ҳаволаси       |
| 9 | Электрон почтаси          |

---

## ⚠️ Muhim qoidalar

* Sertifikat havolasi ustunida **faqat Coursera link** bo‘lishi kerak
* Faylda kamida **1 ta to‘g‘ri havola** bo‘lishi shart
* Bo‘sh qatorlar bo‘lmasligi tavsiya etiladi
* Noto‘g‘ri format → noto‘g‘ri natija

---

## 📊 Hisobot tarkibi

Natija Excel faylida quyidagilar chiqadi:

* F.I.SH
* Kurs yo‘nalishi
* Holati (MAVJUD / XATO / MAVJUD EMAS)
* Natija
* Sertifikat havolasi
* Sertifikat kodi
* Sertifikat olingan sana

---

## ⚙️ Texnologiyalar

* Python
* Streamlit
* Pandas
* Requests
* BeautifulSoup

---

## ⚡ Ishlash bo‘yicha tavsiyalar

| Linklar soni | Tavsiya      |
| ------------ | ------------ |
| 100–300      | Threads: 5   |
| 500–1000     | Threads: 3–5 |
| 2000+        | Threads: 3   |

---

## ⚠️ Cheklovlar

* Coursera juda ko‘p so‘rov bo‘lsa vaqtincha blok qilishi mumkin
* Internet tezligi natijaga ta’sir qiladi
* Juda katta fayllar sekin ishlashi mumkin

---

## 💡 Tavsiya

Agar natijada ko‘p "MAVJUD EMAS" chiqsa:

👉 Parallel tekshiruv sonini kamaytiring
👉 Kichik bo‘laklarga bo‘lib tekshiring

---

## 👨‍💻 Muallif

**Azamat Madrimov**
Informatika va AT mutaxassisi

---

## 📈 Loyiha haqida

Ushbu loyiha ta’lim muassasalari uchun:

* o‘quvchilar sertifikatlarini tez va aniq tekshirish
* qo‘lda tekshirish vaqtini kamaytirish
* xatoliklarni minimallashtirish

uchun ishlab chiqilgan.
