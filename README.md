![logo jpg](banner.jpg "Logo")

<div align="center">
  <h1>Python va Django framework orqali excel fayllar bilan ishlash</h1>
</div>

<div align="center">
  Ushbu video seriyamizda excel fayllar bilan ishlashni o'rganamiz. Excel fayldagi ma'lumotlarni o'qiymiz, o'qilgan ma'lumotlar asosida Djangodagi ma'lumotlar omborida objectlar yaratamiz. Objectlar yaratish vaqtida filter, zip, map kabi funksiyalar bilan tanishamiz va ushbu funksiyalar yordamida ma'lumotlarimizni filterlaymiz. Filterlangan ma'lumotlarni bulk_create methodi yordamida djangodagi ma'lumotlar omboriga yozamiz. Djangodagi ma'lumotlar omboridan ma'lumotlarni qayta filterlab olib excel faylga yozamiz. Excel fayldagi yacheyka va ustunlarga stil berib, ma'lumotlarni mavjud faylga va yangi faylga yozishni o'rganamiz. Excel faylni ma’lumotlar omboriga saqlashdan tashqari foydalanuvchiga virtual holatda qanday jo’natish mumkinligini ko’rib chiqamiz. FileResponse yordamida virtual yaratilgan faylimizni foydalanuvchimizga taqdim etamiz. Turli xildagi hisobotlar yaratib foydalauvchiga taqdim etish vaqtida huddi shu usuldan foydalanishimiz kerak, sababi har bir hisobotni ma’lumotlar omobirga saqlaydigan bo’lsak, ma’lumotlar omborimiz keraksiz ma’lumotlar bilan to’lib toshadi. Keyingi videodarsimizda esa excel fayllarni email orqali jo’natishni o’rganamiz. Ushbu dars jarayonida biz emailimizni xabar jo’natish uchun sozlab olamiz. Sozlab olingan email ishlayotganligini testdan o’tkazamiz. Undan so’ng oldingi darslarda yaratilgan excel faylimizni virtual holatdan bytes turiga o’tkazamiz. Excel faylimizni Django orqali jo’natish uchun kerakli sozlamalarni amalga oshiramiz va SmtpServer classini yaratib olamiz. Ushbu class yordamida excel ko’rinishidagi hisobotimizni superuserlarimizga jo’natamiz.
</div>

<br>

<div align="center">
  Webdasturlashga oid ko'proq yangiliklardan xabardor bo'lish uchun bizni kuzatib boring: <br>
  <a href="https://www.youtube.com/channel/UCeJ6Sc3SaKKArAurnCwlJBw">YouTube</a>
  <span> | </span>
  <a href="https://www.instagram.com/yakubovdeveloper">Instagram</a>
  <span> | </span>
  <a href="https://www.facebook.com/yakubovdeveloper">Facebook</a>
  <span> | </span>
  <a href="https://www.tiktok.com/@yakubovdeveloper">TikTok</a>
  <span> | </span>
  <a href="https://t.me/yakubovdeveloper">Telegram</a>
</div>

## Video seriyamiz qismlari va mazmuni

<details>
<summary><b>Part-1 Python orqali excel fayldagi ma’lumotlarni o’qish.</b>
</summary>
<br>
<ul>
    <li>Openpyxl paketini o’rnatamiz.</li>
    <li>Exceldagi listlar haqida tushunchaga ega bo’lamiz.</li>
    <li>Yacheykadagi ma’lumotlarni o’qishni o’rganamiz.</li>
    <li>Diapazon bo’yicha yacheykalarni o’qishni o’rganamiz.</li>
    <li>iter_cols, iter_rows, cell methodlari bilan tanishamiz va amaliyotda tekshirib ko’ramiz.</li>
    <li>values, columns, rows generatorlari bilan tanishamiz va amaliyotda tekshirib ko’ramiz.</li>
    <li>Ma’lumotlarni dict holatiga o’tkazamiz va ma’lumotlar omboriga yozish uchun qulay ko’rinishga keltiramiz.</li>
    <li>Ma’lumotlarni filter funksiyasi orqali filterlaymiz.</li>
</ul>

To'liq video qo'llanma bu yerda: https://www.youtube.com/watch?v=Li8FNtZ5wJQ

</details>

<details>
<summary><b>Part-2 Python orqali excel fayldagi ma’lumotlarni saqlash.</b>
</summary>
<br>
<ul>
    <li>Excel fayldagi ma'lumotlarni o'qiymiz.</li>
    <li>zip funksiyasi orqali kerakli kalit so'zlar bilan qiymatlarni o'qishga qulay ko'rinishga keltiramiz.</li>
    <li>Olingan ma'lumotlarni map, filter funksiyalarini ishlatgan holda filterlaymiz.</li> 
    <li>Filterlangan ma'lumotlarni django ma'lumotlar omboriga yozish uchun tayyorlaymiz.</li>
    <li>bulk_create methodi yordamida ma'lumotlarni tezkorlik bilan django ma'lumot omboriga yozamiz.</li>
</ul>

To'liq video qo'llanma bu yerda: https://www.youtube.com/watch?v=dH2kIfaa6i8

</details>

<details>
<summary><b>Part-3 Python orqali excel faylga ma'lumot yozish.</b>
</summary>
<br>
<ul>
    <li>Ma'lumotlar omboridan django orm orqali foydalanuvchilar ma'lumotlarini filterlab olamiz.</li>
    <li>Mavjud excel faylni ochamiz va yangi sheet yaratamiz.</li>
    <li>Excel faylning ustunlari va yacheykalariga stil beramiz.</li>
    <li>Excel faylga filterlangan ma'lumotlarni yozamiz.</li>
    <li>Yangi excel fayl yaratamiz va yangi faylga ma'lumot yozishni ham o'rganamiz. </li>
</ul>

To'liq video qo'llanma bu yerda: https://www.youtube.com/watch?v=w--Ayie81ko

</details>


<details>
<summary><b>Part-4 Python orqali excel faylga formula yozish va faylni yuklab olish.</b></summary>
<br>

<ul>
    <li>Openpyxl paketi yordamida virtual excel fayl yaratamiz</li>
    <li>Excel faylimiz sheetiga nom beramiz</li>
    <li>Excel fayl ustunlari hajimlarini kattalashtiramiz</li>
    <li>Excel fayl ustunlariga nom beramiz</li>
    <li>Djangodagi foydalanuvchilarimizning oylik ish xaqlarini excel faylga yozamiz</li>
    <li>Foydalanuvchilarning oylik ish xaqilarining jami summasini hisoblash uchun python orqali excel faylga formula yozamiz</li>
    <li>openpyxl paketining merge_cells methodi yordamida bir nechta yacheykalarni birlashtiramiz</li>
    <li>Django FileResponse orqali attachment ko’rnishda faylni foydalanuvchiga taqdim etamiz</li>
</ul>

To'liq video qo'llanma bu yerda: https://www.youtube.com/watch?v=nQNWsJ7c_9c

</details>


<details>
<summary><b>Part-5 Djangoda email orqali excel faylni jo’natish.</b></summary>
<br>

<ul>
    <li>Elektron pochtamizni xabar jo’natish uchun sozlab olamiz.</li>
    <li>Elektron pochta orqali xabar jo’natib testdan o’tkazamiz.</li>
    <li>Tayyor excel faylimizni bytes turiga o’girib olamiz.</li>
    <li>Django orqali email jo’natish uchun kerakli ma’lumotlarni sozlamalarda ko’rsatamiz.</li>
    <li>Excel faylni jo’natishi uchun SmtpServer classini yaratamiz.</li>
</ul>

To'liq video qo'llanma bu yerda: https://www.youtube.com/watch?v=X65PSzuTO6w

</details>

## Installation
```bash
python -m venv venv

source venv/bin/activate

pip install -r requirements.txt

python manage.py makemigrations

python manage.py migrate
```
## Run server
```bash
python manage.py createsuperuser

python manage.py runserver
```
## Author
[Sirojiddin Yakubov](https://t.me/Sirojiddin_Yakubov)