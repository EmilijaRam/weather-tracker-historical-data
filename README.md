# Weather Tracker - Historical and Forecast Data 📊🌤️

Оваа Python скрипта автоматски собира историски и прогностички временски податоци за избран град (во случајов: Скопје) од Open-Meteo API, ги снима во Excel и визуелизира температурни трендови.

## ✨ Функционалности
- Преземање на дневни податоци (максимална, минимална температура и врнежи) за последниве неколку години.
- Прогноза за наредните 16 дена.
- Снимање на податоците во Excel со различни sheets.
- Визуелизација со matplotlib.
- Споредба по датуми за различни години.
- Автоматско бојадисување на редови по година.

## 📦 Зависности
Пред да ја стартувате скриптата, инсталирајте ги потребните библиотеки:

```bash
pip install requests pandas matplotlib openpyxl
▶️ Како се користи?
Подеси ги параметрите city_name, latitude, longitude и start_date во скриптата.

Стартувај ја Python скриптата:

bash
Copy
Edit
python ime_na_skriptata.py
Ќе се генерира Excel фајл со име Weather_Skopje_...xlsx и ќе се појави график.

👩‍💻 Автор
Овој проект е изработен од Emilija Ramova
GitHub: EmilijaRam