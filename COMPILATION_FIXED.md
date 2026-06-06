# 🛠️ Compilation Troubleshooting Guide

## Переглянуто та виправлено всі проблеми компіляції

### ✅ Що було додано для успішної компіляції:

#### 1. **Оновлено requirements.txt**
Додано всі необхідні залежності:
- `pytz>=2023.3` - Часові зони для pandas
- `tzdata>=2023.3` - Дані часових зон
- `et-xmlfile>=1.1.0` - XML обробка для openpyxl
- `python-dateutil>=2.8.2` - Обробка дат

#### 2. **Покращено convert_csv_to_excel_v3.spec**
Додано:
- Повний список прихованих імпортів для pandas, openpyxl
- Внутрішні модулі pandas (_libs, _tslibs)
- Всі підмодулі openpyxl (styles, cell, worksheet, workbook)
- Кодування (utf-8, cp1251, cp1252)
- XML обробку (et_xmlfile, lxml)
- Файли даних (header_map.json, logo.ico)

#### 3. **Створено PyInstaller hooks**
Додано спеціальні хуки для автоматичного збору залежностей:
- `hooks/hook-pandas.py` - Для повного включення pandas
- `hooks/hook-openpyxl.py` - Для повного включення openpyxl

#### 4. **Додано build_and_verify.bat**
Автоматизований скрипт для:
- Перевірки середовища
- Встановлення залежностей
- Збірки executable
- Верифікації результату

---

## 🚀 Як компілювати

### Метод 1: Автоматична збірка (Рекомендовано)
```bash
build_and_verify.bat
```

### Метод 2: Ручна збірка
```bash
# 1. Встановити залежності
pip install -r requirements.txt

# 2. Зібрати executable
pyinstaller convert_csv_to_excel_v3.spec
```

---

## 🔍 Розв'язання типових проблем

### Проблема: "ModuleNotFoundError: No module named 'pandas._libs'"
**Рішення:** ✅ Виправлено додаванням hiddenimports у .spec файлі

### Проблема: "ModuleNotFoundError: No module named 'pytz'"
**Рішення:** ✅ Виправлено додаванням pytz у requirements.txt

### Проблема: "ImportError: cannot import name 'ElementTree'"
**Рішення:** ✅ Виправлено додаванням et_xmlfile та xml.etree.ElementTree

### Проблема: Програма не запускається без помилок
**Рішення:** ✅ Додано правильну обробку шляхів через get_app_directory()

### Проблема: "tkdnd not found"
**Рішення:** ✅ Виправлено правильним включенням tkdnd бібліотеки в datas

### Проблема: header_map.json не знайдено
**Рішення:** ✅ Додано header_map.json до datas у .spec файлі

---

## 📋 Checklist перед компіляцією

- [x] Python 3.8+ встановлено
- [x] requirements.txt оновлено
- [x] convert_csv_to_excel_v3.spec налаштовано
- [x] logo.ico присутній
- [x] header_map.json присутній
- [x] hooks/ директорія створена
- [x] PyInstaller hooks додано

---

## 📦 Результат компіляції

Після успішної компіляції:
- **Файл:** `dist/CSVtoExcel.exe`
- **Розмір:** ~60-100 MB (нормально для standalone app)
- **Включає:** Всі залежності, іконка, drag-and-drop підтримка

---

## 🎯 Тестування після компіляції

1. Запустіть `dist/CSVtoExcel.exe`
2. Перевірте:
   - ✅ Вікно відкривається
   - ✅ Іконка відображається
   - ✅ Drag-and-drop працює
   - ✅ Конвертація CSV файлів працює
   - ✅ Словник заголовків зберігається
   - ✅ Немає помилок консолі

---

## 💡 Додаткові поради

### Для зменшення розміру EXE:
- Використовується UPX compression (вже включено)
- Excluded: matplotlib, scipy, pytest (вже виключено)

### Для швидшої компіляції:
- Використовуйте `--clean` тільки при необхідності
- Кеш PyInstaller прискорює повторні збірки

### Для debug режиму:
В .spec файлі змініть:
```python
debug=True,
console=True,
```

---

## 📞 Підтримка

Якщо виникають проблеми:
1. Перевірте версію Python: `python --version`
2. Оновіть pip: `python -m pip install --upgrade pip`
3. Переустановіть залежності: `pip install -r requirements.txt --force-reinstall`
4. Очистіть кеш: видаліть папки `build/` та `dist/`
5. Запустіть `build_and_verify.bat`

---

**Версія документа:** 1.0  
**Дата оновлення:** 2026-06-06  
**Статус:** Всі проблеми компіляції виправлено ✅
