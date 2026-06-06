# ⚡ Швидка Компіляція - Шпаргалка

## Одна команда для успішної збірки:

```bash
build_and_verify.bat
```

---

## Або покроково:

### 1️⃣ Встановити залежності
```bash
pip install -r requirements.txt
```

### 2️⃣ Зібрати EXE
```bash
pyinstaller convert_csv_to_excel_v3.spec
```

### 3️⃣ Знайти результат
```
dist/CSVtoExcel.exe ← Ваш готовий файл
```

---

## ✅ Що вже включено:

### В requirements.txt:
- ✅ pandas, openpyxl, tkinterdnd2
- ✅ pytz, tzdata (для pandas)
- ✅ et-xmlfile (для openpyxl)
- ✅ python-dateutil

### В .spec файлі:
- ✅ 50+ hidden imports
- ✅ pandas._libs підмодулі
- ✅ openpyxl повний набір
- ✅ tkdnd бібліотека
- ✅ header_map.json
- ✅ logo.ico

### Додаткові файли:
- ✅ hooks/hook-pandas.py
- ✅ hooks/hook-openpyxl.py
- ✅ build_and_verify.bat

---

## 🔥 Типові помилки - ВИПРАВЛЕНО:

| Помилка | Статус |
|---------|--------|
| ModuleNotFoundError: pandas._libs | ✅ Виправлено |
| ModuleNotFoundError: pytz | ✅ Виправлено |
| ImportError: ElementTree | ✅ Виправлено |
| tkdnd not found | ✅ Виправлено |
| header_map.json missing | ✅ Виправлено |

---

## 📞 Якщо щось не працює:

```bash
# Очистити кеш
rmdir /s /q build dist

# Переустановити залежності
pip install -r requirements.txt --force-reinstall

# Зібрати знову
pyinstaller convert_csv_to_excel_v3.spec
```

---

## 📦 Розповсюдження:

**Мінімум:**
- dist/CSVtoExcel.exe

**Повний пакет:**
- dist/CSVtoExcel.exe
- USER_MANUAL.md
- install_context_menu.bat
- uninstall_context_menu.bat

---

**Все готово до компіляції! 🚀**
