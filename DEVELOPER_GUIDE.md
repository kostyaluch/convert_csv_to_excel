# Developer Guide / Інструкція для розробників

## Підготовка середовища розробки

### Встановлення Python

Потрібен Python 3.8 або новіший. Завантажте з [python.org](https://www.python.org/downloads/).

### Встановлення залежностей

```bash
pip install -r requirements.txt
```

Або встановіть вручну:

```bash
pip install pandas openpyxl tkinterdnd2 pyinstaller
```

## Запуск з вихідного коду

```bash
python convert_csv_to_excel_v3.py
```

## Збірка .exe файлу

### Використання spec файлу

```bash
pyinstaller convert_csv_to_excel_v3.spec
```

Готовий виконуваний файл з'явиться у папці `dist/`.

### Збірка з нуля (без spec файлу)

```bash
pyinstaller --onefile --windowed --icon=logo.ico --name=convert_csv_to_excel_v3 convert_csv_to_excel_v3.py
```

Потім додайте до згенерованого `.spec` файлу підтримку tkinterdnd2:

```python
import sys
from pathlib import Path

# Get tkinterdnd2 package path
import tkinterdnd2
tkdnd_path = Path(tkinterdnd2.__file__).parent

a = Analysis(
    # ...
    datas=[(str(tkdnd_path / 'tkdnd'), 'tkdnd')],
    hiddenimports=['pandas', 'openpyxl', 'openpyxl.styles', 'openpyxl.utils', 
                   'html', 're', 'queue', 'threading', 'json', 'tkinterdnd2'],
    # ...
)
```

## Структура коду

### Основні модулі

- **Data Processing** - функції для очищення даних (`clean_cell_value`, `clean_text`)
- **Excel Operations** - збереження у Excel з форматуванням (`save_dataframe_to_excel`)
- **Column Pairing** - групування (ua) та (pl) стовпців
- **UI Components** - tkinter інтерфейс та drag-and-drop
- **Header Mapping** - словник для перейменування заголовків

### Потоки виконання

Програма використовує багатопотоковість:
- **Main Thread** - GUI інтерфейс (tkinter)
- **Worker Thread** - конвертація файлів
- **Queue** - комунікація між потоками

## Тестування

### Ручне тестування

1. **Drag-and-Drop:**
   - Запустіть програму
   - Перетягніть CSV файл у вікно
   - Перевірте, що файл додано до списку

2. **Command Line:**
   ```bash
   python convert_csv_to_excel_v3.py test.csv
   ```
   Програма має автоматично почати конвертацію

3. **Context Menu:**
   - Встановіть контекстне меню за допомогою `install_context_menu.bat`
   - Клікніть правою кнопкою на CSV файл
   - Виберіть "Конвертувати у Excel"

## Відладка

### Увімкнення консолі для .exe

У файлі `.spec` змініть:

```python
exe = EXE(
    # ...
    console=True,  # Було False
    # ...
)
```

Тепер при запуску .exe відкриється консоль з виводом помилок.

## Поширені проблеми

### tkinterdnd2 не працює у зібраному .exe

**Проблема:** Drag-and-drop не працює після збірки PyInstaller.

**Рішення:** Переконайтеся, що в `.spec` файлі додано:
```python
datas=[(str(tkdnd_path / 'tkdnd'), 'tkdnd')]
```

### Помилка імпорту tkinterdnd2

**Проблема:** `ModuleNotFoundError: No module named 'tkinterdnd2'`

**Рішення:**
```bash
pip install tkinterdnd2
```

Для Windows може знадобитися:
```bash
pip install tkinterdnd2 --no-cache-dir
```

## Додавання нових функцій

### Додавання нового режиму очищення

1. Додайте функцію очищення в секцію Data helpers
2. Оновіть `clean_cell_value()` для виклику нової функції
3. Додайте опцію у GUI (якщо потрібно)

### Додавання нового формату файлів

1. Додайте обробку у `select_files()` та `on_drop()`
2. Створіть функцію для читання нового формату
3. Оновіть `_convert_worker()` для обробки

## Ліцензія та внесок

Проект відкритий для внесків. Пропозиції та pull requests вітаються!
