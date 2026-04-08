import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
import re
import html

def clean_cell_value(value):
    if pd.isnull(value):
        return value
    if isinstance(value, (int, float)):
        return value
    txt = str(value)
    for _ in range(5):
        new_txt = html.unescape(txt)
        if new_txt == txt:
            break
        txt = new_txt
    txt = txt.replace("\\'", "'")
    for ch in ["\r\n", "\n", "\r", "\t", "\\", "\s", "\f", "\v"]:
        txt = txt.replace(ch, "")
    txt = re.sub(r',(?=[^\s])', '; ', txt)
    txt = re.sub(' +', ' ', txt)
    txt = txt.strip()
    return txt

LIGHT_GREEN = "#e5fbe2"
DARK_GREEN = "#56ad4c"
BUTTON_GREEN = "#8fdc8a"
FONT = ("Segoe UI", 11)
TITLE_FONT = ("Segoe UI", 14, "bold")

class CsvToExcelConverterApp:
    def __init__(self, master):
        self.master = master
        master.title("Конвертер CSV у Excel")
        master.geometry("700x530")
        master.configure(bg=LIGHT_GREEN)

        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TButton',
                        font=FONT,
                        background=BUTTON_GREEN,
                        foreground='black',
                        borderwidth=1,
                        focusthickness=3,
                        focuscolor=DARK_GREEN)
        style.map('TButton', background=[('active', DARK_GREEN)])
        style.configure('TLabel',
                        background=LIGHT_GREEN,
                        font=FONT)
        style.configure('TFrame', background=LIGHT_GREEN)
        style.configure('Vertical.TScrollbar', background=LIGHT_GREEN)

        main_frame = ttk.Frame(master, padding=(20, 10, 20, 10), style='TFrame')
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="Конвертер CSV у Excel", font=TITLE_FONT, background=LIGHT_GREEN, foreground=DARK_GREEN).pack(pady=(0, 8))

        self.select_button = ttk.Button(main_frame, text="Вибрати файли", command=self.select_files)
        self.select_button.pack(pady=7)

        self.files_listbox = tk.Listbox(main_frame, width=90, height=8, font=("Segoe UI", 10), bg="white", borderwidth=2, relief="groove", highlightcolor=DARK_GREEN, selectbackground=BUTTON_GREEN)
        self.files_listbox.pack(pady=7)

        # === Перемикач очищення/оригінальність ===
        self.clean_var = tk.StringVar(value="clean")
        mode_frame = ttk.Frame(main_frame)
        mode_frame.pack(pady=4)
        ttk.Label(mode_frame, text="Форматування:").pack(side=tk.LEFT, padx=(0,8))
        ttk.Radiobutton(mode_frame, text="Очищений", variable=self.clean_var, value="clean").pack(side=tk.LEFT)
        ttk.Radiobutton(mode_frame, text="Оригінальний", variable=self.clean_var, value="original").pack(side=tk.LEFT)

        self.convert_button = ttk.Button(main_frame, text="Конвертувати у Excel", command=self.convert_files)
        self.convert_button.pack(pady=10)

        ttk.Label(main_frame, text="Журнал процесу:", font=("Segoe UI", 12, "bold"), background=LIGHT_GREEN, foreground=DARK_GREEN).pack(anchor="w", pady=(18, 2), padx=4)

        self.log_text = scrolledtext.ScrolledText(main_frame, width=90, height=9, font=("Segoe UI", 10), bg="#f7fdf7", borderwidth=2, relief="groove")
        self.log_text.pack(pady=5)
        self.log_text.configure(state='disabled')

        self.selected_files = []

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Виберіть CSV-файли для конвертації",
            filetypes=[("CSV or any file", "*.*")]
        )
        if files:
            self.selected_files = files
            self.files_listbox.delete(0, tk.END)
            for f in self.selected_files:
                self.files_listbox.insert(tk.END, f)
            self.log("Файли вибрано. Готові до конвертації.")

    def convert_files(self):
        if not self.selected_files:
            messagebox.showwarning("Увага", "Будь ласка, виберіть файли для конвертації.")
            return
        success_count = 0
        error_count = 0
        for file_path in self.selected_files:
            try:
                self.log(f"Читаю файл: {file_path}")
                df = pd.read_csv(file_path, encoding='utf-8', sep=',')
                # --- Очищати чи ні ---
                if self.clean_var.get() == "clean":
                    df = df.applymap(clean_cell_value)
                excel_file_path = os.path.splitext(file_path)[0] + '.xlsx'
                df.to_excel(excel_file_path, index=False)

                # === Форматування заголовка ===
                wb = load_workbook(excel_file_path)
                ws = wb.active
                header_row = ws[1]
                for cell in header_row:
                    cell.font = Font(bold=True, color="000000")    # Чорний, жирний
                    cell.alignment = Alignment(
                        wrap_text=True,                            # Перенесення тексту
                        vertical="top",                            # Вирівнювання по верхньому краю
                        horizontal="left"                          # Вирівнювання по лівому краю
                    )
                ws.row_dimensions[1].height = 50
                # Опціонально: автоширина колонок
                for col in ws.columns:
                    max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
                    ws.column_dimensions[get_column_letter(col[0].column)].width = max(15, min(50, max_length * 1.2))
                wb.save(excel_file_path)
                # ===

                self.log(f"✅ Успішно конвертовано: {os.path.basename(excel_file_path)}")
                success_count += 1
            except Exception as e:
                self.log(f"❌ Помилка у файлі {os.path.basename(file_path)}: {e}")
                error_count += 1
        self.log(f"\nЗавершено! Успішно: {success_count}, з помилкою: {error_count}")
        messagebox.showinfo("Готово", f"Успішно конвертовано: {success_count}\nЗ помилкою: {error_count}")

    def log(self, message):
        self.log_text.configure(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state='disabled')

if __name__ == "__main__":
    root = tk.Tk()
    app = CsvToExcelConverterApp(root)
    root.mainloop()
