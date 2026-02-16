from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from src.excel_processor import iter_excel_openpyxl


class ReportTab:
    def __init__(self, master, app):
        self.app = app
        self.tab = master.add(f"{'Отчёт':^30}")
        # Переменные приложения
        self.excel_path = None
        self.xml_path = Path(__file__).parent / "static" / "data" / "data.xml"
        self.excel_processor = None
        self.xml_parser = None
        self.osg = "70"
        self.text = ""
        self.create_widgets_report()

    def create_widgets_report(self):
        """Создание всех виджетов интерфейса"""
        # Заголовок (строка 0, занимает все колонки)
        title_label = ctk.CTkLabel(
            self.tab, text="Генератор отчетов из Excel", font=("Arial", 24, "bold")
        )

        title_label.pack(pady=10, padx=20, fill="x")

        # Фрейм для выбора файла (строка 1)
        file_frame = ctk.CTkFrame(self.tab)
        file_frame.pack(pady=10, padx=20, fill="x")

        # Кнопка выбора Excel файла
        self.select_btn = ctk.CTkButton(
            file_frame,
            text="📁 Выбрать Excel файл",
            command=self.select_excel_file,
            height=40,
            width=200,
        )
        self.select_btn.pack(pady=10)

        # Метка с выбранным файлом
        self.file_label = ctk.CTkLabel(
            file_frame, text="Файл не выбран", font=("Arial", 12)
        )
        self.file_label.pack(pady=5)

        # Фрейм для кнопок действий (строка 2)
        action_frame = ctk.CTkFrame(self.tab)
        # action_frame.grid(row=2, column=0, columnspan=4, padx=20, pady=10, sticky="nsew")
        action_frame.pack(pady=10, padx=20, fill="x")

        # Кнопка проверки
        self.check_btn = ctk.CTkButton(
            action_frame,
            text="✅ Сформировать отчёт",
            command=self.check_xlsx_with_xml,
            state="disabled",
            height=40,
        )

        self.check_btn.pack(side="left", padx=10, pady=10, expand=True)

        self.osg_label = ctk.CTkLabel(action_frame, text="ОСГ (%)", font=("Arial", 12))
        self.osg_label.pack(side="left", padx=10, pady=10, expand=True)

        self.osg_entry = ctk.CTkEntry(action_frame)
        self.osg_entry.insert(0, self.osg)
        self.osg_entry.pack(side="left", padx=10, pady=10, expand=True)

        # Текстовое поле для отчета (строка 3)
        self.text_frame = ctk.CTkFrame(self.tab)
        # self.text_frame.grid(row=3, column=0, columnspan=4, padx=20, pady=10, sticky="nsew")
        self.text_frame.pack(pady=10, padx=20, fill="both", expand=True)

        # Заголовок текстового поля
        text_label = ctk.CTkLabel(
            self.text_frame,
            text="Текст отчета для мессенджера:",
            font=("Arial", 14, "bold"),
        )
        text_label.pack(pady=5)

        # Текстовое поле с прокруткой
        self.text_box = ctk.CTkTextbox(self.text_frame, font=("Arial", 12), wrap="word")
        self.text_box.pack(pady=10, padx=10, fill="both", expand=True)

        # Кнопка копирования
        self.copy_btn = ctk.CTkButton(
            self.text_frame,
            text="📋 Копировать в буфер",
            command=self.copy_to_clipboard,
            # state="disabled"
        )
        self.copy_btn.pack(pady=5)

    # Остальные методы остаются без изменений...
    def select_excel_file(self):
        """Выбор Excel файла"""
        file_path = filedialog.askopenfilename(
            title="Выберите Excel файл", filetypes=[("Excel files", "*.xlsx")]
        )

        if file_path:
            self.excel_path = Path(file_path)
            self.file_label.configure(text=f"Выбран: {Path(file_path).name}")
            self.check_btn.configure(state="normal")
            messagebox.showinfo("Успех", "Файл успешно выбран!")

    def check_xlsx_with_xml(self):
        """Проверка данных с XML"""
        try:
            self.osg = int(self.osg_entry.get())
            path_and_text_report = iter_excel_openpyxl(self.excel_path, int(self.osg))
            messagebox.showinfo("Инфо", f"Файл сохранен!\n{path_and_text_report[0]}")
            self.generate_report(path_and_text_report[1])

        except:
            messagebox.showerror("Ошибка", "Ошибка чтения файла")

    def save_edited_excel(self):
        """Сохранение отредактированного Excel файла"""
        messagebox.showinfo("Инфо", "Файл сохранен!")

    def generate_report(self, report_text):
        """Генерация текста отчета"""
        self.text_box.delete("1.0", "end")
        self.text_box.insert("1.0", f"СЕРЫШЕВО!\n{report_text}")
        self.copy_btn.configure(state="normal")
        messagebox.showinfo("Успех", "Отчет сгенерирован!")

    def copy_to_clipboard(self):
        """Копирование текста в буфер обмена"""
        text = self.text_box.get("1.0", "end-1c")
        self.app.clipboard_clear()
        self.app.clipboard_append(text)
        messagebox.showinfo("Успех", "Текст скопирован в буфер обмена!")
