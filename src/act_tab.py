from pathlib import Path
from tkinter import filedialog, messagebox, StringVar

import customtkinter as ctk


from src.settings.settings import load_data, settings, append_data
from src.create_excel_act import pars_invoice, create_act
from src.utils.utils import remake_file


class ActsTab:
    def __init__(self, master):
        self.excel_path_list = []
        self.invoices_data = []
        self.tab = master.add(f"{'Акты':^30}")

        self.position1 = "Руководитель СК"
        try:
            self.name1 = load_data(str(settings.NAMES1_FILE))[-1::-1]
            self.name1str = self.name1[-1]
        except IndexError:
            self.name1str = ""
            self.name1 = [self.name1str]
        self.name1var = StringVar(value=self.name1[-1])
        self.position2 = "Кладовщик"
        try:
            self.name2 = load_data(str(settings.NAMES2_FILE))[-1::-1]
            self.name2str = self.name2[-1]
        except IndexError:
            self.name2str = ""
            self.name2 = [self.name2str]
        self.name2var = StringVar(value=self.name2[-1])

        try:
            self.reason_for_write_off = load_data(str(settings.REASON_FILE))[-1]
        except IndexError:
            self.reason_for_write_off = ""
        self.create_widgets_act()

    def create_widgets_act(self):
        """Создание всех виджетов интерфейса"""
        # Заголовок (строка 0, занимает все колонки)
        title_label_act = ctk.CTkLabel(
            self.tab, text="Генератор актов", font=("Arial", 24, "bold")
        )
        title_label_act.pack(pady=10, padx=20, fill="x")

        file_frame = ctk.CTkFrame(self.tab)
        file_frame.pack(pady=10, padx=20, fill="x")

        # Кнопка выбора Excel файла
        self.select_btn = ctk.CTkButton(
            file_frame,
            text="📁 Выбрать Excel файлы",
            command=self._select_excel_files,
            height=40,
            width=200,
        )
        self.select_btn.pack(pady=10)

        # Метка с выбранным файлом
        self.file_label = ctk.CTkLabel(
            file_frame,
            text="Файлы не выбраны",
            font=("Arial", 12),
        )
        self.file_label.pack(pady=5)

        # Фрейм для полей настроек
        setting_frame = ctk.CTkFrame(self.tab)
        setting_frame.pack(pady=10, padx=20, fill="x")

        # Настройка растягивания при изменении размера окна

        setting_frame.grid_rowconfigure(0, weight=1)
        setting_frame.grid_rowconfigure(1, weight=1)
        setting_frame.grid_rowconfigure(2, weight=1)
        setting_frame.grid_rowconfigure(3, weight=1)

        setting_frame.grid_columnconfigure(0, weight=1)
        setting_frame.grid_columnconfigure(1, weight=1)

        # Поля настроек
        self.position1_label = ctk.CTkLabel(
            setting_frame, text="Должность", font=("Arial", 12)
        )
        self.position1_label.grid(row=0, column=0, padx=5, pady=(0, 0), sticky="ew")

        self.position1_entry = ctk.CTkEntry(
            setting_frame,
        )
        self.position1_entry.insert(0, self.position1)
        self.position1_entry.grid(row=1, column=0, padx=5, pady=(0, 15), sticky="ew")

        self.name1_label = ctk.CTkLabel(
            setting_frame, text="Ф.И.О.", font=("Arial", 12)
        )
        self.name1_label.grid(row=0, column=1, padx=5, pady=(0, 0), sticky="ew")
        self.name1_combo_box = ctk.CTkComboBox(setting_frame, values=self.name1)
        # self.name1_combo_box.insert(0, self.name1)

        self.name1_combo_box.grid(row=1, column=1, padx=5, pady=(0, 15), sticky="ew")

        self.position2_label = ctk.CTkLabel(
            setting_frame, text="Должность", font=("Arial", 12)
        )
        self.position2_label.grid(row=2, column=0, padx=5, pady=(0, 0), sticky="ew")

        self.position2_entry = ctk.CTkEntry(setting_frame)
        self.position2_entry.insert(0, self.position2)
        self.position2_entry.grid(row=3, column=0, padx=5, pady=(0, 15), sticky="ew")

        self.name2_label = ctk.CTkLabel(
            setting_frame, text="Ф.И.О.", font=("Arial", 12)
        )
        self.name2_label.grid(row=2, column=1, padx=5, pady=(0, 0), sticky="ew")

        self.name2_combo_box = ctk.CTkComboBox(setting_frame, values=self.name2)

        # self.name2_entry.insert(0, self.name2)
        self.name2_combo_box.grid(row=3, column=1, padx=5, pady=(0, 15), sticky="ew")

        self.reason_for_write_off_label = ctk.CTkLabel(
            setting_frame, text="Причина списания", font=("Arial", 12)
        )
        self.reason_for_write_off_label.grid(
            row=4, column=0, padx=5, pady=(15, 0), sticky="ew"
        )
        self.reason_for_write_off_entry = ctk.CTkEntry(setting_frame)
        self.reason_for_write_off_entry.insert(0, self.reason_for_write_off)
        self.reason_for_write_off_entry.grid(
            row=4, column=1, padx=5, pady=(15, 0), sticky="ew"
        )

        # Фрейм для кнопки действия (строка 2)
        action_frame = ctk.CTkFrame(self.tab)
        # action_frame.grid(row=2, column=0, columnspan=4, padx=20, pady=10, sticky="nsew")
        action_frame.pack(pady=10, padx=20, fill="x")
        self.check_btn = ctk.CTkButton(
            action_frame,
            text="✅ Сформировать акты",
            command=self.create_act,
            state="disabled",
            height=40,
        )
        self.check_btn.pack(pady=10, padx=20, fill="x")

    def _select_excel_files(self):
        """Выбор Excel файлов"""
        files_path = filedialog.askopenfilenames(
            title="Выберите Excel файлы", filetypes=[("Excel files", "*.xlsx")]
        )

        if files_path:
            text_label = ""
            self.excel_path_list.clear()
            for file in files_path:
                self.excel_path_list.append(Path(file))
                text_label += Path(file).name + "\n"

            self.file_label.configure(
                text=f"Выбрано файлов - {len(files_path)} шт.:\n{text_label}"
            )
            self.check_btn.configure(state="normal")
            messagebox.showinfo("Успех", "Файлы успешно выбраны!")

    def _get_invoice_data(self):

        for file_path in self.excel_path_list:
            remake_file(file_path)
            invoice_data = pars_invoice(Path(file_path))
            self.invoices_data.append(invoice_data)

    def _create_setting_data(self):
        self.position1 = self.position1_entry.get()
        self.name1str = self.name1_combo_box.get()
        self.position2 = self.position2_entry.get()
        self.name2str = self.name2_combo_box.get()
        self.reason_for_write_off = self.reason_for_write_off_entry.get()
        append_data(self.name1str, str(settings.NAMES1_FILE))
        append_data(self.name2str, str(settings.NAMES2_FILE))
        append_data(self.reason_for_write_off, str(settings.REASON_FILE))
        self.name1 = load_data(str(settings.NAMES1_FILE))
        self.name2 = load_data(str(settings.NAMES2_FILE))

        self.name1_combo_box.configure(values=self.name1[-1::-1])
        self.name2_combo_box.configure(
            values=self.name2[-1::-1]
        )  # отображает недавно введенные вверху
        # self.name2_entry.update()
        if not all(
            (
                self.position1,
                self.name1str,
                self.position2,
                self.name2str,
                self.reason_for_write_off,
            )
        ):
            messagebox.showerror("Ошибка", "Заполните все поля")
            raise ValueError

    def create_act(self):
        try:
            self._get_invoice_data()
        except Exception:
            messagebox.showerror("Ошибка", "Ошибка при анализе требования-накладной.")
            return
        try:
            self._create_setting_data()
        except ValueError:
            return
        try:
            for invoice in self.invoices_data:
                create_act(
                    invoice,
                    name1=self.name1str,
                    position1=self.position1,
                    name2=self.name2str,
                    position2=self.position2,
                    reason=self.reason_for_write_off,
                )
            messagebox.showinfo("Успех", "Акты успешно сформированы")
        except Exception:
            messagebox.showerror("Ошибка", "Ошибка при формировании акта")
