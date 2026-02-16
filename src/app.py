import customtkinter as ctk

from src.act_tab import ActsTab
from src.report_tab import ReportTab


class MainApp(ctk.CTk):  # ← Наследуем от ctk.CTk
    def __init__(self):
        super().__init__()  # ← Инициализируем родительский класс

        # Настройка окна
        self.title("Генератор отчетов из Excel")
        # self.geometry("500x600")

        # Настройка внешнего вида
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")


        self.create_nodebook()

    def create_nodebook(self):
        self.notebook = ctk.CTkTabview(master=self, anchor="w", width=470, height=620)
        self.notebook.pack(fill="both", expand=True, pady=10, padx=10)

        self.report_tab = ReportTab(self.notebook, self)
        self.acts_tab = ActsTab(self.notebook)
