import os
import tempfile
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import numpy as np
import matplotlib.pyplot as plt

# Бібліотеки для Excel та Word
import openpyxl
from openpyxl.chart import LineChart, Reference

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

MONTHS = [
    "Січень", "Лютий", "Березень", "Квітень", 
    "Травень", "Червень", "Липень", "Серпень", 
    "Вересень", "Жовтень", "Листопад", "Грудень"
]

class StaffMovementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Облік руху кадрів та некомплекту")
        self.root.geometry("680x750")
        self.root.minsize(650, 550)

        main_frame = ttk.Frame(root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Заголовок програми
        title_label = ttk.Label(
            main_frame, 
            text="Генератор графіків руху кадрів (12 місяців)", 
            font=("Helvetica", 13, "bold")
        )
        title_label.pack(pady=(0, 10))

        # Поле для назви графіку/звіту
        title_frame = ttk.LabelFrame(main_frame, text="Назва графіку / звіту", padding="10")
        title_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.entry_title = ttk.Entry(title_frame, font=("Helvetica", 10))
        self.entry_title.insert(0, "")
        self.entry_title.pack(fill=tk.X, expand=True)

        # Фрейм для 12 місяців із прокруткою
        months_frame_outer = ttk.LabelFrame(main_frame, text="Введення даних по місяцях", padding="10")
        months_frame_outer.pack(fill=tk.BOTH, expand=True, pady=5)

        canvas = tk.Canvas(months_frame_outer, height=240)
        scrollbar = ttk.Scrollbar(months_frame_outer, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Шапка таблиці місяців
        ttk.Label(self.scrollable_frame, text="Місяць", font=("Helvetica", 9, "bold")).grid(row=0, column=0, padx=10, pady=2, sticky=tk.W)
        ttk.Label(self.scrollable_frame, text="Приток кадрів (осіб)", font=("Helvetica", 9, "bold")).grid(row=0, column=1, padx=15, pady=2)
        ttk.Label(self.scrollable_frame, text="Відтік кадрів (осіб)", font=("Helvetica", 9, "bold")).grid(row=0, column=2, padx=15, pady=2)

        self.entries_inflow = []
        self.entries_outflow = []

        # Створення рядків для кожного місяця (дефолт = 0)
        for i, month in enumerate(MONTHS, start=1):
            ttk.Label(self.scrollable_frame, text=month, width=12).grid(row=i, column=0, padx=10, pady=4, sticky=tk.W)
            
            e_in = ttk.Entry(self.scrollable_frame, width=15, justify="center")
            e_in.insert(0, "0")
            e_in.grid(row=i, column=1, padx=15, pady=4)
            self.entries_inflow.append(e_in)

            e_out = ttk.Entry(self.scrollable_frame, width=15, justify="center")
            e_out.insert(0, "0")
            e_out.grid(row=i, column=2, padx=15, pady=4)
            self.entries_outflow.append(e_out)

        # Вибір формату та кнопка експорту
        export_frame = ttk.LabelFrame(main_frame, text="Експорт результату", padding="10")
        export_frame.pack(fill=tk.X, pady=10)

        self.export_format = tk.StringVar(value="excel")
        
        ttk.Radiobutton(
            export_frame, text="Excel (.xlsx)", variable=self.export_format, value="excel"
        ).pack(side=tk.LEFT, padx=15)

        ttk.Radiobutton(
            export_frame, text="Word (.docx)", variable=self.export_format, value="word"
        ).pack(side=tk.LEFT, padx=15)

        btn_generate = ttk.Button(
            export_frame, text="Сформувати файл", command=self.generate_report
        )
        btn_generate.pack(side=tk.RIGHT, padx=10)

    def get_data(self):
        data = []
        for i, month in enumerate(MONTHS):
            try:
                inflow = int(self.entries_inflow[i].get().strip() or 0)
            except ValueError:
                inflow = 0
            try:
                outflow = int(self.entries_outflow[i].get().strip() or 0)
            except ValueError:
                outflow = 0
            data.append((month, inflow, outflow))
        return data

    def generate_report(self):
        title = self.entry_title.get().strip()
        if not title:
            messagebox.showwarning("Помилка", "Введіть назву графіку/звіту!")
            return

        data = self.get_data()
        fmt = self.export_format.get()

        if fmt == "excel":
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")],
                title="Зберегти Excel файл"
            )
            if file_path:
                self.create_excel_report(title, data, file_path)
                messagebox.showinfo("Успіх", f"Excel графік успішно збережено:\n{file_path}")

        elif fmt == "word":
            file_path = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word Files", "*.docx")],
                title="Зберегти Word файл"
            )
            if file_path:
                self.create_word_report(title, data, file_path)
                messagebox.showinfo("Успіх", f"Word графік успішно збережено:\n{file_path}")

    def create_excel_report(self, title, data, path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Графік"

        # Технічні дані для побудови графіка (записуються в куток)
        ws["A1"] = "Місяць"
        ws["B1"] = "Приток кадрів"
        ws["C1"] = "Відтік кадрів"

        for row_idx, (month, inflow, outflow) in enumerate(data, start=2):
            ws.cell(row=row_idx, column=1, value=month)
            ws.cell(row=row_idx, column=2, value=inflow)
            ws.cell(row=row_idx, column=3, value=outflow)

        # Побудова лінійного графіка в Excel
        chart = LineChart()
        chart.title = title
        chart.style = 13
        chart.y_axis.title = "Кількість осіб"
        chart.x_axis.title = "Місяць"
        chart.width = 20
        chart.height = 11

        data_ref = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=len(data) + 1)
        cats_ref = Reference(ws, min_col=1, min_row=2, max_row=len(data) + 1)

        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)

        # Розміщуємо графік поверх стовпців (з комірки A1)
        ws.add_chart(chart, "E2")

        # Повністю приховуємо службові стовпці (A, B, C), щоб їх не було видно
        ws.column_dimensions['A'].hidden = True
        ws.column_dimensions['B'].hidden = True
        ws.column_dimensions['C'].hidden = True

        wb.save(path)

    def create_word_report(self, title, data, path):
        months = [item[0] for item in data]
        inflows = [item[1] for item in data]
        outflows = [item[2] for item in data]

        # Генерація графіка через Matplotlib
        x = np.arange(len(months))

        fig, ax = plt.subplots(figsize=(10, 5.5))
        
        # Стиль сірого фону та сіток
        ax.set_facecolor('#EAEAEA')
        ax.grid(True, which='both', color='white', linestyle='-', linewidth=1)
        fig.patch.set_facecolor('#F2F2F2')

        color_in = '#385723'   # Зелений (Приток)
        color_out = '#C00000'  # Червоний (Відтік)

        # Лінії графіка
        ax.plot(x, inflows, marker='o', linewidth=2.5, markersize=8, label='Приток кадрів', color=color_in)
        ax.plot(x, outflows, marker='o', linewidth=2.5, markersize=8, label='Відтік кадрів', color=color_out)

        # Цифри на точках
        for i, txt in enumerate(inflows):
            ax.annotate(str(txt), (x[i], inflows[i]),
                        textcoords="offset points", xytext=(0, 12), ha='center',
                        fontsize=9, fontweight='bold', color='white',
                        bbox=dict(boxstyle="circle,pad=0.3", fc=color_in, ec="none"))

        for i, txt in enumerate(outflows):
            ax.annotate(str(txt), (x[i], outflows[i]),
                        textcoords="offset points", xytext=(0, 12), ha='center',
                        fontsize=9, fontweight='bold', color='white',
                        bbox=dict(boxstyle="circle,pad=0.3", fc=color_out, ec="none"))

        ax.set_title(title, fontsize=13, fontweight='bold', pad=20, color='#333333')
        ax.set_xticks(x)
        ax.set_xticklabels(months, rotation=25, ha='right', fontsize=9, fontweight='bold')
        
        max_val = max(max(inflows), max(outflows)) if (inflows or outflows) else 10
        ax.set_ylim(-0.5, max_val + 2.5)

        ax.legend(loc='lower center', bbox_to_anchor=(0.5, -0.22), ncol=2, frameon=True, facecolor='white')
        plt.tight_layout()

        # Збереження у тимчасовий файл
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
            img_path = tmp.name

        plt.savefig(img_path, dpi=300, facecolor=fig.get_facecolor(), edgecolor='none')
        plt.close()

        # Створення Word документа (тільки назва та графік)
        doc = docx.Document()

        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p_title.add_run(title)
        run.font.size = Pt(15)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

        doc.add_paragraph() # Відступ
        
        # Вставка самого графіка як зображення на весь документ
        p_graph = doc.add_paragraph()
        p_graph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_graph.add_run().add_picture(img_path, width=Inches(6.5))

        doc.save(path)

        # Видалення тимчасового файлу зображення
        if os.path.exists(img_path):
            os.remove(img_path)

if __name__ == "__main__":
    root = tk.Tk()
    app = StaffMovementApp(root)
    root.mainloop()
