from calendar import monthrange
from datetime import datetime, timedelta
import os
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
import pandas as pd

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# Зберігаємо файл у системній папці "Документи" (Documents) на Windows
DOCUMENTS_PATH = os.path.join(os.path.expanduser("~"), "Documents")
DEFAULT_EXCEL_PATH = os.path.join(DOCUMENTS_PATH, "kontroly.xlsx")


class ControlApp(ctk.CTk):

  def __init__(self):
    super().__init__()

    self.title("Контролі")
    self.geometry("850x720")

    self.excel_path = DEFAULT_EXCEL_PATH

    # Верхня панель статусу файлу
    self.file_frame = ctk.CTkFrame(self)
    self.file_frame.pack(pady=10, padx=20, fill="x")

    self.btn_select_file = ctk.CTkButton(
        self.file_frame,
        text="Змінити шлях до файлу",
        command=self.select_file,
        fg_color="#4A5568",
    )
    self.btn_select_file.pack(side="left", padx=10, pady=10)

    self.lbl_file_status = ctk.CTkLabel(
        self.file_frame, text="", text_color="green", anchor="w"
    )
    self.lbl_file_status.pack(
        side="left", padx=10, pady=10, fill="x", expand=True
    )

    # Кнопки дій (Головне меню)
    self.actions_frame = ctk.CTkFrame(self)
    self.actions_frame.pack(pady=10, padx=20, fill="x")

    self.btn_add_window = ctk.CTkButton(
        self.actions_frame,
        text="➕ Додати контроль",
        command=self.open_add_control_window,
    )
    self.btn_add_window.pack(side="left", padx=10, pady=10)

    self.btn_check_controls = ctk.CTkButton(
        self.actions_frame,
        text="🔍 Перевірка контролів",
        command=self.open_check_window,
        fg_color="#2B7A78",
    )
    self.btn_check_controls.pack(side="left", padx=10, pady=10)

    # Окрема кнопка для швидкого відкриття та редагування Excel файлу
    self.btn_open_excel = ctk.CTkButton(
        self.actions_frame,
        text="📄 Відкрити Excel",
        command=self.open_excel_file,
        fg_color="#D97706",
    )
    self.btn_open_excel.pack(side="left", padx=10, pady=10)

    # Інформаційна область
    self.output_frame = ctk.CTkFrame(self)
    self.output_frame.pack(pady=10, padx=20, fill="both", expand=True)

    self.lbl_output_title = ctk.CTkLabel(
        self.output_frame,
        text=(
            "Головне меню. Файл 'kontroly.xlsx' завантажено з папки «Документи».\nНатисніть «Перевірка контролів», «Додати контроль» або «Відкрити Excel»."
        ),
        font=("Arial", 13),
        text_color="gray",
    )
    self.lbl_output_title.pack(anchor="w", padx=15, pady=15)

    self.init_default_file()

  def init_default_file(self):
    if not os.path.exists(DOCUMENTS_PATH):
      os.makedirs(DOCUMENTS_PATH, exist_ok=True)

    if not os.path.exists(self.excel_path):
      wb = openpyxl.Workbook()
      ws = wb.active
      columns = [
          "Назва контролю",
          "Орган куди направив контроль",
          "Класифікація",
          "Термін виконання",
          "Виконавець",
          "Виконано",
      ]
      for col_num, col_name in enumerate(columns, 1):
        ws.cell(row=1, column=col_num, value=col_name)
      wb.save(self.excel_path)

    filename = os.path.basename(self.excel_path)
    self.lbl_file_status.configure(
        text=f"Файл в роботі: {filename} (Документи)", text_color="green"
    )
    self._ensure_excel_structure()

  def select_file(self):
    filepath = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if filepath:
      self.excel_path = filepath
      filename = os.path.basename(filepath)
      self.lbl_file_status.configure(
          text=f"Обрано: {filename}", text_color="green"
      )
      self._ensure_excel_structure()

  def open_excel_file(self):
    try:
      if not os.path.exists(self.excel_path):
        self.init_default_file()
      
      if sys.platform == "win32":
        os.startfile(self.excel_path)
      else:
        subprocess.run(["open", self.excel_path])
    except Exception as e:
      messagebox.showerror("Помилка", f"Не вдалося відкрити файл: {str(e)}")

  def _ensure_excel_structure(self):
    columns = [
        "Назва контролю",
        "Орган куди направив контроль",
        "Класифікація",
        "Термін виконання",
        "Виконавець",
        "Виконано",
    ]
    try:
      wb = openpyxl.load_workbook(self.excel_path)
      ws = wb.active
      headers = [cell.value for cell in ws[1]]
      if not headers or all(h is None for h in headers):
        for col_num, col_name in enumerate(columns, 1):
          ws.cell(row=1, column=col_num, value=col_name)
        wb.save(self.excel_path)
      else:
        updated = False
        if "Виконавець" not in headers:
          if "Виконано" in headers:
            idx = headers.index("Виконано")
            headers.insert(idx, "Виконавець")
          else:
            headers.append("Виконавець")
          updated = True

        if "Виконано" not in headers:
          headers.append("Виконано")
          updated = True

        if updated:
          ws.delete_rows(1, ws.max_row)
          for col_num, col_name in enumerate(headers, 1):
            ws.cell(row=1, column=col_num, value=col_name)
          wb.save(self.excel_path)
    except Exception as e:
      print(f"Помилка структури Excel: {e}")

  def open_add_control_window(self):
    AddControlDialog(self, self.excel_path, self.apply_excel_formatting)

  def open_check_window(self):
    CheckControlsDialog(self, self.excel_path, self.apply_excel_formatting)

  def apply_excel_formatting(self):
    try:
      wb = openpyxl.load_workbook(self.excel_path)
      ws = wb.active

      data = []
      for row in ws.iter_rows(min_row=2, values_only=True):
        if any(row):
          data.append(list(row))

      if not data:
        return

      num_cols = len(ws[1])
      if num_cols >= 6:
        columns = [
            "Назва контролю",
            "Орган куди направив контроль",
            "Класифікація",
            "Термін виконання",
            "Виконавець",
            "Виконано",
        ]
      else:
        columns = [
            "Назва контролю",
            "Орган куди направив контроль",
            "Класифікація",
            "Термін виконання",
            "Виконано",
        ]

      df = pd.DataFrame(data, columns=columns[: len(data[0])])

      # Конвертуємо дати в об'єкти datetime для сортування
      df["TempDate"] = pd.to_datetime(
          df["Термін виконання"], format="%d.%m.%Y", errors="coerce"
      )
      df = df.sort_values(by="TempDate", ascending=True)

      ws.delete_rows(2, ws.max_row + 1)

      # Записуємо дані назад у таблицю як реальні дати Excel
      for r_idx, (_, row) in enumerate(df.iterrows(), start=2):
        row_values = row.drop(labels=["TempDate"]).values
        date_obj = row["TempDate"]
        
        for c_idx, val in enumerate(row_values, start=1):
          cell = ws.cell(row=r_idx, column=c_idx)
          if c_idx == 4 and not pd.isna(date_obj):
            cell.value = date_obj.date()
            cell.number_format = 'DD.MM.YYYY'
          else:
            cell.value = val

      thin_border = Border(
          left=Side(style="thin", color="CCCCCC"),
          right=Side(style="thin", color="CCCCCC"),
          top=Side(style="thin", color="CCCCCC"),
          bottom=Side(style="thin", color="CCCCCC"),
      )

      for row in ws.iter_rows(
          min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column
      ):
        for cell in row:
          cell.border = thin_border
          if cell.row == 1:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(
                horizontal="center", vertical="center"
            )
          else:
            if cell.column == ws.max_column:
              cell.alignment = Alignment(
                  horizontal="center", vertical="center"
              )
              cell.font = Font(name="Segoe UI Symbol", size=12, bold=True)
            else:
              cell.alignment = Alignment(vertical="center")

      for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
          if cell.value:
            max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 5, 15)

      # --- УМОВНЕ ФОРМАТУВАННЯ ДЛЯ КОЛОНОК A–E (Без F) ---
      ws.conditional_formatting.formulaRule = []

      max_r = max(ws.max_row, 2)
      target_range = f"A2:E{max_r}"

      # Кольори відповідно до вашого вибору
      overdue_fill = PatternFill(start_color="750377", end_color="750377", fill_type="solid") # Протерміновано
      red_fill = PatternFill(start_color="DC2626", end_color="DC2626", fill_type="solid")       # Сьогодні / Завтра
      orange_fill = PatternFill(start_color="D4652A", end_color="D4652A", fill_type="solid")    # Увага (2-4 дні)
      yellow_fill = PatternFill(start_color="E3F06F", end_color="E3F06F", fill_type="solid")    # Через 5-7 днів

      # 1. Протерміновано (< сьогодні) -> Заливка #750377
      ws.conditional_formatting.add(
          target_range,
          FormulaRule(formula=['AND($D2<>"", $D2<TODAY())'], stopIfTrue=True, fill=overdue_fill)
      )

      # 2. Червоний (Сьогодні / Завтра: від сьогодні до сьогодні + 1)
      ws.conditional_formatting.add(
          target_range,
          FormulaRule(formula=['AND($D2<>"", $D2>=TODAY(), $D2<=TODAY()+1)'], stopIfTrue=True, fill=red_fill)
      )

      # 3. Помаранчевий (Увага: 2-4 дні)
      ws.conditional_formatting.add(
          target_range,
          FormulaRule(formula=['AND($D2<>"", $D2>=TODAY()+2, $D2<=TODAY()+4)'], stopIfTrue=True, fill=orange_fill)
      )

      # 4. Жовтий (5-7 днів)
      ws.conditional_formatting.add(
          target_range,
          FormulaRule(formula=['AND($D2<>"", $D2>=TODAY()+5, $D2<=TODAY()+7)'], stopIfTrue=True, fill=yellow_fill)
      )

      wb.save(self.excel_path)
    except Exception as e:
      print(f"Помилка форматування Excel: {e}")


class CheckControlsDialog(ctk.CTkToplevel):

  def __init__(self, parent, excel_path, on_save_callback):
    super().__init__(parent)
    self.excel_path = excel_path
    self.on_save_callback = on_save_callback

    self.title("Перевірка контролів поточного місяця")
    self.geometry("820x580")
    self.grab_set()

    self.lbl_info = ctk.CTkLabel(
        self,
        text=(
            "Усі контролі поточного місяця. Позначте галочками виконані та"
            " збережіть:"
        ),
        font=("Arial", 13, "bold"),
    )
    self.lbl_info.pack(anchor="w", padx=20, pady=(15, 10))

    self.scroll_frame = ctk.CTkScrollableFrame(self, width=770, height=400)
    self.scroll_frame.pack(padx=20, pady=10, fill="both", expand=True)

    self.checkboxes = []
    self.load_controls()

    self.btn_save_done = ctk.CTkButton(
        self,
        text="💾 Зберегти виконання (прибрати виконані зі списку)",
        command=self.save_completed,
        fg_color="green",
        font=("Arial", 13, "bold"),
    )
    self.btn_save_done.pack(pady=15)

  def load_controls(self):
    for widget in self.scroll_frame.winfo_children():
      widget.destroy()
    self.checkboxes = []

    try:
      df = pd.read_excel(self.excel_path)
      if df.empty or "Термін виконання" not in df.columns:
        lbl = ctk.CTkLabel(
            self.scroll_frame, text="У файлі немає даних для перевірки."
        )
        lbl.pack(anchor="w", padx=10, pady=10)
        return

      today = datetime.now().date()
      current_month = today.month
      current_year = today.year
      found_any = False

      for index, row in df.iterrows():
        done_val = row.get("Виконано")
        if not pd.isna(done_val) and str(done_val).strip() in [
            "✔",
            "true",
            "1",
            "так",
            "yes",
            "v",
        ]:
          continue

        raw_date = row.get("Термін виконання")
        if pd.isna(raw_date):
          continue

        try:
          term_date = pd.to_datetime(
              raw_date, format="%d.%m.%Y", errors="coerce"
          ).date()
          if pd.isna(term_date):
            term_date = pd.to_datetime(raw_date, errors="coerce").date()
        except Exception:
          try:
            term_date = pd.to_datetime(raw_date).date()
          except:
            continue

        if pd.isna(term_date):
          continue

        if (
            term_date.month == current_month and term_date.year == current_year
        ) or term_date < today:
          found_any = True
          days_left = (term_date - today).days

          if days_left > 7:
            status_tag = "[ПЛАН]"
            text_color = "#718096"
          elif days_left >= 5:
            status_tag = "[ЧЕРЕЗ 5-7 ДНІВ]"
            text_color = "#E3F06F"
          elif days_left >= 2:
            status_tag = "[УВАГА]"
            text_color = "#D4652A"
          elif days_left == 1:
            status_tag = "[ЗАВТРА!]"
            text_color = "#DC2626"
          elif days_left == 0:
            status_tag = "[СЬОГОДНІ!]"
            text_color = "#DC2626"
          else:
            status_tag = "[ПРОТЕРМІНОВАНО!]"
            text_color = "#750377"

          control_name = row.get("Назва контролю", "Без назви")
          organ = row.get("Орган куди направив контроль", "-")
          classif = row.get("Класифікація", "-")
          executor = row.get("Виконавець", "-")
          date_str = term_date.strftime("%d.%m.%Y")

          text_label = (
              f"{status_tag} Термін: {date_str} (дн: {days_left}) | {control_name}"
              f" | Орган: {organ} | Тип: {classif} | Виконавець: {executor}"
          )

          var = tk.BooleanVar(value=False)
          chk = ctk.CTkCheckBox(
              self.scroll_frame,
              text=text_label,
              variable=var,
              font=("Arial", 12, "bold"),
              text_color=text_color
          )
          chk.pack(anchor="w", padx=10, pady=8)
          self.checkboxes.append({"index": index, "var": var})

      if not found_any:
        lbl = ctk.CTkLabel(
            self.scroll_frame,
            text="На поточний місяць активних невиконаних контролів немає!",
            text_color="gray",
        )
        lbl.pack(anchor="w", padx=10, pady=10)

    except Exception as e:
      messagebox.showerror("Помилка", f"Не вдалося зчитати файл: {str(e)}")

  def save_completed(self):
    try:
      wb = openpyxl.load_workbook(self.excel_path)
      ws = wb.active

      indices_to_done = [
          item["index"] for item in self.checkboxes if item["var"].get()
      ]

      if not indices_to_done:
        messagebox.showinfo(
            "Інформація",
            "Ви не позначили жодного контролю як виконаний.",
        )
        return

      header_row = [cell.value for cell in ws[1]]
      if "Виконано" in header_row:
        done_col_idx = header_row.index("Виконано") + 1
      else:
        done_col_idx = len(header_row) + 1
        ws.cell(row=1, column=done_col_idx, value="Виконано")

      for idx in indices_to_done:
        excel_row_num = idx + 2
        cell = ws.cell(row=excel_row_num, column=done_col_idx, value="✔")
        cell.font = Font(name="Segoe UI Symbol", size=12, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

      wb.save(self.excel_path)
      self.on_save_callback()

      messagebox.showinfo(
          "Успіх",
          (
              "Виконані контролі успішно збережено у вигляді галочок у Excel та"
              " приховано зі списку!"
          ),
      )
      self.load_controls()

    except Exception as e:
      messagebox.showerror(
          "Помилка", f"Не вдалося оновити статус у Excel: {str(e)}"
      )


class AddControlDialog(ctk.CTkToplevel):

  def __init__(self, parent, excel_path, on_save_callback):
    super().__init__(parent)
    self.excel_path = excel_path
    self.on_save_callback = on_save_callback

    self.title("Додати новий контроль")
    self.geometry("460x580")
    self.grab_set()

    ctk.CTkLabel(self, text="Назва контролю:").pack(
        anchor="w", padx=20, pady=(15, 0)
    )
    self.entry_name = ctk.CTkEntry(
        self, placeholder_text="Введіть назву контролю"
    )
    self.entry_name.pack(fill="x", padx=20, pady=(0, 10))

    ctk.CTkLabel(self, text="Орган куди направив контроль:").pack(
        anchor="w", padx=20, pady=(5, 0)
    )
    self.entry_organ = ctk.CTkEntry(
        self, placeholder_text="Вкажіть орган або підрозділ"
    )
    self.entry_organ.pack(fill="x", padx=20, pady=(0, 10))

    ctk.CTkLabel(self, text="Класифікація:").pack(
        anchor="w", padx=20, pady=(5, 0)
    )
    self.combo_class = ctk.CTkOptionMenu(
        self,
        values=["Разовий", "Щомісячний", "Квартальний", "Річний"],
        command=self.on_classification_change,
    )
    self.combo_class.pack(fill="x", padx=20, pady=(0, 10))

    self.lbl_date = ctk.CTkLabel(
        self, text="Термін виконання (ДД.ММ.РРРР):"
    )
    self.lbl_date.pack(anchor="w", padx=20, pady=(5, 0))
    self.entry_date = ctk.CTkEntry(
        self, placeholder_text="Наприклад, 15.11.2026"
    )
    self.entry_date.pack(fill="x", padx=20, pady=(0, 10))

    ctk.CTkLabel(self, text="Виконавець:").pack(
        anchor="w", padx=20, pady=(5, 0)
    )
    self.entry_executor = ctk.CTkEntry(
        self, placeholder_text="Введіть ПІБ або посаду виконавця"
    )
    self.entry_executor.pack(fill="x", padx=20, pady=(0, 15))

    self.next_year_var = tk.BooleanVar(value=False)
    self.chk_next_year = ctk.CTkCheckBox(
        self,
        text="Поширити на наступний рік (створити також на наступний рік)",
        variable=self.next_year_var,
        font=("Arial", 12),
    )
    self.chk_next_year.pack(anchor="w", padx=20, pady=(0, 15))

    self.btn_save = ctk.CTkButton(
        self,
        text="Зберегти та впорядкувати за датою",
        command=self.save_control,
        fg_color="green",
    )
    self.btn_save.pack(pady=5)

  def on_classification_change(self, choice):
    if choice == "Разовий":
      self.lbl_date.configure(text="Термін виконання (ДД.ММ.РРРР):")
      self.entry_date.configure(placeholder_text="Наприклад, 15.11.2026")
      self.chk_next_year.configure(state="disabled")
    else:
      self.chk_next_year.configure(state="normal")
      if choice == "Щомісячний":
        self.lbl_date.configure(
            text="Стартовий день/дата (програма створить контролі):"
        )
      elif choice == "Квартальний":
        self.lbl_date.configure(text="Дата початку (програма створить по кварталах):")
      elif choice == "Річний":
        self.lbl_date.configure(text="Дата початку (програма створить річні):")

  def save_control(self):
    name = self.entry_name.get().strip()
    organ = self.entry_organ.get().strip()
    classification = self.combo_class.get()
    date_str = self.entry_date.get().strip()
    executor = self.entry_executor.get().strip()
    include_next_year = self.next_year_var.get()

    if not name or not date_str:
      messagebox.showwarning(
          "Увага", "Заповніть назву контролю та поле дати!"
      )
      return

    try:
      start_date = datetime.strptime(date_str, "%d.%m.%Y")
    except ValueError:
      messagebox.showerror(
          "Помилка дати",
          "Введіть дату у правильному форматі: ДД.ММ.РРРР (наприклад, 15.11.2026)",
      )
      return

    try:
      df = pd.read_excel(self.excel_path)
      new_rows = []
      target_year = (
          start_date.year + 1 if include_next_year else start_date.year
      )
      end_date = datetime(target_year, 12, 31)
      original_day = start_date.day

      if classification == "Разовий":
        new_rows.append({
            "Назва контролю": name,
            "Орган куди направив контроль": organ,
            "Класифікація": classification,
            "Термін виконання": start_date.strftime("%d.%m.%Y"),
            "Виконавець": executor,
            "Виконано": "",
        })

      elif classification == "Щомісячний":
        curr = start_date
        while curr <= end_date:
          new_rows.append({
              "Назва контролю": name,
              "Орган куди направив контроль": organ,
              "Класифікація": classification,
              "Термін виконання": curr.strftime("%d.%m.%Y"),
              "Виконавець": executor,
              "Виконано": "",
          })
          if curr.month == 12 and curr.year >= target_year:
            break
          
          next_month = curr.month + 1
          year_val = curr.year
          if next_month > 12:
            next_month = 1
            year_val += 1
          
          last_day = monthrange(year_val, next_month)[1]
          day_val = min(original_day, last_day)
          curr = datetime(year_val, next_month, day_val)

      elif classification == "Квартальний":
        curr = start_date
        while curr <= end_date:
          new_rows.append({
              "Назва контролю": name,
              "Орган куди направив контроль": organ,
              "Класифікація": classification,
              "Термін виконання": curr.strftime("%d.%m.%Y"),
              "Виконавець": executor,
              "Виконано": "",
          })
          next_month = curr.month + 3
          year_val = curr.year
          if next_month > 12:
            next_month -= 12
            year_val += 1
            
          last_day = monthrange(year_val, next_month)[1]
          day_val = min(original_day, last_day)
          curr = datetime(year_val, next_month, day_val)

      elif classification == "Річний":
        curr = start_date
        while curr <= end_date:
          new_rows.append({
              "Назва контролю": name,
              "Орган куди направив контроль": organ,
              "Класифікація": classification,
              "Термін виконання": curr.strftime("%d.%m.%Y"),
              "Виконавець": executor,
              "Виконано": "",
          })
          year_val = curr.year + 1
          next_month = curr.month
          last_day = monthrange(year_val, next_month)[1]
          day_val = min(original_day, last_day)
          curr = datetime(year_val, next_month, day_val)

      df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)
      df.to_excel(self.excel_path, index=False)

      self.on_save_callback()

      messagebox.showinfo(
          "Успіх",
          (
              "Контроль(і) успішно додано, впорядковано за датою та оформлено у"
              " таблиці!"
          ),
      )
      self.destroy()

    except Exception as e:
      messagebox.showerror(
          "Помилка", f"Не вдалося зберегти дані у Excel: {str(e)}"
      )


if __name__ == "__main__":
  app = ControlApp()
  app.mainloop()
