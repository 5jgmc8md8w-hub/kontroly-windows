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

          # Використовуємо надійні текстові мітки та кольори для Windows
          if days_left > 7:
            status_tag = "[ПЛАН]"
            text_color = "gray"
          elif days_left >= 5:
            status_tag = "[ЧЕРЕЗ 5-7 ДНІВ]"
            text_color = "#D97706"  # Жовто-оранжевий
          elif days_left >= 2:
            status_tag = "[УВАГА]"
            text_color = "orange"
          elif days_left == 1:
            status_tag = "[ЗАВТРА!]"
            text_color = "red"
          elif days_left == 0:
            status_tag = "[СЬОГОДНІ!]"
            text_color = "red"
          else:
            status_tag = "[ПРОТЕРМІНОВАНО!]"
            text_color = "#991B1B"  # Темно-червоний

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
              font=("Arial", 12, "bold" if days_left <= 1 else "normal"),
              text_color=text_color if days_left <= 1 else None
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
