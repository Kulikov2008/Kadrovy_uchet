import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
from db import get_connection

class HRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Кадровый учёт предприятия")
        self.root.geometry("1150x700")

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.tab_emp = tk.Frame(self.notebook)
        self.tab_dept = tk.Frame(self.notebook)
        self.tab_order = tk.Frame(self.notebook)

        self.notebook.add(self.tab_emp, text="Сотрудники")
        self.notebook.add(self.tab_dept, text="Отделы")
        self.notebook.add(self.tab_order, text="Приказы")

        self._build_emp_tab()
        self._build_dept_tab()
        self._build_order_tab()

        self.load_all_data()

    def _build_emp_tab(self):
        # Верхняя панель кнопок
        bar = tk.Frame(self.tab_emp, padx=5, pady=5)
        bar.pack(fill="x")
        tk.Button(bar, text="Обновить", command=self.load_all_data).pack(side="left", padx=2)
        tk.Button(bar, text="Добавить", command=self.add_employee).pack(side="left", padx=2)
        tk.Button(bar, text="Редактировать", command=self.edit_employee, bg="#e2e3e5").pack(side="left", padx=2)
        tk.Button(bar, text="Удалить", command=self.del_employee).pack(side="left", padx=2)

        # Кнопки экспорта/импорта
        tk.Button(bar, text="Экспорт в DOCX", command=self.export_docx, bg="#cce5ff").pack(side="left", padx=(20, 2))
        tk.Button(bar, text="Экспорт в XLSX", command=self.export_xlsx, bg="#d4edda").pack(side="left", padx=2)
        tk.Button(bar, text="Экспорт в PDF", command=self.export_pdf, bg="#f8d7da").pack(side="left", padx=2)
        tk.Button(bar, text="Импорт из CSV", command=self.import_csv, bg="#fff3cd").pack(side="left", padx=(20, 2))

        # НОВАЯ КНОПКА: Форма отчетов
        tk.Button(bar, text="Формирование отчетов", command=self.show_report_form, bg="#e2e3e5").pack(side="left", padx=(20, 2))

        # Панель поиска и фильтрации
        filter_bar = tk.Frame(self.tab_emp, padx=5, pady=3)
        filter_bar.pack(fill="x")

        tk.Label(filter_bar, text="Поиск (ФИО / должность):").pack(side="left", padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *args: self.load_employees())
        ent_search = tk.Entry(filter_bar, textvariable=self.search_var, width=30)
        ent_search.pack(side="left", padx=(0, 15))

        tk.Label(filter_bar, text="Фильтр по отделу:").pack(side="left", padx=(0, 5))
        self.filter_dept_var = tk.StringVar(value="Все отделы")
        self.combo_filter_dept = ttk.Combobox(filter_bar, textvariable=self.filter_dept_var, state="readonly", width=22)
        self.combo_filter_dept.pack(side="left", padx=(0, 10))
        self.combo_filter_dept.bind("<<ComboboxSelected>>", lambda e: self.load_employees())

        tk.Button(filter_bar, text="Сбросить", command=self.reset_filters).pack(side="left")

        # Таблица сотрудников
        cols = ("id", "full_name", "position", "dept_name", "salary")
        self.tree_emp = ttk.Treeview(self.tab_emp, columns=cols, show="headings")
        for c, t, w in [("id", "ID", 40), ("full_name", "ФИО", 250),
                        ("position", "Должность", 200), ("dept_name", "Отдел", 150),
                        ("salary", "Оклад", 100)]:
            self.tree_emp.heading(c, text=t)
            self.tree_emp.column(c, width=w)
        self.tree_emp.pack(fill="both", expand=True, padx=5, pady=5)
        self.tree_emp.bind("<Double-1>", lambda e: self.edit_employee())

    def _build_dept_tab(self):
        bar = tk.Frame(self.tab_dept, padx=5, pady=5)
        bar.pack(fill="x")
        tk.Button(bar, text="Обновить", command=self.load_all_data).pack(side="left", padx=2)
        tk.Button(bar, text="Добавить отдел", command=self.add_dept).pack(side="left", padx=2)
        tk.Button(bar, text="Удалить отдел", command=self.del_dept).pack(side="left", padx=2)

        cols = ("id", "name", "description")
        self.tree_dept = ttk.Treeview(self.tab_dept, columns=cols, show="headings")
        for c, t, w in [("id", "ID", 50), ("name", "Название", 200), ("description", "Описание", 400)]:
            self.tree_dept.heading(c, text=t)
            self.tree_dept.column(c, width=w)
        self.tree_dept.pack(fill="both", expand=True, padx=5, pady=5)

    def _build_order_tab(self):
        bar = tk.Frame(self.tab_order, padx=5, pady=5)
        bar.pack(fill="x")
        tk.Button(bar, text="Обновить", command=self.load_all_data).pack(side="left", padx=2)
        tk.Button(bar, text="Добавить приказ", command=self.add_order).pack(side="left", padx=2)
        tk.Button(bar, text="Удалить приказ", command=self.del_order).pack(side="left", padx=2)

        cols = ("id", "emp_name", "order_type", "details", "order_date")
        self.tree_order = ttk.Treeview(self.tab_order, columns=cols, show="headings")
        for c, t, w in [("id", "ID", 40), ("emp_name", "Сотрудник", 250),
                        ("order_type", "Тип", 150), ("details", "Детали", 250),
                        ("order_date", "Дата", 100)]:
            self.tree_order.heading(c, text=t)
            self.tree_order.column(c, width=w)
        self.tree_order.pack(fill="both", expand=True, padx=5, pady=5)

    def reset_filters(self):
        self.search_var.set("")
        self.filter_dept_var.set("Все отделы")
        self.load_employees()

    def load_employees(self):
        for i in self.tree_emp.get_children():
            self.tree_emp.delete(i)
        search_text = f"%{self.search_var.get().strip()}%"
        selected_dept = self.filter_dept_var.get()

        try:
            with get_connection() as conn:
                with conn.cursor() as cur:
                    query = """
                        SELECT e.id, e.full_name, e.position, d.name as dept_name, e.salary 
                        FROM employees e 
                        LEFT JOIN departments d ON e.department_id = d.id
                        WHERE (e.full_name ILIKE %s OR e.position ILIKE %s)
                    """
                    params = [search_text, search_text]

                    if selected_dept and selected_dept != "Все отделы":
                        query += " AND d.name = %s"
                        params.append(selected_dept)

                    query += " ORDER BY e.id"
                    cur.execute(query, tuple(params))

                    for row in cur.fetchall():
                        self.tree_emp.insert("", "end", values=(
                            row['id'], row['full_name'], row['position'],
                            row['dept_name'] or "—", row['salary']
                        ))
        except Exception as e:
            messagebox.showerror("Ошибка БД", str(e))

    def load_all_data(self):
        for t in [self.tree_dept, self.tree_order]:
            for i in t.get_children():
                t.delete(i)
        try:
            with get_connection() as conn:
                with conn.cursor() as cur:
                    cur.execute("SELECT * FROM departments ORDER BY id")
                    dept_names = ["Все отделы"]
                    for row in cur.fetchall():
                        self.tree_dept.insert("", "end", values=(row['id'], row['name'], row['description'] or ""))
                        dept_names.append(row['name'])

                    self.combo_filter_dept['values'] = dept_names
                    if self.filter_dept_var.get() not in dept_names:
                        self.filter_dept_var.set("Все отделы")

                    cur.execute("""
                        SELECT o.id, e.full_name as emp_name, o.order_type, o.details, o.order_date 
                        FROM orders o 
                        LEFT JOIN employees e ON o.employee_id = e.id ORDER BY o.id
                    """)
                    for row in cur.fetchall():
                        self.tree_order.insert("", "end", values=(
                            row['id'], row['emp_name'] or "—", row['order_type'],
                            row['details'] or "", row['order_date']
                        ))
            self.load_employees()
        except Exception as e:
            messagebox.showerror("Ошибка БД", str(e))

    # --- ФОРМЫ ДОБАВЛЕНИЯ И РЕДАКТИРОВАНИЯ ---
    def add_employee(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Новый сотрудник")
        dlg.geometry("320x320")
        dlg.grab_set()

        tk.Label(dlg, text="ФИО:").pack(pady=(10, 0))
        ent_fio = tk.Entry(dlg, width=32)
        ent_fio.pack()

        tk.Label(dlg, text="Должность:").pack(pady=(8, 0))
        ent_pos = tk.Entry(dlg, width=32)
        ent_pos.pack()

        tk.Label(dlg, text="Оклад:").pack(pady=(8, 0))
        ent_sal = tk.Entry(dlg, width=32)
        ent_sal.pack()

        tk.Label(dlg, text="Отдел:").pack(pady=(8, 0))
        combo_dept = ttk.Combobox(dlg, width=29, state="readonly")
        combo_dept.pack()

        depts = {}
        try:
            with get_connection() as conn:
                with conn.cursor() as cur:
                    cur.execute("SELECT id, name FROM departments ORDER BY id")
                    for r in cur.fetchall():
                        depts[r['name']] = r['id']
            combo_dept['values'] = list(depts.keys())
            if depts:
                combo_dept.current(0)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e), parent=dlg)

        def save():
            try:
                fio = ent_fio.get().strip()
                pos = ent_pos.get().strip()
                sal = float(ent_sal.get())
                dept_name = combo_dept.get()
                dept_id = depts.get(dept_name)

                if not fio or not pos:
                    raise ValueError("ФИО и должность обязательны!")

                with get_connection() as conn:
                    with conn.cursor() as cur:
                        cur.execute(
                            "INSERT INTO employees (full_name, position, salary, department_id) VALUES (%s, %s, %s, %s)",
                            (fio, pos, sal, dept_id)
                        )
                        conn.commit()
                dlg.destroy()
                self.load_all_data()
            except Exception as e:
                messagebox.showerror("Ошибка", str(e), parent=dlg)

        tk.Button(dlg, text="Сохранить", command=save, width=20, bg="#d4edda").pack(pady=20)

    def edit_employee(self):
        sel = self.tree_emp.selection()
        if not sel:
            messagebox.showwarning("Внимание", "Выберите сотрудника в таблице для редактирования!")
            return

        values = self.tree_emp.item(sel[0])["values"]
        emp_id, cur_fio, cur_pos, cur_dept, cur_sal = values[0], values[1], values[2], values[3], values[4]

        dlg = tk.Toplevel(self.root)
        dlg.title(f"Карточка сотрудника #{emp_id}")
        dlg.geometry("320x320")
        dlg.grab_set()

        tk.Label(dlg, text="ФИО:").pack(pady=(10, 0))
        ent_fio = tk.Entry(dlg, width=32)
        ent_fio.insert(0, str(cur_fio))
        ent_fio.pack()

        tk.Label(dlg, text="Должность:").pack(pady=(8, 0))
        ent_pos = tk.Entry(dlg, width=32)
        ent_pos.insert(0, str(cur_pos))
        ent_pos.pack()

        tk.Label(dlg, text="Оклад:").pack(pady=(8, 0))
        ent_sal = tk.Entry(dlg, width=32)
        ent_sal.insert(0, str(cur_sal))
        ent_sal.pack()

        tk.Label(dlg, text="Отдел:").pack(pady=(8, 0))
        combo_dept = ttk.Combobox(dlg, width=29, state="readonly")
        combo_dept.pack()

        depts = {}
        try:
            with get_connection() as conn:
                with conn.cursor() as cur:
                    cur.execute("SELECT id, name FROM departments ORDER BY id")
                    for r in cur.fetchall():
                        depts[r['name']] = r['id']
            dept_list = list(depts.keys())
            combo_dept['values'] = dept_list
            if cur_dept in dept_list:
                combo_dept.set(cur_dept)
            elif dept_list:
                combo_dept.current(0)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e), parent=dlg)

        def update():
            try:
                fio = ent_fio.get().strip()
                pos = ent_pos.get().strip()
                sal = float(ent_sal.get())
                dept_name = combo_dept.get()
                dept_id = depts.get(dept_name)

                if not fio or not pos:
                    raise ValueError("ФИО и должность обязательны!")

                with get_connection() as conn:
                    with conn.cursor() as cur:
                        cur.execute(
                            """UPDATE employees 
                               SET full_name = %s, position = %s, salary = %s, department_id = %s 
                               WHERE id = %s""",
                            (fio, pos, sal, dept_id, emp_id)
                        )
                        conn.commit()
                dlg.destroy()
                self.load_all_data()
            except Exception as e:
                messagebox.showerror("Ошибка", str(e), parent=dlg)

        tk.Button(dlg, text="Сохранить изменения", command=update, width=22, bg="#cce5ff").pack(pady=20)

    def add_dept(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Новый отдел")
        dlg.geometry("300x200")
        dlg.grab_set()

        tk.Label(dlg, text="Название:").pack(pady=(10, 0))
        ent_name = tk.Entry(dlg, width=30)
        ent_name.pack()

        tk.Label(dlg, text="Описание:").pack()
        ent_desc = tk.Entry(dlg, width=30)
        ent_desc.pack()

        def save():
            try:
                name = ent_name.get().strip()
                desc = ent_desc.get().strip()
                if not name:
                    raise ValueError("Название обязательно!")

                with get_connection() as conn:
                    with conn.cursor() as cur:
                        cur.execute(
                            "INSERT INTO departments (name, description) VALUES (%s, %s)",
                            (name, desc)
                        )
                        conn.commit()
                dlg.destroy()
                self.load_all_data()
            except Exception as e:
                messagebox.showerror("Ошибка", str(e), parent=dlg)

        tk.Button(dlg, text="Сохранить", command=save, width=20).pack(pady=20)

    def add_order(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Новый приказ")
        dlg.geometry("300x300")
        dlg.grab_set()

        tk.Label(dlg, text="Тип приказа:").pack(pady=(10, 0))
        combo_type = ttk.Combobox(dlg, values=["Прием", "Перевод", "Увольнение"], state="readonly")
        combo_type.current(0)
        combo_type.pack()

        tk.Label(dlg, text="Сотрудник:").pack()
        combo_emp = ttk.Combobox(dlg, state="readonly")
        combo_emp.pack()

        tk.Label(dlg, text="Детали:").pack()
        ent_details = tk.Entry(dlg, width=30)
        ent_details.pack()

        emps = {}
        try:
            with get_connection() as conn:
                with conn.cursor() as cur:
                    cur.execute("SELECT id, full_name FROM employees")
                    for r in cur.fetchall():
                        emps[r['full_name']] = r['id']
            combo_emp['values'] = list(emps.keys())
            if emps:
                combo_emp.current(0)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e), parent=dlg)

        def save():
            try:
                o_type = combo_type.get()
                emp_name = combo_emp.get()
                details = ent_details.get().strip()
                emp_id = emps.get(emp_name)

                if not emp_id:
                    raise ValueError("Выберите сотрудника!")

                with get_connection() as conn:
                    with conn.cursor() as cur:
                        cur.execute(
                            "INSERT INTO orders (employee_id, order_type, details) VALUES (%s, %s, %s)",
                            (emp_id, o_type, details)
                        )
                        conn.commit()
                dlg.destroy()
                self.load_all_data()
            except Exception as e:
                messagebox.showerror("Ошибка", str(e), parent=dlg)

        tk.Button(dlg, text="Сохранить", command=save, width=20).pack(pady=20)

    # --- УДАЛЕНИЕ ---
    def del_employee(self):
        sel = self.tree_emp.selection()
        if sel:
            eid = self.tree_emp.item(sel[0])["values"][0]
            if messagebox.askyesno("Удалить", f"Уволить сотрудника #{eid}?"):
                with get_connection() as conn:
                    with conn.cursor() as cur:
                        cur.execute("DELETE FROM employees WHERE id = %s", (eid,))
                        conn.commit()
                self.load_all_data()

    def del_dept(self):
        sel = self.tree_dept.selection()
        if sel:
            did = self.tree_dept.item(sel[0])["values"][0]
            if messagebox.askyesno("Удалить", f"Удалить отдел #{did}?"):
                with get_connection() as conn:
                    with conn.cursor() as cur:
                        cur.execute("DELETE FROM departments WHERE id = %s", (did,))
                        conn.commit()
                self.load_all_data()

    def del_order(self):
        sel = self.tree_order.selection()
        if sel:
            oid = self.tree_order.item(sel[0])["values"][0]
            if messagebox.askyesno("Удалить", f"Удалить приказ #{oid}?"):
                with get_connection() as conn:
                    with conn.cursor() as cur:
                        cur.execute("DELETE FROM orders WHERE id = %s", (oid,))
                        conn.commit()
                self.load_all_data()

    # --- ЭКСПОРТ И ИМПОРТ ---
    def export_docx(self):
        from export import export_to_docx
        items = [self.tree_emp.item(i)["values"] for i in self.tree_emp.get_children()]
        if not items:
            messagebox.showwarning("Внимание", "Нет данных для экспорта")
            return
        path = export_to_docx(items)
        messagebox.showinfo("Успех", f"Сохранено:\n{path}")

    def export_xlsx(self):
        from export import export_to_xlsx
        items = [self.tree_emp.item(i)["values"] for i in self.tree_emp.get_children()]
        if not items:
            messagebox.showwarning("Внимание", "Нет данных для экспорта")
            return
        path = export_to_xlsx(items)
        messagebox.showinfo("Успех", f"Сохранено:\n{path}")

    def export_pdf(self):
        from export import export_to_pdf
        items = [self.tree_emp.item(i)["values"] for i in self.tree_emp.get_children()]
        if not items:
            messagebox.showwarning("Внимание", "Нет данных для экспорта")
            return
        path = export_to_pdf(items)
        messagebox.showinfo("Успех", f"Сохранено:\n{path}")

    def import_csv(self):
        filepath = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if not filepath:
            return
        try:
            with open(filepath, mode='r', encoding='utf-8-sig') as f:
                reader = csv.reader(f, delimiter=';')
                count = 0
                with get_connection() as conn:
                    with conn.cursor() as cur:
                        for row in reader:
                            if len(row) >= 2:
                                cur.execute(
                                    "INSERT INTO employees (full_name, position, salary) VALUES (%s, %s, %s)",
                                    (row[0], row[1], float(row[2]) if len(row) > 2 else 0)
                                )
                                count += 1
                        conn.commit()
            self.load_all_data()
            messagebox.showinfo("Успех", f"Импортировано записей: {count}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка импорта: {e}")

    # --- ФОРМА ОТЧЕТОВ ---
    def show_report_form(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Формирование отчетов")
        dlg.geometry("350x300")
        dlg.grab_set()

        tk.Label(dlg, text="Выберите тип отчета:", font=("Arial", 10, "bold")).pack(pady=(15, 5))

        report_type = tk.StringVar(value="docx")
        
        tk.Radiobutton(dlg, text="Штатное расписание (DOCX)", variable=report_type, value="docx").pack(anchor="w", padx=40, pady=2)
        tk.Radiobutton(dlg, text="Ведомость сотрудников (XLSX)", variable=report_type, value="xlsx").pack(anchor="w", padx=40, pady=2)
        tk.Radiobutton(dlg, text="Справка о доходах (PDF)", variable=report_type, value="pdf").pack(anchor="w", padx=40, pady=2)

        tk.Label(dlg, text="Период (для ведомости):").pack(pady=(15, 0))
        ent_period = tk.Entry(dlg, width=20)
        ent_period.insert(0, "Текущий месяц")
        ent_period.pack(pady=5)

        def generate():
            items = [self.tree_emp.item(i)["values"] for i in self.tree_emp.get_children()]
            if not items:
                messagebox.showwarning("Внимание", "Нет данных для отчета", parent=dlg)
                return

            r_type = report_type.get()
            try:
                if r_type == "docx":
                    from export import export_to_docx
                    path = export_to_docx(items)
                elif r_type == "xlsx":
                    from export import export_to_xlsx
                    path = export_to_xlsx(items)
                elif r_type == "pdf":
                    from export import export_to_pdf
                    path = export_to_pdf(items)
                
                messagebox.showinfo("Успех", f"Отчет сформирован и сохранен:\n{path}", parent=dlg)
                dlg.destroy()
            except Exception as e:
                messagebox.showerror("Ошибка", str(e), parent=dlg)

        tk.Button(dlg, text="Сформировать", command=generate, width=20, bg="#d4edda").pack(pady=20)

if __name__ == "__main__":
    root = tk.Tk()
    app = HRApp(root)
    root.mainloop()