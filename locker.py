import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import sqlite3
from datetime import datetime
import os

# إعداد قاعدة البيانات
conn = sqlite3.connect("treasury.db")
c = conn.cursor()

# إنشاء الجداول (بدون عمود date أولاً)
c.execute('''CREATE TABLE IF NOT EXISTS income (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, batch REAL, date TEXT)''')
c.execute('''CREATE TABLE IF NOT EXISTS goods (id INTEGER PRIMARY KEY AUTOINCREMENT, item TEXT, cost REAL)''')
c.execute('''CREATE TABLE IF NOT EXISTS workers (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, salary REAL)''')
c.execute('''CREATE TABLE IF NOT EXISTS factory (id INTEGER PRIMARY KEY AUTOINCREMENT, description TEXT, cost REAL)''')
c.execute('''CREATE TABLE IF NOT EXISTS settings (id INTEGER PRIMARY KEY AUTOINCREMENT, password TEXT)''')

# التحقق من وجود عمود date وإضافته إذا لم يكن موجودًا
def add_date_column_if_not_exists(table_name):
    # جلب هيكل الجدول
    c.execute(f"PRAGMA table_info({table_name})")
    columns = [info[1] for info in c.fetchall()]
    if "date" not in columns:
        c.execute(f"ALTER TABLE {table_name} ADD COLUMN date TEXT")
        # تعبئة العمود الجديد بتاريخ افتراضي (مثلاً تاريخ اليوم) للسجلات الموجودة
        default_date = datetime.now().strftime("%Y-%m-%d")
        c.execute(f"UPDATE {table_name} SET date = ? WHERE date IS NULL", (default_date,))

# تطبيق التعديل على الجداول التي تحتاج إلى عمود date
add_date_column_if_not_exists("goods")
add_date_column_if_not_exists("workers")
add_date_column_if_not_exists("factory")

# إضافة الفهارس لتحسين أداء الاستعلامات
c.execute("CREATE INDEX IF NOT EXISTS idx_income_date ON income(date)")
c.execute("CREATE INDEX IF NOT EXISTS idx_income_name ON income(name)")
c.execute("CREATE INDEX IF NOT EXISTS idx_goods_item ON goods(item)")
c.execute("CREATE INDEX IF NOT EXISTS idx_workers_name ON workers(name)")
c.execute("CREATE INDEX IF NOT EXISTS idx_factory_description ON factory(description)")

# إعداد كلمة المرور الافتراضية
c.execute("SELECT password FROM settings WHERE id = 1")
if not c.fetchone():
    c.execute("INSERT INTO settings (password) VALUES (?)", ("123",))
conn.commit()

# جلب كلمة المرور
c.execute("SELECT password FROM settings WHERE id = 1")
PASSWORD = c.fetchone()[0]

# إعداد النافذة الرئيسية للبرنامج
root = tk.Tk()
root.title("نظام إدارة الخزنة")
root.geometry("1200x700")
root.configure(bg="#F3F4F6")

# تحسين التصميم باستخدام Style (لتحديد أنماط الأزرار والجداول)
style = ttk.Style()
style.theme_use("clam")
# تحديد نمط الجدول (Treeview) مع زيادة حجم الخط
style.configure("Treeview", font=("Arial", 12), rowheight=35, foreground="black", background="#FFFFFF")
# تحديد نمط رأس الجدول مع زيادة حجم الخط
style.configure("Treeview.Heading", font=("Arial", 14, "bold"), foreground="white", background="#1E3A8A")
style.map("Treeview", background=[("selected", "#A3BFFA")])
# تحديد نمط الأزرار
style.configure("TButton", padding=8, font=("Arial", 12), foreground="white", background="#1E3A8A")
style.map("TButton", background=[("active", "#1E40AF"), ("hover", "#3B82F6")])
# تحديد نمط النصوص
style.configure("TLabel", font=("Arial", 12), foreground="black", background="#F3F4F6")
style.configure("TNotebook", background="#F3F4F6")
style.configure("TNotebook.Tab", font=("Arial", 12), padding=[10, 5], foreground="black")

# إعداد الشريط العلوي
header_frame = tk.Frame(root, bg="#1E3A8A", height=60)
header_frame.pack(fill="x")

# أزرار الشريط العلوي (الجانب الأيسر)
left_button_frame = ttk.Frame(header_frame)
left_button_frame.pack(side="left", padx=10, pady=10)
ttk.Button(left_button_frame, text="تحديث", command=lambda: refresh_all()).pack(side="left", padx=5)

# عنوان البرنامج في الشريط العلوي
tk.Label(header_frame, text="Areez", font=("Arial", 20, "bold"), fg="white", bg="#1E3A8A").pack(side="left", expand=True)

# أزرار الشريط العلوي (الجانب الأيمن)
right_button_frame = ttk.Frame(header_frame)
right_button_frame.pack(side="right", padx=10, pady=10)
ttk.Button(right_button_frame, text="تغيير كلمة المرور", command=lambda: change_password_window()).pack(side="left", padx=5)
ttk.Button(right_button_frame, text="طباعة", command=lambda: print_section()).pack(side="left", padx=5)
ttk.Button(right_button_frame, text="تصدير إلى Excel", command=lambda: export_to_excel()).pack(side="left", padx=5)

# دالة للتحقق من كلمة المرور عند تسجيل الدخول
def check_password():
    def verify():
        if password_entry.get() == PASSWORD:
            login_window.destroy()
            root.deiconify()
        else:
            messagebox.showerror("خطأ", "كلمة المرور غير صحيحة")
    root.withdraw()
    login_window = tk.Toplevel(bg="#F3F4F6")
    login_window.title("تسجيل الدخول")
    login_window.geometry("350x200")
    tk.Label(login_window, text="تسجيل الدخول", font=("Arial", 14, "bold"), bg="#F3F4F6").pack(pady=10)
    tk.Label(login_window, text="كلمة المرور:", font=("Arial", 12), bg="#F3F4F6").pack(pady=5)
    password_entry = tk.Entry(login_window, show="*", font=("Arial", 12), width=25)
    password_entry.pack(pady=5)
    ttk.Button(login_window, text="دخول", command=verify).pack(pady=15)
    login_window.protocol("WM_DELETE_WINDOW", root.quit)

check_password()

# إنشاء علامات التبويب (Tabs) للأقسام المختلفة
notebook = ttk.Notebook(root)
notebook.pack(pady=10, padx=10, fill="both", expand=True)

# -------------------------------------- قسم الداخل --------------------------------------
# إعداد إطار قسم الداخل
income_frame = ttk.Frame(notebook, padding=10)
notebook.add(income_frame, text="الداخل")

# إعداد الجدول (Treeview) مع شريط تمرير
income_scroll = ttk.Scrollbar(income_frame, orient="vertical")
income_tree = ttk.Treeview(income_frame, columns=("ID", "Date", "Name", "Batch"), show="headings", style="Treeview", yscrollcommand=income_scroll.set)
income_scroll.config(command=income_tree.yview)
income_scroll.pack(side="right", fill="y")
income_tree.heading("ID", text="المعرف")
income_tree.heading("Date", text="التاريخ")
income_tree.heading("Name", text="الاسم")
income_tree.heading("Batch", text="الدفعة")
# ضبط عرض الأعمدة
income_tree.column("ID", width=100, anchor="center")
income_tree.column("Date", width=200, anchor="center")
income_tree.column("Name", width=300, anchor="center")
income_tree.column("Batch", width=200, anchor="center")
income_tree.pack(fill="both", expand=True, pady=10)

# إعداد الألوان المتناوبة للصفوف
income_tree.tag_configure("oddrow", background="#F9FAFB")
income_tree.tag_configure("evenrow", background="#FFFFFF")

# عرض الإجمالي أسفل الجدول
total_label_income = ttk.Label(income_frame, text="الإجمالي: 0", font=("Arial", 14, "bold"))
total_label_income.pack(pady=5)

# دالة لحساب إجمالي الدفعات في قسم الداخل
def update_total_income():
    total = sum(float(income_tree.item(item)["values"][3]) for item in income_tree.get_children() if income_tree.item(item)["values"][3])
    total_label_income.config(text=f"الإجمالي: {total:,.2f}")

# دالة لتحميل بيانات قسم الداخل من قاعدة البيانات
def load_income_data():
    for item in income_tree.get_children():
        income_tree.delete(item)
    c.execute("SELECT id, date, name, batch FROM income")
    rows = c.fetchall()
    for i, row in enumerate(rows):
        tag = "evenrow" if i % 2 == 0 else "oddrow"
        income_tree.insert("", "end", values=row, tags=(tag,))
    update_total_income()

load_income_data()

# إعداد إطار البحث في قسم الداخل
search_frame = ttk.Frame(income_frame)
search_frame.pack(pady=10, fill="x")
tk.Label(search_frame, text="البحث:", font=("Arial", 12)).pack(side="right", padx=5)
search_name_entry = ttk.Entry(search_frame, width=15, font=("Arial", 10))
search_name_entry.pack(side="right", padx=5)
tk.Label(search_frame, text="الاسم:").pack(side="right", padx=5)
search_date_entry = ttk.Entry(search_frame, width=15, font=("Arial", 10))
search_date_entry.pack(side="right", padx=5)
tk.Label(search_frame, text="التاريخ:").pack(side="right", padx=5)
search_batch_entry = ttk.Entry(search_frame, width=15, font=("Arial", 10))
search_batch_entry.pack(side="right", padx=5)
tk.Label(search_frame, text="الدفعة:").pack(side="right", padx=5)

# دالة للبحث في بيانات قسم الداخل
def search_income():
    name_term = search_name_entry.get().strip().lower()
    date_term = search_date_entry.get().strip()
    batch_term = search_batch_entry.get().strip()
    all_items = income_tree.get_children()
    for item in all_items:
        income_tree.detach(item)
    for item in all_items:
        values = income_tree.item(item)["values"]
        if (not name_term or name_term in str(values[2]).lower()) and \
           (not date_term or date_term in str(values[1])) and \
           (not batch_term or batch_term in str(values[3])):
            income_tree.reattach(item, "", "end")
    update_total_income()

ttk.Button(search_frame, text="بحث", command=search_income).pack(side="right", padx=5)

# دالة لإضافة بيانات جديدة في قسم الداخل
def add_income_data():
    add_window = tk.Toplevel(root, bg="#F3F4F6")
    add_window.title("إضافة بيانات الداخل")
    add_window.geometry("400x250")
    add_window.resizable(False, False)
    
    tk.Label(add_window, text="إضافة بيانات جديدة", font=("Arial", 16, "bold"), bg="#F3F4F6").pack(pady=15)
    frame = ttk.Frame(add_window)
    frame.pack(padx=20, pady=10)
    
    tk.Label(frame, text="الاسم:", font=("Arial", 12), bg="#F3F4F6").grid(row=0, column=0, padx=10, pady=10, sticky="e")
    name_entry = ttk.Entry(frame, width=25, font=("Arial", 12))
    name_entry.grid(row=0, column=1)
    tk.Label(frame, text="الدفعة:", font=("Arial", 12), bg="#F3F4F6").grid(row=1, column=0, padx=10, pady=10, sticky="e")
    batch_entry = ttk.Entry(frame, width=25, font=("Arial", 12))
    batch_entry.grid(row=1, column=1)
    
    def save_income():
        name, batch = name_entry.get().strip(), batch_entry.get().strip()
        if not name or not batch:
            messagebox.showwarning("خطأ", "يرجى ملء جميع الحقول")
            return
        try:
            batch = float(batch)
            if batch < 0:
                raise ValueError("الدفعة يجب أن تكون موجبة")
            date = datetime.now().strftime("%Y-%m-%d")  # إضافة التاريخ تلقائيًا
            with conn:
                c.execute("INSERT INTO income (name, batch, date) VALUES (?, ?, ?)", (name, batch, date))
            tag = "evenrow" if len(income_tree.get_children()) % 2 == 0 else "oddrow"
            income_tree.insert("", "end", values=(c.lastrowid, date, name, batch), tags=(tag,))
            update_total_income()
            add_window.destroy()
        except ValueError as e:
            messagebox.showwarning("خطأ", str(e) or "يرجى إدخال قيمة عددية صحيحة")

    ttk.Button(add_window, text="حفظ", command=save_income).pack(pady=20)

# دالة لتعديل بيانات موجودة في قسم الداخل
def edit_income_data():
    selected = income_tree.selection()
    if not selected:
        messagebox.showwarning("خطأ", "يرجى تحديد سجل لتعديله")
        return
    item_values = income_tree.item(selected[0])["values"]
    edit_window = tk.Toplevel(root, bg="#F3F4F6")
    edit_window.title("تعديل بيانات الداخل")
    edit_window.geometry("400x250")
    edit_window.resizable(False, False)
    
    tk.Label(edit_window, text="تعديل البيانات", font=("Arial", 16, "bold"), bg="#F3F4F6").pack(pady=15)
    frame = ttk.Frame(edit_window)
    frame.pack(padx=20, pady=10)
    
    tk.Label(frame, text="الاسم:", font=("Arial", 12), bg="#F3F4F6").grid(row=0, column=0, padx=10, pady=10, sticky="e")
    name_entry = ttk.Entry(frame, width=25, font=("Arial", 12))
    name_entry.insert(0, item_values[2])
    name_entry.grid(row=0, column=1)
    tk.Label(frame, text="الدفعة:", font=("Arial", 12), bg="#F3F4F6").grid(row=1, column=0, padx=10, pady=10, sticky="e")
    batch_entry = ttk.Entry(frame, width=25, font=("Arial", 12))
    batch_entry.insert(0, item_values[3])
    batch_entry.grid(row=1, column=1)
    
    def save_edit():
        name, batch = name_entry.get().strip(), batch_entry.get().strip()
        if not name or not batch:
            messagebox.showwarning("خطأ", "يرجى ملء جميع الحقول")
            return
        try:
            batch = float(batch)
            if batch < 0:
                raise ValueError("الدفعة يجب أن تكون موجبة")
            with conn:
                c.execute("UPDATE income SET name=?, batch=? WHERE id=?", (name, batch, item_values[0]))
            income_tree.item(selected[0], values=(item_values[0], item_values[1], name, batch))
            update_total_income()
            edit_window.destroy()
        except ValueError as e:
            messagebox.showwarning("خطأ", str(e) or "يرجى إدخال قيمة عددية صحيحة")

    ttk.Button(edit_window, text="حفظ التغييرات", command=save_edit).pack(pady=20)

# دالة لحذف بيانات من قسم الداخل
def delete_income_data():
    selected = income_tree.selection()
    if selected and messagebox.askyesno("تأكيد", "هل تريد حذف السجل المحدد؟"):
        values = income_tree.item(selected[0])["values"]
        with conn:
            c.execute("DELETE FROM income WHERE id=?", (values[0],))
        income_tree.delete(selected[0])
        update_total_income()

# أزرار الإضافة، التعديل، والحذف في قسم الداخل
button_frame_income = ttk.Frame(income_frame)
button_frame_income.pack(pady=10, fill="x")
ttk.Button(button_frame_income, text="إضافة", command=add_income_data).pack(side="left", padx=5)
ttk.Button(button_frame_income, text="تعديل", command=edit_income_data).pack(side="left", padx=5)
ttk.Button(button_frame_income, text="حذف", command=delete_income_data).pack(side="left", padx=5)

# -------------------------------------- قسم البضائع --------------------------------------
# إعداد إطار قسم البضائع
goods_frame = ttk.Frame(notebook, padding=10)
notebook.add(goods_frame, text="المصنع")

# إعداد الجدول (Treeview) مع شريط تمرير
goods_scroll = ttk.Scrollbar(goods_frame, orient="vertical")
goods_tree = ttk.Treeview(goods_frame, columns=("ID", "Date", "Cost", "Item"), show="headings", style="Treeview", yscrollcommand=goods_scroll.set)
goods_scroll.config(command=goods_tree.yview)
goods_scroll.pack(side="right", fill="y")
goods_tree.heading("ID", text="المعرف")
goods_tree.heading("Date", text="التاريخ")
goods_tree.heading("Cost", text="التكلفة")
goods_tree.heading("Item", text="الصنف")
# ضبط عرض الأعمدة
goods_tree.column("ID", width=100, anchor="center")
goods_tree.column("Date", width=200, anchor="center")
goods_tree.column("Cost", width=200, anchor="center")
goods_tree.column("Item", width=300, anchor="center")
goods_tree.pack(fill="both", expand=True, pady=10)

# إعداد الألوان المتناوبة للصفوف
goods_tree.tag_configure("oddrow", background="#F9FAFB")
goods_tree.tag_configure("evenrow", background="#FFFFFF")

# عرض الإجمالي أسفل الجدول
total_label_goods = ttk.Label(goods_frame, text="الإجمالي: 0", font=("Arial", 14, "bold"))
total_label_goods.pack(pady=5)

# دالة لحساب إجمالي التكاليف في قسم البضائع
def update_total_goods():
    total = sum(float(goods_tree.item(item)["values"][2]) for item in goods_tree.get_children() if goods_tree.item(item)["values"][2])
    total_label_goods.config(text=f"الإجمالي: {total:,.2f}")

# دالة لتحميل بيانات قسم البضائع من قاعدة البيانات
def load_goods_data():
    for item in goods_tree.get_children():
        goods_tree.delete(item)
    c.execute("SELECT id, date, cost, item FROM goods")
    rows = c.fetchall()
    for i, row in enumerate(rows):
        tag = "evenrow" if i % 2 == 0 else "oddrow"
        goods_tree.insert("", "end", values=row, tags=(tag,))
    update_total_goods()

load_goods_data()

# دالة لإضافة بيانات جديدة في قسم البضائع
def add_goods_data():
    add_window = tk.Toplevel(root, bg="#F3F4F6")
    add_window.title("إضافة بيانات البضائع")
    add_window.geometry("400x250")
    add_window.resizable(False, False)
    
    tk.Label(add_window, text="إضافة بيانات جديدة", font=("Arial", 16, "bold"), bg="#F3F4F6").pack(pady=15)
    frame = ttk.Frame(add_window)
    frame.pack(padx=20, pady=10)
    
    tk.Label(frame, text="الصنف:", font=("Arial", 12), bg="#F3F4F6").grid(row=0, column=0, padx=10, pady=10, sticky="e")
    item_entry = ttk.Entry(frame, width=25, font=("Arial", 12))
    item_entry.grid(row=0, column=1)
    tk.Label(frame, text="التكلفة:", font=("Arial", 12), bg="#F3F4F6").grid(row=1, column=0, padx=10, pady=10, sticky="e")
    cost_entry = ttk.Entry(frame, width=25, font=("Arial", 12))
    cost_entry.grid(row=1, column=1)
    
    def save_goods():
        item, cost = item_entry.get().strip(), cost_entry.get().strip()
        if not item or not cost:
            messagebox.showwarning("خطأ", "يرجى ملء جميع الحقول")
            return
        try:
            cost = float(cost)
            if cost < 0:
                raise ValueError("التكلفة يجب أن تكون موجبة")
            date = datetime.now().strftime("%Y-%m-%d")  # إضافة التاريخ تلقائيًا
            with conn:
                c.execute("INSERT INTO goods (item, cost, date) VALUES (?, ?, ?)", (item, cost, date))
            tag = "evenrow" if len(goods_tree.get_children()) % 2 == 0 else "oddrow"
            goods_tree.insert("", "end", values=(c.lastrowid, date, cost, item), tags=(tag,))
            update_total_goods()
            add_window.destroy()
        except ValueError as e:
            messagebox.showwarning("خطأ", str(e) or "يرجى إدخال قيمة عددية صحيحة")

    ttk.Button(add_window, text="حفظ", command=save_goods).pack(pady=20)

# دالة لتعديل بيانات موجودة في قسم البضائع
def edit_goods_data():
    selected = goods_tree.selection()
    if not selected:
        messagebox.showwarning("خطأ", "يرجى تحديد سجل لتعديله")
        return
    item_values = goods_tree.item(selected[0])["values"]
    edit_window = tk.Toplevel(root, bg="#F3F4F6")
    edit_window.title("تعديل بيانات البضائع")
    edit_window.geometry("400x250")
    edit_window.resizable(False, False)
    
    tk.Label(edit_window, text="تعديل البيانات", font=("Arial", 16, "bold"), bg="#F3F4F6").pack(pady=15)
    frame = ttk.Frame(edit_window)
    frame.pack(padx=20, pady=10)
    
    tk.Label(frame, text="الصنف:", font=("Arial", 12), bg="#F3F4F6").grid(row=0, column=0, padx=10, pady=10, sticky="e")
    item_entry = ttk.Entry(frame, width=25, font=("Arial", 12))
    item_entry.insert(0, item_values[3])
    item_entry.grid(row=0, column=1)
    tk.Label(frame, text="التكلفة:", font=("Arial", 12), bg="#F3F4F6").grid(row=1, column=0, padx=10, pady=10, sticky="e")
    cost_entry = ttk.Entry(frame, width=25, font=("Arial", 12))
    cost_entry.insert(0, item_values[2])
    cost_entry.grid(row=1, column=1)
    
    def save_edit():
        item, cost = item_entry.get().strip(), cost_entry.get().strip()
        if not item or not cost:
            messagebox.showwarning("خطأ", "يرجى ملء جميع الحقول")
            return
        try:
            cost = float(cost)
            if cost < 0:
                raise ValueError("التكلفة يجب أن تكون موجبة")
            with conn:
                c.execute("UPDATE goods SET item=?, cost=? WHERE id=?", (item, cost, item_values[0]))
            goods_tree.item(selected[0], values=(item_values[0], item_values[1], cost, item))
            update_total_goods()
            edit_window.destroy()
        except ValueError as e:
            messagebox.showwarning("خطأ", str(e) or "يرجى إدخال قيمة عددية صحيحة")

    ttk.Button(edit_window, text="حفظ التغييرات", command=save_edit).pack(pady=20)

# دالة لحذف بيانات من قسم البضائع
def delete_goods_data():
    selected = goods_tree.selection()
    if selected and messagebox.askyesno("تأكيد", "هل تريد حذف السجل المحدد؟"):
        values = goods_tree.item(selected[0])["values"]
        with conn:
            c.execute("DELETE FROM goods WHERE id=?", (values[0],))
        goods_tree.delete(selected[0])
        update_total_goods()

# أزرار الإضافة، التعديل، والحذف في قسم البضائع
button_frame_goods = ttk.Frame(goods_frame)
button_frame_goods.pack(pady=10, fill="x")
ttk.Button(button_frame_goods, text="إضافة", command=add_goods_data).pack(side="left", padx=5)
ttk.Button(button_frame_goods, text="تعديل", command=edit_goods_data).pack(side="left", padx=5)
ttk.Button(button_frame_goods, text="حذف", command=delete_goods_data).pack(side="left", padx=5)

# -------------------------------------- قسم العمال --------------------------------------
# إعداد إطار قسم العمال
workers_frame = ttk.Frame(notebook, padding=10)
notebook.add(workers_frame, text="العمال")

# إعداد الجدول (Treeview) مع شريط تمرير
workers_scroll = ttk.Scrollbar(workers_frame, orient="vertical")
workers_tree = ttk.Treeview(workers_frame, columns=("ID", "Date", "Salary", "Name"), show="headings", style="Treeview", yscrollcommand=workers_scroll.set)
workers_scroll.config(command=workers_tree.yview)
workers_scroll.pack(side="right", fill="y")
workers_tree.heading("ID", text="المعرف")
workers_tree.heading("Date", text="التاريخ")
workers_tree.heading("Salary", text="المرتب")
workers_tree.heading("Name", text="اسم الموظف")
# ضبط عرض الأعمدة
workers_tree.column("ID", width=100, anchor="center")
workers_tree.column("Date", width=200, anchor="center")
workers_tree.column("Salary", width=200, anchor="center")
workers_tree.column("Name", width=300, anchor="center")
workers_tree.pack(fill="both", expand=True, pady=10)

# إعداد الألوان المتناوبة للصفوف
workers_tree.tag_configure("oddrow", background="#F9FAFB")
workers_tree.tag_configure("evenrow", background="#FFFFFF")

# عرض الإجمالي أسفل الجدول
total_label_workers = ttk.Label(workers_frame, text="الإجمالي: 0", font=("Arial", 14, "bold"))
total_label_workers.pack(pady=5)

# دالة لحساب إجمالي المرتبات في قسم العمال
def update_total_workers():
    total = sum(float(workers_tree.item(item)["values"][2]) for item in workers_tree.get_children() if workers_tree.item(item)["values"][2])
    total_label_workers.config(text=f"الإجمالي: {total:,.2f}")

# دالة لتحميل بيانات قسم العمال من قاعدة البيانات
def load_workers_data():
    for item in workers_tree.get_children():
        workers_tree.delete(item)
    c.execute("SELECT id, date, salary, name FROM workers")
    rows = c.fetchall()
    for i, row in enumerate(rows):
        tag = "evenrow" if i % 2 == 0 else "oddrow"
        workers_tree.insert("", "end", values=row, tags=(tag,))
    update_total_workers()

load_workers_data()

# دالة لإضافة بيانات جديدة في قسم العمال
def add_workers_data():
    add_window = tk.Toplevel(root, bg="#F3F4F6")
    add_window.title("إضافة بيانات العمال")
    add_window.geometry("400x250")
    add_window.resizable(False, False)
    
    tk.Label(add_window, text="إضافة بيانات جديدة", font=("Arial", 16, "bold"), bg="#F3F4F6").pack(pady=15)
    frame = ttk.Frame(add_window)
    frame.pack(padx=20, pady=10)
    
    tk.Label(frame, text="اسم الموظف:", font=("Arial", 12), bg="#F3F4F6").grid(row=0, column=0, padx=10, pady=10, sticky="e")
    name_entry = ttk.Entry(frame, width=25, font=("Arial", 12))
    name_entry.grid(row=0, column=1)
    tk.Label(frame, text="المرتب:", font=("Arial", 12), bg="#F3F4F6").grid(row=1, column=0, padx=10, pady=10, sticky="e")
    salary_entry = ttk.Entry(frame, width=25, font=("Arial", 12))
    salary_entry.grid(row=1, column=1)
    
    def save_workers():
        name, salary = name_entry.get().strip(), salary_entry.get().strip()
        if not name or not salary:
            messagebox.showwarning("خطأ", "يرجى ملء جميع الحقول")
            return
        try:
            salary = float(salary)
            if salary < 0:
                raise ValueError("المرتب يجب أن يكون موجبًا")
            date = datetime.now().strftime("%Y-%m-%d")  # إضافة التاريخ تلقائيًا
            with conn:
                c.execute("INSERT INTO workers (name, salary, date) VALUES (?, ?, ?)", (name, salary, date))
            tag = "evenrow" if len(workers_tree.get_children()) % 2 == 0 else "oddrow"
            workers_tree.insert("", "end", values=(c.lastrowid, date, salary, name), tags=(tag,))
            update_total_workers()
            add_window.destroy()
        except ValueError as e:
            messagebox.showwarning("خطأ", str(e) or "يرجى إدخال قيمة عددية صحيحة")

    ttk.Button(add_window, text="حفظ", command=save_workers).pack(pady=20)

# دالة لتعديل بيانات موجودة في قسم العمال
def edit_workers_data():
    selected = workers_tree.selection()
    if not selected:
        messagebox.showwarning("خطأ", "يرجى تحديد سجل لتعديله")
        return
    item_values = workers_tree.item(selected[0])["values"]
    edit_window = tk.Toplevel(root, bg="#F3F4F6")
    edit_window.title("تعديل بيانات العمال")
    edit_window.geometry("400x250")
    edit_window.resizable(False, False)
    
    tk.Label(edit_window, text="تعديل البيانات", font=("Arial", 16, "bold"), bg="#F3F4F6").pack(pady=15)
    frame = ttk.Frame(edit_window)
    frame.pack(padx=20, pady=10)
    
    tk.Label(frame, text="اسم الموظف:", font=("Arial", 12), bg="#F3F4F6").grid(row=0, column=0, padx=10, pady=10, sticky="e")
    name_entry = ttk.Entry(frame, width=25, font=("Arial", 12))
    name_entry.insert(0, item_values[3])
    name_entry.grid(row=0, column=1)
    tk.Label(frame, text="المرتب:", font=("Arial", 12), bg="#F3F4F6").grid(row=1, column=0, padx=10, pady=10, sticky="e")
    salary_entry = ttk.Entry(frame, width=25, font=("Arial", 12))
    salary_entry.insert(0, item_values[2])
    salary_entry.grid(row=1, column=1)
    
    def save_edit():
        name, salary = name_entry.get().strip(), salary_entry.get().strip()
        if not name or not salary:
            messagebox.showwarning("خطأ", "يرجى ملء جميع الحقول")
            return
        try:
            salary = float(salary)
            if salary < 0:
                raise ValueError("المرتب يجب أن يكون موجبًا")
            with conn:
                c.execute("UPDATE workers SET name=?, salary=? WHERE id=?", (name, salary, item_values[0]))
            workers_tree.item(selected[0], values=(item_values[0], item_values[1], salary, name))
            update_total_workers()
            edit_window.destroy()
        except ValueError as e:
            messagebox.showwarning("خطأ", str(e) or "يرجى إدخال قيمة عددية صحيحة")

    ttk.Button(edit_window, text="حفظ التغييرات", command=save_edit).pack(pady=20)

# دالة لحذف بيانات من قسم العمال
def delete_workers_data():
    selected = workers_tree.selection()
    if selected and messagebox.askyesno("تأكيد", "هل تريد حذف السجل المحدد؟"):
        values = workers_tree.item(selected[0])["values"]
        with conn:
            c.execute("DELETE FROM workers WHERE id=?", (values[0],))
        workers_tree.delete(selected[0])
        update_total_workers()

# أزرار الإضافة، التعديل، والحذف في قسم العمال
button_frame_workers = ttk.Frame(workers_frame)
button_frame_workers.pack(pady=10, fill="x")
ttk.Button(button_frame_workers, text="إضافة", command=add_workers_data).pack(side="left", padx=5)
ttk.Button(button_frame_workers, text="تعديل", command=edit_workers_data).pack(side="left", padx=5)
ttk.Button(button_frame_workers, text="حذف", command=delete_workers_data).pack(side="left", padx=5)

# -------------------------------------- قسم المصنع --------------------------------------
# إعداد إطار قسم المصنع
factory_frame = ttk.Frame(notebook, padding=10)
notebook.add(factory_frame, text="اليوميات")

# إعداد الجدول (Treeview) مع شريط تمرير
factory_scroll = ttk.Scrollbar(factory_frame, orient="vertical")
factory_tree = ttk.Treeview(factory_frame, columns=("ID", "Date", "Cost", "Description"), show="headings", style="Treeview", yscrollcommand=factory_scroll.set)
factory_scroll.config(command=factory_tree.yview)
factory_scroll.pack(side="right", fill="y")
factory_tree.heading("ID", text="المعرف")
factory_tree.heading("Date", text="التاريخ")
factory_tree.heading("Cost", text="التكلفة")
factory_tree.heading("Description", text="الوصف")
# ضبط عرض الأعمدة
factory_tree.column("ID", width=100, anchor="center")
factory_tree.column("Date", width=200, anchor="center")
factory_tree.column("Cost", width=200, anchor="center")
factory_tree.column("Description", width=300, anchor="center")
factory_tree.pack(fill="both", expand=True, pady=10)

# إعداد الألوان المتناوبة للصفوف
factory_tree.tag_configure("oddrow", background="#F9FAFB")
factory_tree.tag_configure("evenrow", background="#FFFFFF")

# عرض الإجمالي أسفل الجدول
total_label_factory = ttk.Label(factory_frame, text="الإجمالي: 0", font=("Arial", 14, "bold"))
total_label_factory.pack(pady=5)

# دالة لحساب إجمالي التكاليف في قسم المصنع
def update_total_factory():
    total = sum(float(factory_tree.item(item)["values"][2]) for item in factory_tree.get_children() if factory_tree.item(item)["values"][2])
    total_label_factory.config(text=f"الإجمالي: {total:,.2f}")

# دالة لتحميل بيانات قسم المصنع من قاعدة البيانات
def load_factory_data():
    for item in factory_tree.get_children():
        factory_tree.delete(item)
    c.execute("SELECT id, date, cost, description FROM factory")
    rows = c.fetchall()
    for i, row in enumerate(rows):
        tag = "evenrow" if i % 2 == 0 else "oddrow"
        factory_tree.insert("", "end", values=row, tags=(tag,))
    update_total_factory()

load_factory_data()

# دالة لإضافة بيانات جديدة في قسم المصنع
def add_factory_data():
    add_window = tk.Toplevel(root, bg="#F3F4F6")
    add_window.title("إضافة بيانات المصنع")
    add_window.geometry("400x250")
    add_window.resizable(False, False)
    
    tk.Label(add_window, text="إضافة بيانات جديدة", font=("Arial", 16, "bold"), bg="#F3F4F6").pack(pady=15)
    frame = ttk.Frame(add_window)
    frame.pack(padx=20, pady=10)
    
    tk.Label(frame, text="الوصف:", font=("Arial", 12), bg="#F3F4F6").grid(row=0, column=0, padx=10, pady=10, sticky="e")
    desc_entry = ttk.Entry(frame, width=25, font=("Arial", 12))
    desc_entry.grid(row=0, column=1)
    tk.Label(frame, text="التكلفة:", font=("Arial", 12), bg="#F3F4F6").grid(row=1, column=0, padx=10, pady=10, sticky="e")
    cost_entry = ttk.Entry(frame, width=25, font=("Arial", 12))
    cost_entry.grid(row=1, column=1)
    
    def save_factory():
        desc, cost = desc_entry.get().strip(), cost_entry.get().strip()
        if not desc or not cost:
            messagebox.showwarning("خطأ", "يرجى ملء جميع الحقول")
            return
        try:
            cost = float(cost)
            if cost < 0:
                raise ValueError("التكلفة يجب أن تكون موجبة")
            date = datetime.now().strftime("%Y-%m-%d")  # إضافة التاريخ تلقائيًا
            with conn:
                c.execute("INSERT INTO factory (description, cost, date) VALUES (?, ?, ?)", (desc, cost, date))
            tag = "evenrow" if len(factory_tree.get_children()) % 2 == 0 else "oddrow"
            factory_tree.insert("", "end", values=(c.lastrowid, date, cost, desc), tags=(tag,))
            update_total_factory()
            add_window.destroy()
        except ValueError as e:
            messagebox.showwarning("خطأ", str(e) or "يرجى إدخال قيمة عددية صحيحة")

    ttk.Button(add_window, text="حفظ", command=save_factory).pack(pady=20)

# دالة لتعديل بيانات موجودة في قسم المصنع
def edit_factory_data():
    selected = factory_tree.selection()
    if not selected:
        messagebox.showwarning("خطأ", "يرجى تحديد سجل لتعديله")
        return
    item_values = factory_tree.item(selected[0])["values"]
    edit_window = tk.Toplevel(root, bg="#F3F4F6")
    edit_window.title("تعديل بيانات المصنع")
    edit_window.geometry("400x250")
    edit_window.resizable(False, False)
    
    tk.Label(edit_window, text="تعديل البيانات", font=("Arial", 16, "bold"), bg="#F3F4F6").pack(pady=15)
    frame = ttk.Frame(edit_window)
    frame.pack(padx=20, pady=10)
    
    tk.Label(frame, text="الوصف:", font=("Arial", 12), bg="#F3F4F6").grid(row=0, column=0, padx=10, pady=10, sticky="e")
    desc_entry = ttk.Entry(frame, width=25, font=("Arial", 12))
    desc_entry.insert(0, item_values[3])
    desc_entry.grid(row=0, column=1)
    tk.Label(frame, text="التكلفة:", font=("Arial", 12), bg="#F3F4F6").grid(row=1, column=0, padx=10, pady=10, sticky="e")
    cost_entry = ttk.Entry(frame, width=25, font=("Arial", 12))
    cost_entry.insert(0, item_values[2])
    cost_entry.grid(row=1, column=1)
    
    def save_edit():
        desc, cost = desc_entry.get().strip(), cost_entry.get().strip()
        if not desc or not cost:
            messagebox.showwarning("خطأ", "يرجى ملء جميع الحقول")
            return
        try:
            cost = float(cost)
            if cost < 0:
                raise ValueError("التكلفة يجب أن تكون موجبة")
            with conn:
                c.execute("UPDATE factory SET description=?, cost=? WHERE id=?", (desc, cost, item_values[0]))
            factory_tree.item(selected[0], values=(item_values[0], item_values[1], cost, desc))
            update_total_factory()
            edit_window.destroy()
        except ValueError as e:
            messagebox.showwarning("خطأ", str(e) or "يرجى إدخال قيمة عددية صحيحة")

    ttk.Button(edit_window, text="حفظ التغييرات", command=save_edit).pack(pady=20)

# دالة لحذف بيانات من قسم المصنع
def delete_factory_data():
    selected = factory_tree.selection()
    if selected and messagebox.askyesno("تأكيد", "هل تريد حذف السجل المحدد؟"):
        values = factory_tree.item(selected[0])["values"]
        with conn:
            c.execute("DELETE FROM factory WHERE id=?", (values[0],))
        factory_tree.delete(selected[0])
        update_total_factory()

# أزرار الإضافة، التعديل، والحذف في قسم المصنع
button_frame_factory = ttk.Frame(factory_frame)
button_frame_factory.pack(pady=10, fill="x")
ttk.Button(button_frame_factory, text="إضافة", command=add_factory_data).pack(side="left", padx=5)
ttk.Button(button_frame_factory, text="تعديل", command=edit_factory_data).pack(side="left", padx=5)
ttk.Button(button_frame_factory, text="حذف", command=delete_factory_data).pack(side="left", padx=5)

# -------------------------------------- قسم التقارير --------------------------------------
# إعداد إطار قسم التقارير
reports_frame = ttk.Frame(notebook, padding=10)
notebook.add(reports_frame, text="تقارير")

tk.Label(reports_frame, text="تقرير الخزنة", font=("Arial", 18, "bold"), bg="#F3F4F6").pack(pady=20)
report_frame = ttk.Frame(reports_frame)
report_frame.pack(pady=10, fill="both", expand=True)

# دالة لتحديث التقارير (تم تصحيح حساب إجمالي الخارج)
def update_reports():
    # حساب إجمالي الداخل
    c.execute("SELECT SUM(batch) FROM income")
    income_total = c.fetchone()[0] or 0
    
    # حساب إجمالي الخارج (البضائع + العمال + المصنع)
    c.execute("SELECT SUM(cost) FROM goods")
    goods_total = c.fetchone()[0] or 0
    c.execute("SELECT SUM(salary) FROM workers")
    workers_total = c.fetchone()[0] or 0
    c.execute("SELECT SUM(cost) FROM factory")
    factory_total = c.fetchone()[0] or 0
    expenses_total = goods_total + workers_total + factory_total
    
    # حساب الرصيد الحالي
    balance = income_total - expenses_total
    
    # تحديث النصوص في واجهة التقارير
    income_label.config(text=f"إجمالي الداخل: {income_total:,.2f}")
    expenses_label.config(text=f"إجمالي الخارج: {expenses_total:,.2f}")
    balance_label.config(text=f"الرصيد الحالي: {balance:,.2f}")

# إعداد النصوص في واجهة التقارير
income_label = ttk.Label(report_frame, text="إجمالي الداخل: 0", font=("Arial", 14))
income_label.pack(pady=10)
expenses_label = ttk.Label(report_frame, text="إجمالي الخارج: 0", font=("Arial", 14))
expenses_label.pack(pady=10)
balance_label = ttk.Label(report_frame, text="الرصيد الحالي: 0", font=("Arial", 16, "bold"))
balance_label.pack(pady=15)
update_reports()

# -------------------------------------- دوال الأزرار العلوية --------------------------------------
# دالة تغيير كلمة المرور
def change_password_window():
    change_window = tk.Toplevel(root, bg="#F3F4F6")
    change_window.title("تغيير كلمة المرور")
    change_window.geometry("400x300")
    change_window.resizable(False, False)

    tk.Label(change_window, text="تغيير كلمة المرور", font=("Arial", 16, "bold"), bg="#F3F4F6").pack(pady=15)
    frame = ttk.Frame(change_window)
    frame.pack(padx=20, pady=10)

    tk.Label(frame, text="كلمة المرور القديمة:", font=("Arial", 12), bg="#F3F4F6").grid(row=0, column=0, padx=10, pady=10, sticky="e")
    old_password_entry = ttk.Entry(frame, width=25, font=("Arial", 12), show="*")
    old_password_entry.grid(row=0, column=1)

    tk.Label(frame, text="كلمة المرور الجديدة:", font=("Arial", 12), bg="#F3F4F6").grid(row=1, column=0, padx=10, pady=10, sticky="e")
    new_password_entry = ttk.Entry(frame, width=25, font=("Arial", 12), show="*")
    new_password_entry.grid(row=1, column=1)

    tk.Label(frame, text="تأكيد كلمة المرور الجديدة:", font=("Arial", 12), bg="#F3F4F6").grid(row=2, column=0, padx=10, pady=10, sticky="e")
    confirm_password_entry = ttk.Entry(frame, width=25, font=("Arial", 12), show="*")
    confirm_password_entry.grid(row=2, column=1)

    def save_new_password():
        old_password = old_password_entry.get().strip()
        new_password = new_password_entry.get().strip()
        confirm_password = confirm_password_entry.get().strip()

        c.execute("SELECT password FROM settings WHERE id = 1")
        current_password = c.fetchone()[0]

        if not old_password or not new_password or not confirm_password:
            messagebox.showwarning("خطأ", "يرجى ملء جميع الحقول")
            return
        if old_password != current_password:
            messagebox.showerror("خطأ", "كلمة المرور القديمة غير صحيحة")
            return
        if new_password != confirm_password:
            messagebox.showerror("خطأ", "كلمة المرور الجديدة وتأكيدها غير متطابقين")
            return
        if len(new_password) < 3:
            messagebox.showwarning("خطأ", "كلمة المرور الجديدة يجب أن تكون 3 أحرف على الأقل")
            return

        with conn:
            c.execute("UPDATE settings SET password = ? WHERE id = 1", (new_password,))
        global PASSWORD
        PASSWORD = new_password
        messagebox.showinfo("نجاح", "تم تغيير كلمة المرور بنجاح!")
        change_window.destroy()

    ttk.Button(change_window, text="حفظ", command=save_new_password).pack(pady=20)

# دالة لتحديث جميع البيانات في الأقسام
def refresh_all():
    root.config(cursor="wait")
    load_income_data()
    load_goods_data()
    load_workers_data()
    load_factory_data()
    update_reports()
    search_name_entry.delete(0, tk.END)
    search_date_entry.delete(0, tk.END)
    search_batch_entry.delete(0, tk.END)
    root.config(cursor="")
    messagebox.showinfo("تحديث", "تم تحديث جميع البيانات بنجاح!")

# دالة لتصدير البيانات إلى ملف Excel
def export_to_excel():
    if messagebox.askyesno("تأكيد", "هل تريد تصدير البيانات إلى ملف Excel؟"):
        root.config(cursor="wait")
        income_data = [income_tree.item(item)["values"] for item in income_tree.get_children()]
        income_df = pd.DataFrame(income_data, columns=["المعرف", "التاريخ", "الاسم", "الدفعة"])
        
        goods_data = [goods_tree.item(item)["values"] for item in goods_tree.get_children()]
        goods_df = pd.DataFrame(goods_data, columns=["المعرف", "التاريخ", "التكلفة", "الصنف"])
        
        workers_data = [workers_tree.item(item)["values"] for item in workers_tree.get_children()]
        workers_df = pd.DataFrame(workers_data, columns=["المعرف", "التاريخ", "المرتب", "اسم الموظف"])
        
        factory_data = [factory_tree.item(item)["values"] for item in factory_tree.get_children()]
        factory_df = pd.DataFrame(factory_data, columns=["المعرف", "التاريخ", "التكلفة", "الوصف"])
        
        with pd.ExcelWriter("تقرير_الخزنة.xlsx") as writer:
            income_df.to_excel(writer, sheet_name="الداخل", index=False)
            goods_df.to_excel(writer, sheet_name="البضائع", index=False)
            workers_df.to_excel(writer, sheet_name="العمال", index=False)
            factory_df.to_excel(writer, sheet_name="المصنع", index=False)
        
        root.config(cursor="")
        messagebox.showinfo("تم التصدير", "تم تصدير البيانات إلى ملف Excel بنجاح!")

# دالة لطباعة البيانات
def print_section():
    current_tab = notebook.tab(notebook.select(), "text")
    if current_tab == "الداخل":
        data = "\n".join([f"{item[0]} | {item[1]} | {item[2]} | {item[3]}" for item in [income_tree.item(item)["values"] for item in income_tree.get_children()]])
    elif current_tab == "البضائع":
        data = "\n".join([f"{item[0]} | {item[1]} | {item[2]} | {item[3]}" for item in [goods_tree.item(item)["values"] for item in goods_tree.get_children()]])
    elif current_tab == "العمال":
        data = "\n".join([f"{item[0]} | {item[1]} | {item[2]} | {item[3]}" for item in [workers_tree.item(item)["values"] for item in workers_tree.get_children()]])
    elif current_tab == "المصنع":
        data = "\n".join([f"{item[0]} | {item[1]} | {item[2]} | {item[3]}" for item in [factory_tree.item(item)["values"] for item in factory_tree.get_children()]])
    elif current_tab == "تقارير":
        data = f"{income_label.cget('text')}\n{expenses_label.cget('text')}\n{balance_label.cget('text')}"
    else:
        return
    with open("طباعة.txt", "w", encoding="utf-8") as f:
        f.write(data)
    os.startfile("طباعة.txt", "print")

# دالة لإغلاق البرنامج والاتصال بقاعدة البيانات
def on_closing(): 
    conn.close()
    root.destroy()

root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()