import sqlite3
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import pandas as pd

# --- إعداد قاعدة البيانات ---
def init_db():
    conn = sqlite3.connect("expenses.db")
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS expenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            amount REAL NOT NULL,
            category TEXT NOT NULL,
            date TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    conn.close()

# --- نافذة التطبيق الرئيسية ---
class ExpenseApp:
    def __init__(self, root):
        self.root = root
        self.root.title("برنامج المصروفات اليومية المتطور")
        self.root.geometry("600x620")
        self.root.config(bg="#f4f4f4")
        
        init_db()
        
        # العنوان الرئيسي
        title_label = tk.Label(root, text="إدارة المصروفات اليومية", font=("Arial", 16, "bold"), bg="#f4f4f4", fg="#333")
        title_label.pack(pady=10)
        
        # إطار الإدخال
        frame_input = tk.LabelFrame(root, text=" إضافة مصروف جديد ", font=("Arial", 11), bg="#f4f4f4", padx=10, pady=10)
        frame_input.pack(fill="x", padx=20, pady=5)
        
        # الوصف
        tk.Label(frame_input, text="بيان المصروف:", bg="#f4f4f4").grid(row=0, column=1, sticky="e", pady=5)
        self.title_entry = tk.Entry(frame_input, width=28, font=("Arial", 11))
        self.title_entry.grid(row=0, column=0, pady=5, padx=5)
        
        # المبلغ
        tk.Label(frame_input, text="المبلغ:", bg="#f4f4f4").grid(row=1, column=1, sticky="e", pady=5)
        self.amount_entry = tk.Entry(frame_input, width=28, font=("Arial", 11))
        self.amount_entry.grid(row=1, column=0, pady=5, padx=5)
        
        # التصنيف
        tk.Label(frame_input, text="التصنيف:", bg="#f4f4f4").grid(row=2, column=1, sticky="e", pady=5)
        self.category_var = tk.StringVar(value="طعام")
        categories = ["طعام", "مواصلات", "فواتير", "ترفيه", "أخرى"]
        self.category_menu = ttk.Combobox(frame_input, textvariable=self.category_var, values=categories, width=26, state="readonly")
        self.category_menu.grid(row=2, column=0, pady=5, padx=5)
        
        # أزرار الإجراءات (حفظ وتصدير)
        btn_frame = tk.Frame(frame_input, bg="#f4f4f4")
        btn_frame.grid(row=3, column=0, columnspan=2, pady=10)
        
        add_btn = tk.Button(btn_frame, text="حفظ المصروف", bg="#28a745", fg="white", font=("Arial", 10, "bold"), width=15, command=self.add_expense)
        add_btn.pack(side="left", padx=5)
        
        export_btn = tk.Button(btn_frame, text="تصدير إلى Excel", bg="#17a2b8", fg="white", font=("Arial", 10, "bold"), width=15, command=self.export_to_excel)
        export_btn.pack(side="left", padx=5)
        
        # إطار العرض والملخص
        frame_view = tk.Frame(root, bg="#f4f4f4")
        frame_view.pack(fill="both", expand=True, padx=20, pady=5)
        
        # جدول عرض البيانات
        self.tree = ttk.Treeview(frame_view, columns=("ID", "Title", "Amount", "Category", "Date"), show="headings", height=8)
        self.tree.heading("ID", text="م")
        self.tree.heading("Title", text="البيان")
        self.tree.heading("Amount", text="المبلغ")
        self.tree.heading("Category", text="التصنيف")
        self.tree.heading("Date", text="التاريخ")
        
        self.tree.column("ID", width=30, anchor="center")
        self.tree.column("Title", width=130, anchor="center")
        self.tree.column("Amount", width=80, anchor="center")
        self.tree.column("Category", width=90, anchor="center")
        self.tree.column("Date", width=130, anchor="center")
        self.tree.pack(fill="both", expand=True, side="left")
        
        # شريط تمرير للجدول
        scrollbar = ttk.Scrollbar(frame_view, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # زر حذف المحدد
        delete_btn = tk.Button(root, text="حذف المصروف المحدد", bg="#dc3545", fg="white", font=("Arial", 10, "bold"), command=self.delete_expense)
        delete_btn.pack(pady=5)
        
        # إجمالي المصروفات
        self.total_label = tk.Label(root, text="إجمالي المصروفات: 0.00", font=("Arial", 12, "bold"), bg="#f4f4f4", fg="#dc3545")
        self.total_label.pack(pady=5)
        
        self.load_expenses()

    def add_expense(self):
        title = self.title_entry.get().strip()
        amount = self.amount_entry.get().strip()
        category = self.category_var.get()
        
        if not title or not amount:
            messagebox.showerror("خطأ", "الرجاء إدخال البيان والمبلغ!")
            return
            
        try:
            amount = float(amount)
        except ValueError:
            messagebox.showerror("خطأ", "المبلغ يجب أن يكون رقماً!")
            return
            
        conn = sqlite3.connect("expenses.db")
        cursor = conn.cursor()
        cursor.execute("INSERT INTO expenses (title, amount, category) VALUES (?, ?, ?)", (title, amount, category))
        conn.commit()
        conn.close()
        
        self.title_entry.delete(0, tk.END)
        self.amount_entry.delete(0, tk.END)
        self.load_expenses()
        messagebox.showinfo("نجاح", "تم إضافة المصروف بنجاح.")

    def delete_expense(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("تنبيه", "الرجاء اختيار المصروف المراد حذفه من الجدول!")
            return
            
        item_data = self.tree.item(selected_item)
        expense_id = item_data['values'][0]
        
        confirm = messagebox.askyesno("تأكيد الحذف", "هل أنت متأكد من حذف هذا المصروف؟")
        if confirm:
            conn = sqlite3.connect("expenses.db")
            cursor = conn.cursor()
            cursor.execute("DELETE FROM expenses WHERE id = ?", (expense_id,))
            conn.commit()
            conn.close()
            
            self.load_expenses()
            messagebox.showinfo("نجاح", "تم حذف المصروف بنجاح.")

    def export_to_excel(self):
        conn = sqlite3.connect("expenses.db")
        df = pd.read_sql_query("SELECT id AS 'المسلسل', title AS 'البيان', amount AS 'المبلغ', category AS 'التصنيف', date AS 'التاريخ' FROM expenses", conn)
        conn.close()
        
        if df.empty:
            messagebox.showwarning("تنبيه", "لا توجد بيانات لتصديرها!")
            return
            
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="حفظ باسم")
        if file_path:
            df.to_excel(file_path, index=False)
            messagebox.showinfo("نجاح", f"تم تصدير البيانات بنجاح إلى:\n{file_path}")

    def load_expenses(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        conn = sqlite3.connect("expenses.db")
        cursor = conn.cursor()
        cursor.execute("SELECT id, title, amount, category, date FROM expenses ORDER BY id DESC")
        rows = cursor.fetchall()
        
        total = 0.0
        for row in rows:
            self.tree.insert("", "end", values=row)
            total += row[2]
            
        conn.close()
        self.total_label.config(text=f"إجمالي المصروفات: {total:.2f}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExpenseApp(root)
    root.mainloop()
