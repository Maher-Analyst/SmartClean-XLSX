import pandas as pd
import re
import os
import customtkinter as ctk
from tkinter import filedialog, messagebox


def clean_text_logic(val):
    """تطهير أعمدة النصوص من الخلايا الرقمية البحتة فقط"""
    if pd.isna(val) or str(val).lower() == 'nan': 
        return 'Unknown'

    text = str(val).strip()
    text = " ".join(text.split())

    # التعديل المطلوب: إذا كانت الخلية "كلها أرقام" فقط، حولها لـ Unknown
    # هذا السطر يتأكد أن النص لا يحتوي إلا على أرقام من البداية للنهاية
    if text.isdigit(): 
        return 'Unknown'

    # إذا كان النص يحتوي على حروف (مثل ahmed1)، سيبقى كما هو ويتحول لـ Title Case فقط
    return text.title()

def clean_numeric_logic(val):
    """تنظيف الأرقام وتصحيح أنواع البيانات"""
    if pd.isna(val): return 0
    val = str(val)
    val = re.sub(r'[$,]', '', val) # حذف رموز العملة
    return pd.to_numeric(val, errors='coerce')

# --- 2. الدالة الرئيسية ---

def process_data():
    file_path = file_entry.get()
    if not file_path:
        messagebox.showwarning("تنبيه", "يرجى اختيار ملف أولاً")
        return

    try:
        df = pd.read_excel(file_path)

        # حذف الصفوف والأعمدة الفارغة تماماً (التحقق المنطقي النهائي)
        df.dropna(how='all', inplace=True)
        df.dropna(axis=1, how='all', inplace=True)
        df.columns = df.columns.str.strip()

        # معالجة النصوص 
        text_input = text_cols_entry.get()
        if text_input.strip():
            t_cols = [c.strip() for c in text_input.split(',')]
            for col in t_cols:
                if col in df.columns:
                    df[col] = df[col].apply(clean_text_logic)

        # معالجة الأرقام
        num_input = num_cols_entry.get()
        if num_input.strip():
            n_cols = [c.strip() for c in num_input.split(',')]
            for col in n_cols:
                if col in df.columns:
                    df[col] = df[col].apply(clean_numeric_logic)

        # حذف التكرار
        dup_input = dup_cols_entry.get()
        if dup_input.strip():
            d_cols = [c.strip() for c in dup_input.split(',')]
            valid_d_cols = [c for c in d_cols if c in df.columns]
            if valid_d_cols:
                df.drop_duplicates(subset=valid_d_cols, keep='first', inplace=True)

        output_name = "Final_Cleaned_Professional.xlsx"
        df.to_excel(output_name, index=False)
        messagebox.showinfo("نجاح", "اكتملت المعالجة! تم الحفاظ على النصوص المركبة بنجاح.")

    except Exception as e:
        messagebox.showerror("خطأ", "حدث خطأ: " + str(e))

# --- 3. الواجهة الحديثة (Modern GUI) ---

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("Amazon Data Cleaner Pro v2.2")
app.geometry("600x620")

# العنوان
title_label = ctk.CTkLabel(app, text="نظام التحقق وتطهير البيانات", font=("Arial", 24, "bold"))
title_label.pack(pady=30)

# اختيار الملف
file_frame = ctk.CTkFrame(app)
file_frame.pack(pady=10, padx=30, fill="x")

file_entry = ctk.CTkEntry(file_frame, placeholder_text="اختر مسار الملف...", width=380)
file_entry.pack(side="left", padx=10, pady=10)

def browse():
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    file_entry.delete(0, "end")
    file_entry.insert(0, path)

browse_btn = ctk.CTkButton(file_frame, text="بحث", width=100, command=browse)
browse_btn.pack(side="right", padx=10)

# الخانات
ctk.CTkLabel(app, text="النصوص اعمدة").pack(pady=(15, 0))
text_cols_entry = ctk.CTkEntry(app, placeholder_text="Name, Category...", width=480)
text_cols_entry.pack(pady=5)

ctk.CTkLabel(app, text="الارقام اعمدة").pack(pady=(15, 0))
num_cols_entry = ctk.CTkEntry(app, placeholder_text="Price, Age...", width=480)
num_cols_entry.pack(pady=5)

ctk.CTkLabel(app, text="التكرار اعمدة").pack(pady=(15, 0))
dup_cols_entry = ctk.CTkEntry(app, placeholder_text="Column names...", width=480)
dup_cols_entry.pack(pady=5)

# زر البدء
start_btn = ctk.CTkButton(app, text="بدء التنظيف الاحترافي", height=55, 
                          font=("Arial", 18, "bold"), fg_color="#2ecc71", hover_color="#27ae60",
                          command=process_data)
start_btn.pack(pady=40)

app.mainloop()

