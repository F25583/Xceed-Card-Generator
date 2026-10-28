import os
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def generate_cards(excel_path, images_dir, template_path, output_path):
    try:
        # قراءة ملف البيانات
        df = pd.read_excel(excel_path)
        
        # فتح قالب الباوربوينت الفارغ
        prs = Presentation(template_path)
        blank_slide_layout = prs.slide_layouts[6] # شريحة فارغة تماماً
        
        for index, row in df.iterrows():
            # تأكد من مطابقة أسماء الأعمدة في ملف الإكسل الخاص بك (ID و Name)
            emp_id = str(row['ID']).strip().split('.')[0] # إزالة الكسور إن وجدت
            emp_name = str(row['Name']).strip()
            
            # إنشاء شريحة جديدة لكل موظف
            slide = prs.slides.add_slide(blank_slide_layout)
            
            # 1. إضافة الخلفية (تصميم كارت Xceed المرفوع)
            slide.shapes.add_picture(template_path, Inches(0), Inches(0), width=prs.slide_width, height=prs.slide_height)
            
            # 2. البحث عن صورة الموظف ومطابقتها بالـ ID
            image_extensions = ['.jpg', '.jpeg', '.png', '.JPG', '.PNG']
            img_added = False
            for ext in image_extensions:
                img_path = os.path.join(images_dir, f"{emp_id}{ext}")
                if os.path.exists(img_path):
                    # المقاسات ومكان الصورة (يمكنك تعديلها بالبوصة حسب رغبتك)
                    slide.shapes.add_picture(img_path, Inches(1.5), Inches(2.0), width=Inches(2.5), height=Inches(3.2))
                    img_added = True
                    break
            
            # 3. إضافة النصوص فوق التصميم
            # اسم الموظف
            txBox_name = slide.shapes.add_textbox(Inches(0.5), Inches(5.5), Inches(4.5), Inches(0.5))
            tf_name = txBox_name.text_frame
            tf_name.word_wrap = True
            p_name = tf_name.paragraphs[0]
            p_name.text = emp_name
            p_name.font.size = Pt(22)
            p_name.font.bold = True
            p_name.font.name = 'Arial'
            
            # الرقم التعريفي ID
            txBox_id = slide.shapes.add_textbox(Inches(0.5), Inches(6.1), Inches(4.5), Inches(0.5))
            tf_id = txBox_id.text_frame
            p_id = tf_id.paragraphs[0]
            p_id.text = f"ID: {emp_id}"
            p_id.font.size = Pt(18)
            p_id.font.name = 'Arial'

        # حفظ الملف النهائي للمراجعة والطباعة
        prs.save(output_path)
        messagebox.showinfo("نجاح", "تم إنشاء الكروت داخل ملف الباوربوينت بنجاح!")
    except Exception as e:
        messagebox.showerror("خطأ", f"حدث خطأ أثناء التشغيل:\n{str(e)}")

# بناء الواجهة الرسومية (GUI)
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Xceed Card Generator")
        self.root.geometry("600x400")
        self.root.resizable(False, False)
        
        # اختيار ملف الإكسل
        tk.Label(root, text="ملف بيانات الموظفين (Excel):", font=("Arial", 11)).pack(anchor="w", px=20, py=5)
        self.excel_entry = tk.Entry(root, width=55)
        self.excel_entry.pack(side="top", anchor="w", px=20)
        tk.Button(root, text="تصفح...", command=self.browse_excel).pack(anchor="w", px=20, py=2)
        
        # اختيار مجلد الصور
        tk.Label(root, text="مجلد صور الموظفين (اسم الصورة = ID):", font=("Arial", 11)).pack(anchor="w", px=20, py=5)
        self.images_entry = tk.Entry(root, width=55)
        self.images_entry.pack(side="top", anchor="w", px=20)
        tk.Button(root, text="تصفح...", command=self.browse_images).pack(anchor="w", px=20, py=2)

        # اختيار قالب التصميم (صورة الكارت الفارغ)
        tk.Label(root, text="صورة تصميم الكارت الفارغ (Template):", font=("Arial", 11)).pack(anchor="w", px=20, py=5)
        self.template_entry = tk.Entry(root, width=55)
        self.template_entry.pack(side="top", anchor="w", px=20)
        tk.Button(root, text="تصفح...", command=self.browse_template).pack(anchor="w", px=20, py=2)

        # زر التشغيل النهائي
        tk.Button(root, text="إنشاء ملف الباوربوينت والطباعة", font=("Arial", 12, "bold"), bg="#007acc", fg="white", px=20, py=10, command=self.start_process).pack(py=30)

    def browse_excel(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, filename)

    def browse_images(self):
        directory = filedialog.askdirectory()
        if directory:
            self.images_entry.delete(0, tk.END)
            self.images_entry.insert(0, directory)

    def browse_template(self):
        filename = filedialog.askopenfilename(filetypes=[("Image files", "*.png *.jpg *.jpeg")])
        if filename:
            self.template_entry.delete(0, tk.END)
            self.template_entry.insert(0, filename)

    def start_process(self):
        excel = self.excel_entry.get()
        images = self.images_entry.get()
        template = self.template_entry.get()
        
        if not excel or not images or not template:
            messagebox.showwarning("تنبيه", "الرجاء تحديد جميع الملفات والمجلدات المطلوبة أولاً!")
            return
            
        output = filedialog.asksaveasfilename(defaultextension=".pptx", filetypes=[("PowerPoint presentations", "*.pptx")])
        if output:
            generate_cards(excel, images, template, output)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
