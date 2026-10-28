import os
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
import tkinter as tk
from tkinter import filedialog, messagebox

def generate_cards(excel_path, images_dir, template_path, output_path):
    # قراءة ملف البيانات
    df = pd.read_excel(excel_path)
    
    # فتح قالب الباوربوينت (التصميم الفارغ)
    prs = Presentation(template_path)
    blank_slide_layout = prs.slide_layouts[6] # شريحة فارغة لتجنب النصوص التلقائية
    
    for index, row in df.iterrows():
        emp_id = str(row['ID']).strip()
        emp_name = str(row['Name']).strip()
        
        # إنشاء شريحة جديدة لكل موظف بناءً على التصميم
        slide = prs.slides.add_slide(blank_slide_layout)
        
        # 1. إضافة الخلفية (تصميم كارت Xceed الفارغ)
        slide.shapes.add_picture(template_path, Inches(0), Inches(0), width=prs.slide_width, height=prs.slide_height)
        
        # 2. البحث عن صورة الموظف ومطابقتها بالـ ID
        image_extensions = ['.jpg', '.jpeg', '.png']
        img_added = False
        for ext in image_extensions:
            img_path = os.path.join(images_dir, f"{emp_id}{ext}")
            if os.path.exists(img_path):
                # تحديد مكان الصورة وأبعادها بدقة (تعدل حسب مقاس التصميم)
                slide.shapes.add_picture(img_path, Inches(1.5), Inches(2.0), width=Inches(2.5), height=Inches(3.5))
                img_added = True
                break
        
        # 3. إضافة النصوص (الاسم والـ ID) فوق التصميم
        # صندوق نص الاسم
        txBox_name = slide.shapes.add_textbox(Inches(1.0), Inches(5.8), Inches(3.5), Inches(0.5))
        tf_name = txBox_name.text_frame
        p_name = tf_name.add_paragraph()
        p_name.text = emp_name
        p_name.font.size = Pt(20)
        p_name.font.bold = True
        
        # صندوق نص الـ ID
        txBox_id = slide.shapes.add_textbox(Inches(1.0), Inches(6.3), Inches(3.5), Inches(0.5))
        tf_id = txBox_id.text_frame
        p_id = tf_id.add_paragraph()
        p_id.text = f"ID: {emp_id}"
        p_id.font.size = Pt(16)

    # حفظ الملف النهائي للمراجعة والطباعة
    prs.save(output_path)
    messagebox.showinfo("نجاح", "تم إنشاء ملف الباوربوينت بنجاح جاهز للمراجعة!")
