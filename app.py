import os
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
import streamlit as st
from io import BytesIO

def generate_cards(excel_file, uploaded_images, template_file):
    # قراءة ملف البيانات
    df = pd.read_excel(excel_file)
    
    # فتح قالب الباوربوينت الفارغ
    prs = Presentation(template_file)
    blank_slide_layout = prs.slide_layouts[6] # شريحة فارغة تماماً لتجنب النصوص التلقائية
    
    # تحويل الصور المرفوعة إلى قاموس ليسهل البحث عنها بالـ ID
    images_dict = {}
    for img in uploaded_images:
        name_without_ext = os.path.splitext(img.name)[0]
        images_dict[str(name_without_ext).strip()] = img

    for index, row in df.iterrows():
        emp_id = str(row['ID']).strip().split('.')[0] # إزالة الكسور أو النقط الزائدة إن وجدت
        emp_name = str(row['Name']).strip()
        
        # إنشاء شريحة جديدة لكل موظف
        slide = prs.slides.add_slide(blank_slide_layout)
        
        # 1. إضافة الخلفية (تصميم كارت Xceed الفارغ)
        template_file.seek(0)
        slide.shapes.add_picture(template_file, Inches(0), Inches(0), width=prs.slide_width, height=prs.slide_height)
        
        # 2. البحث عن صورة الموظف ومطابقتها بالـ ID
        if emp_id in images_dict:
            img_file = images_dict[emp_id]
            img_file.seek(0)
            # المقاسات ومكان الصورة (يمكنك تعديلها بالبوصة حسب رغبتك)
            slide.shapes.add_picture(img_file, Inches(1.5), Inches(2.0), width=Inches(2.5), height=3.2)
        
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

    # حفظ الملف النهائي في الذاكرة لتنزيله عبر المتصفح
    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# تصميم واجهة موقع الويب عبر Streamlit
st.set_page_config(page_title="Xceed Card Generator", layout="centered")
st.title("🪪 مصمم كروت موظفي Xceed الآلي")
st.write("قم برفع ملف البيانات والصور لتوليد ملف الباوربوينت جاهزاً للطباعة فوراً.")

# 1. رفع ملف الإكسيل
excel_file = st.file_uploader("اختر ملف بيانات الموظفين (Excel):", type=["xlsx", "xls"])

# 2. رفع صور الموظفين دفعة واحدة
uploaded_images = st.file_uploader("اختر صور الموظفين (يمكنك تحديد عدة صور معاً بشرط اسم الصورة = ID):", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

# 3. رفع قالب التصميم الفارغ
template_file = st.file_uploader("اختر صورة تصميم الكارت الفارغ (Template):", type=["jpg", "jpeg", "png"])

if excel_file and uploaded_images and template_file:
    if st.button("إنشاء ملف الباوربوينت للمراجعة والطباعة", type="primary"):
        with st.spinner("جاري دمج البيانات والصور وإنشاء الكروت..."):
            try:
                result_pptx = generate_cards(excel_file, uploaded_images, template_file)
                st.success("تم إنشاء الكروت بنجاح! 🎉")
                
                # زر تحميل الملف الناتج للمستخدم
                st.download_button(
                    label="📥 تحميل ملف PowerPoint الناتج",
                    data=result_pptx,
                    file_name="Xceed_Staff_Cards.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            except Exception as e:
                st.error(f"حدث خطأ أثناء المعالجة: {e}")
