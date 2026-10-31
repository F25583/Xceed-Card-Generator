import os
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import streamlit as st
from io import BytesIO
from PIL import Image

def generate_cards(excel_file, uploaded_images, template_image_file):
    # 1. قراءة ملف البيانات
    df = pd.read_excel(excel_file)
    
    # 2. قراءة أبعاد صورة الاستمبة المرفوعة تلقائياً لمنع التمطيط
    template_image_file.seek(0)
    pil_img = Image.open(template_image_file)
    img_width_px, img_height_px = pil_img.size
    
    aspect_ratio = img_width_px / img_height_px
    
    prs = Presentation()
    prs.slide_height = Inches(7.5) 
    prs.slide_width = Inches(7.5 * aspect_ratio) 
    
    # استدعاء مباشر ومستقر للتخطيط الفارغ القياسي
    blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
    
    images_dict = {}
    for img in uploaded_images:
        name_without_ext = os.path.splitext(img.name)[0]
        images_dict[str(name_without_ext).strip()] = img

    # 📌 المقاسات النهائية الموزونة والمطابقة 100% لخصائص الـ Format Shape الحقيقية في ملفك:
    card_positions = [
        # ─────── الصف العلوي (تثبيت الـ Top عند 1.01 مع الخطوة المتزنة الناتجة عن مسطرتك) ───────
        # كارت 1
        {"img_left": Inches(0.77), "img_top": Inches(1.01), "name_left": Inches(0.33), "name_top": Inches(2.67), "id_left": Inches(1.19), "id_top": Inches(3.02)}, 
        # كارت 2
        {"img_left": Inches(3.11), "img_top": Inches(1.01), "name_left": Inches(2.67), "name_top": Inches(2.67), "id_left": Inches(3.53), "id_top": Inches(3.02)}, 
        # كارت 3
        {"img_left": Inches(5.45), "img_top": Inches(1.01), "name_left": Inches(5.01), "name_top": Inches(2.67), "id_left": Inches(5.87), "id_top": Inches(3.02)}, 
        # كارت 4
        {"img_left": Inches(7.79), "img_top": Inches(1.01), "name_left": Inches(7.35), "name_top": Inches(2.67), "id_left": Inches(8.21), "id_top": Inches(3.02)}, 
        
        # ─────── الصف السفلي (تثبيت الـ Top عند 4.66 بناءً على لقطات الشاشة الحية الحالية) ───────
        # كارت 5
        {"img_left": Inches(0.77), "img_top": Inches(4.66), "name_left": Inches(0.33), "name_top": Inches(6.32), "id_left": Inches(1.19), "id_top": Inches(6.67)}, 
        # كارت 6
        {"img_left": Inches(3.15), "img_top": Inches(4.66), "name_left": Inches(2.71), "name_top": Inches(6.32), "id_left": Inches(3.57), "id_top": Inches(6.67)}, 
        # كارت 7
        {"img_left": Inches(5.53), "img_top": Inches(4.66), "name_left": Inches(5.09), "name_top": Inches(6.32), "id_left": Inches(5.95), "id_top": Inches(6.67)}, 
        # كارت 8
        {"img_left": Inches(7.91), "img_top": Inches(4.66), "name_left": Inches(7.47), "name_top": Inches(6.32), "id_left": Inches(8.33), "id_top": Inches(6.67)}  
    ]

    current_slide = None

    for index, row in df.iterrows():
        card_index = index % 8
        
        if card_index == 0:
            current_slide = prs.slides.add_slide(blank_layout)
            template_image_file.seek(0)
            current_slide.shapes.add_picture(template_image_file, Inches(0), Inches(0), width=prs.slide_width, height=prs.slide_height)

        raw_id = str(row['ID']).strip()
        emp_id = raw_id.split('.')[0] if '.' in raw_id else raw_id
        
        full_name = str(row['Name']).strip().split()
        emp_name = " ".join(full_name[:2]) if len(full_name) >= 2 else " ".join(full_name)
        
        pos = card_positions[card_index]
        
        # 1. إسقاط صورة الموظف بالمقاس الصارم المطلوب (1.7 × 1.22 بوصة)
        if emp_id in images_dict:
            img_file = images_dict[emp_id]
            img_file.seek(0)
            current_slide.shapes.add_picture(img_file, pos["img_left"], pos["img_top"], width=Inches(1.22), height=Inches(1.70))
        
        # 2. إضافة صندوق الاسم (حجم 14 أزرق موسط ومتناسق)
        txBox_name = current_slide.shapes.add_textbox(pos["name_left"], pos["name_top"], Inches(2.10), Inches(0.35))
        tf_name = txBox_name.text_frame
        tf_name.word_wrap = True
        tf_name.margin_top = tf_name.margin_bottom = tf_name.margin_left = tf_name.margin_right = Inches(0)
        
        p_name = tf_name.paragraphs[0]
        p_name.text = emp_name
        p_name.alignment = PP_ALIGN.CENTER
        p_name.font.size = Pt(14)
        p_name.font.bold = True
        p_name.font.name = 'Arial'
        p_name.font.color.rgb = RGBColor(0, 102, 204)
        
        # 3. إضافة صندوق رقم الـ ID (حجم 14 أحمر يبدأ بدقة بعد نقطتي الـ :ID)
        txBox_id = current_slide.shapes.add_textbox(pos["id_left"], pos["id_top"], Inches(0.90), Inches(0.3))
        tf_id = txBox_id.text_frame
        tf_id.word_wrap = False
        tf_id.margin_top = tf_id.margin_bottom = tf_id.margin_left = tf_id.margin_right = Inches(0)
        
        p_id = tf_id.paragraphs[0]
        p_id.text = emp_id
        p_id.alignment = PP_ALIGN.LEFT
        p_id.font.size = Pt(14)
        p_id.font.bold = True
        p_id.font.name = 'Arial'
        p_id.font.color.rgb = RGBColor(255, 0, 0)

    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# واجهة مستخدم ستريمليت Streamlit
st.set_page_config(page_title="Xceed Print Generator", layout="centered")
st.title("🪪 طباعة كروت موظفي Xceed الآلية (8 كروت بالشريحة)")
st.write("النسخة النهائية المحدثة: تطبيق كامل لأرقام الـ Format Shape الصارمة للصفين بالتفصيل.")

excel_file = st.file_uploader("1. اختر ملف بيانات الموظفين (Excel):", type=["xlsx", "xls"])
uploaded_images = st.file_uploader("2. اختر صور الموظفين معاً (اسم الصورة = ID):", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
template_image = st.file_uploader("3. اختر صورة الاستمبة الفارغة المجمعة (Template Image):", type=["jpg", "jpeg", "png"])

if excel_file and uploaded_images and template_image:
    if st.button("توليد كروت الطباعة المجمعة", type="primary"):
        with st.spinner("جاري تطبيق الأبعاد الحقيقية المحدثة للصفين..."):
            try:
                result_pptx = generate_cards(excel_file, uploaded_images, template_image)
                st.success("تم التوليد بنجاح تام وتطابق 100% لكل الكروت! 🎉")
                
                st.download_button(
                    label="📥 تحميل ملف الباوربوينت النهائي الجاهز للطباعة",
                    data=result_pptx,
                    file_name="Xceed_Cards_Final_Perfect.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            except Exception as e:
                st.error(f"حدث خطأ أثناء المعالجة: {e}")
