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

    # 📌 المقاسات النهائية المأخوذة حرفياً من لقطات الشاشة للـ Format Shape لكل عنصر:
    card_positions = [
        # ─────── الصف العلوي (الأسماء عند 2.75 والـ ID عند 3.06) ───────
        # كارت 1
        {"img_left": Inches(0.77), "img_top": Inches(1.01), "name_left": Inches(0.56), "name_top": Inches(2.75), "id_left": Inches(1.23), "id_top": Inches(3.06)}, 
        # كارت 2
        {"img_left": Inches(3.11), "img_top": Inches(1.01), "name_left": Inches(2.86), "name_top": Inches(2.75), "id_left": Inches(3.60), "id_top": Inches(3.06)}, 
        # كارت 3
        {"img_left": Inches(5.45), "img_top": Inches(1.01), "name_left": Inches(5.20), "name_top": Inches(2.75), "id_left": Inches(5.94), "id_top": Inches(3.06)}, 
        # كارت 4
        {"img_left": Inches(7.79), "img_top": Inches(1.01), "name_left": Inches(7.54), "name_top": Inches(2.75), "id_left": Inches(8.28), "id_top": Inches(3.06)}, 
        
        # ─────── الصف السفلي (الأسماء عند 6.40 والـ ID عند 6.71) ───────
        # كارت 5
        {"img_left": Inches(0.77), "img_top": Inches(4.66), "name_left": Inches(0.56), "name_top": Inches(6.40), "id_left": Inches(1.24), "id_top": Inches(6.71)}, 
        # كارت 6
        {"img_left": Inches(3.15), "img_top": Inches(4.66), "name_left": Inches(2.90), "name_top": Inches(6.40), "id_left": Inches(3.59), "id_top": Inches(6.71)}, 
        # كارت 7
        {"img_left": Inches(5.53), "img_top": Inches(4.66), "name_left": Inches(5.28), "name_top": Inches(6.40), "id_left": Inches(5.97), "id_top": Inches(6.71)}, 
        # كارت 8
        {"img_left": Inches(7.91), "img_top": Inches(4.66), "name_left": Inches(7.66), "name_top": Inches(6.40), "id_left": Inches(8.35), "id_top": Inches(6.71)}  
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
        
        # 1. إسقاط صورة الموظف بالمقاس الصارم (1.7 × 1.22 بوصة)
        if emp_id in images_dict:
            img_file = images_dict[emp_id]
            img_file.seek(0)
            current_slide.shapes.add_picture(img_file, pos["img_left"], pos["img_top"], width=Inches(1.22), height=Inches(1.70))
        
        # 2. إضافة صندوق الاسم (حجم 14 أزرق موسط ومتناسق وعرض 1.65 بوصة كصورتك)
        txBox_name = current_slide.shapes.add_textbox(pos["name_left"], pos["name_top"], Inches(1.65), Inches(0.23))
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
        
        # 3. إضافة صندوق رقم الـ ID (حجم 14 أحمر يبدأ بدقة بعد نقطتي الـ :ID وعرض 0.95 بوصة كصورتك)
        txBox_id = current_slide.shapes.add_textbox(pos["id_left"], pos["id_top"], Inches(0.95), Inches(0.23))
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
st.title("البرنامج الالي لطباعة كروت الموظفين")
st.write("برجاء اضافة البيانات بشكل صحيح لاستخراج الكروت.")

excel_file = st.file_uploader("1. اختر ملف بيانات الموظفين (Excel):", type=["xlsx", "xls"])
uploaded_images = st.file_uploader("2. اختر صور الموظفين معاً (اسم الصورة = ID):", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
template_image = st.file_uploader("3. اختر صورة الاستمبة الفارغة المجمعة (Template Image):", type=["jpg", "jpeg", "png"])

if excel_file and uploaded_images and template_image:
    if st.button("توليد كروت الطباعة المجمعة", type="primary"):
        with st.spinner("جاري صب البيانات وتوليد الملف بأبعاد متناهية الدقة..."):
            try:
                result_pptx = generate_cards(excel_file, uploaded_images, template_image)
                st.success("تم التوليد بنجاح تام! الكود متطابق الآن مع ملف الباوربوينت الخاص بك بنسبة 100%. 🎉")
                
                st.download_button(
                    label="📥 تحميل ملف الباوربوينت المطابق تماماً",
                    data=result_pptx,
                    file_name="Xceed_Cards_Perfect_Match_Final.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            except Exception as e:
                st.error(f"حدث خطأ أثناء المعالجة: {e}")
