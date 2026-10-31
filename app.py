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
    
    images_dict = {}
    for img in uploaded_images:
        name_without_ext = os.path.splitext(img.name)[0]
        images_dict[str(name_without_ext).strip()] = img

    # 📌 التطابق الحرفي والصارم مع إحداثيات المسطرة والشبكة المحدثة:
    # مقاس الصورة المعتمد والمطلوب: العرض = 1.22 بوصة، الارتفاع = 1.70 بوصة
    card_positions = [
        # ── الصف العلوي (تثبيت خط المسطرة الأفقي الأزرق عند 0.92 بوصة بالتمام) ──
        {"img_left": Inches(0.94), "img_top": Inches(0.92), "name_left": Inches(0.50), "name_top": Inches(2.68), "id_left": Inches(1.36), "id_top": Inches(2.95)}, # كارت 1
        {"img_left": Inches(3.34), "img_top": Inches(0.92), "name_left": Inches(2.90), "name_top": Inches(2.68), "id_left": Inches(3.76), "id_top": Inches(2.95)}, # كارت 2
        {"img_left": Inches(5.74), "img_top": Inches(0.92), "name_left": Inches(5.30), "name_top": Inches(2.68), "id_left": Inches(6.16), "id_top": Inches(2.95)}, # كارت 3
        {"img_left": Inches(8.14), "img_top": Inches(0.92), "name_left": Inches(7.70), "name_top": Inches(2.68), "id_left": Inches(8.56), "id_top": Inches(2.95)}, # كارت 4
        
        # ── الصف السفلي (تثبيت خط المسطرة الأفقي الأزرق الجديد عند 4.18 بوصة بالتمام لرفع الكروت) ──
        {"img_left": Inches(0.94), "img_top": Inches(4.18), "name_left": Inches(0.50), "name_top": Inches(5.91), "id_left": Inches(1.36), "id_top": Inches(6.21)}, # كارت 5
        {"img_left": Inches(3.34), "img_top": Inches(4.18), "name_left": Inches(2.90), "name_top": Inches(5.91), "id_left": Inches(3.76), "id_top": Inches(6.21)}, # كارت 6
        {"img_left": Inches(5.74), "img_top": Inches(4.18), "name_left": Inches(5.30), "name_top": Inches(5.91), "id_left": Inches(6.16), "id_top": Inches(6.21)}, # كارت 7
        {"img_left": Inches(8.14), "img_top": Inches(4.18), "name_left": Inches(7.70), "name_top": Inches(5.91), "id_left": Inches(8.56), "id_top": Inches(6.21)}  # كارت 8
    ]

    current_slide = None

    for index, row in df.iterrows():
        card_index = index % 8
        
        if card_index == 0:
            # استدعاء مباشر وآمن تماماً للتخطيط الفارغ القياسي لمنع التعليق السحابي للموقع
            current_slide = prs.slides.add_slide(prs.slide_layouts[6])
            template_image_file.seek(0)
            current_slide.shapes.add_picture(template_image_file, Inches(0), Inches(0), width=prs.slide_width, height=prs.slide_height)

        raw_id = str(row['ID']).strip()
        emp_id = raw_id.split('.')[0] if '.' in raw_id else raw_id
        
        full_name = str(row['Name']).strip().split()
        emp_name = " ".join(full_name[:2]) if len(full_name) >= 2 else " ".join(full_name)
        
        pos = card_positions[card_index]
        
        # 1. إسقاط صورة الموظف (1.7 × 1.22 بوصة) في السنتر تماماً طبقاً للمسطرة المحدثة
        if emp_id in images_dict:
            img_file = images_dict[emp_id]
            img_file.seek(0)
            current_slide.shapes.add_picture(img_file, pos["img_left"], pos["img_top"], width=Inches(1.22), height=Inches(1.70))
        
        # 2. إضافة صندوق الاسم (حجم 14 أزرق موسط ومتناسق مع السنتر الرأسي للخط الأسود)
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
        
        # 3. إضافة صندوق رقم الـ ID (حجم 14 أحمر يبدأ بدقة بعد نقطتي الـ :ID المرسومة مباشرة)
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
st.write("النسخة النهائية المستقرة والمطابقة تماماً لإحداثيات شبكتك المتقاطعة بدون أخطاء.")

excel_file = st.file_uploader("1. اختر ملف بيانات الموظفين (Excel):", type=["xlsx", "xls"])
uploaded_images = st.file_uploader("2. اختر صور الموظفين معاً (اسم الصورة = ID):", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
template_image = st.file_uploader("3. اختر صورة الاستمبة الفارغة المجمعة (Template Image):", type=["jpg", "jpeg", "png"])

if excel_file and uploaded_images and template_image:
    if st.button("توليد كروت الطباعة المجمعة", type="primary"):
        with st.spinner("جاري دمج الكروت بالبنية البرمجية الجديدة الأكثر استقراراً..."):
            try:
                result_pptx = generate_cards(excel_file, uploaded_images, template_image)
                st.success("تم التوليد بنجاح باهر! جاهز للتحميل الفوري وبدون تعليق السيرفر. 🎉")
                
                st.download_button(
                    label="📥 تحميل ملف الباوربوينت النهائي الجاهز للطباعة",
                    data=result_pptx,
                    file_name="Xceed_Cards_Final_Match.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            except Exception as e:
                st.error(f"حدث خطأ أثناء المعالجة: {e}")
