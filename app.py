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
    
    # 2. معرفة أبعاد صورة الاستمبة المرفوعة تلقائياً لتجنب تمطيطها
    template_image_file.seek(0)
    pil_img = Image.open(template_image_file)
    img_width_px, img_height_px = pil_img.size
    
    aspect_ratio = img_width_px / img_height_px
    
    prs = Presentation()
    prs.slide_height = Inches(7.5) 
    prs.slide_width = Inches(7.5 * aspect_ratio) 
    
    blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
    
    images_dict = {}
    for img in uploaded_images:
        name_without_ext = os.path.splitext(img.name)[0]
        images_dict[str(name_without_ext).strip()] = img

    # 📌 تحديث الإحداثيات الموزونة بالملي لتوسيط الصور والنصوص تماماً داخل إطار الكروت المرفقة:
    w_scale = prs.slide_width / Inches(10.0) 
    
    card_positions = [
        # الصف العلوي (توسيط كامل)
        {"img_left": Inches(0.53) * w_scale, "img_top": Inches(0.82), "txt_left": Inches(0.20) * w_scale, "txt_top": Inches(2.62)},
        {"img_left": Inches(2.93) * w_scale, "img_top": Inches(0.82), "txt_left": Inches(2.60) * w_scale, "txt_top": Inches(2.62)},
        {"img_left": Inches(5.33) * w_scale, "img_top": Inches(0.82), "txt_left": Inches(5.00) * w_scale, "txt_top": Inches(2.62)},
        {"img_left": Inches(7.73) * w_scale, "img_top": Inches(0.82), "txt_left": Inches(7.40) * w_scale, "txt_top": Inches(2.62)},
        # الصف السفلي (توسيط كامل)
        {"img_left": Inches(0.53) * w_scale, "img_top": Inches(4.27), "txt_left": Inches(0.20) * w_scale, "txt_top": Inches(6.07)},
        {"img_left": Inches(2.93) * w_scale, "img_top": Inches(4.27), "txt_left": Inches(2.60) * w_scale, "txt_top": Inches(6.07)},
        {"img_left": Inches(5.33) * w_scale, "img_top": Inches(4.27), "txt_left": Inches(5.00) * w_scale, "txt_top": Inches(6.07)},
        {"img_left": Inches(7.73) * w_scale, "img_top": Inches(4.27), "txt_left": Inches(7.40) * w_scale, "txt_top": Inches(6.07)}
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
        
        # 1. إضافة صورة الموظف بالمنتصف (العرض 1.3 والارتفاع 1.65 بوصة)
        if emp_id in images_dict:
            img_file = images_dict[emp_id]
            img_file.seek(0)
            current_slide.shapes.add_picture(img_file, pos["img_left"], pos["img_top"], width=Inches(1.30), height=Inches(1.65))
        
        # 2. إضافة وتوسيط النصوص (الاسم بالأزرق، والـ ID بالأحمر أمام كلمة ID الأصلية)
        txBox = current_slide.shapes.add_textbox(pos["txt_left"], pos["txt_top"], Inches(2.35) * w_scale, Inches(0.8))
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.margin_top = Inches(0)
        tf.margin_left = Inches(0)
        tf.margin_right = Inches(0)
        tf.margin_bottom = Inches(0)
        
        # تنسيق وتوسيط الاسم الثنائي
        p_name = tf.paragraphs[0]
        p_name.text = emp_name
        p_name.alignment = PP_ALIGN.CENTER  # توسيط الاسم في المنتصف
        p_name.font.size = Pt(10.5)
        p_name.font.bold = True
        p_name.font.name = 'Arial'
        p_name.font.color.rgb = RGBColor(0, 102, 204)
        
        # تنسيق وتوجيه رقم الـ ID ليقع أمام الكلمة المطبوعة بالخلفية مباشرة وبشكل موسط
        p_id = tf.add_paragraph()
        p_id.space_before = Pt(4)
        p_id.alignment = PP_ALIGN.CENTER  # توسيط السطر لتوزيع المساحة بالتساوي
        
        run_indent = p_id.add_run()
        run_indent.text = "   "  # مسافات مدروسة لترحيل الرقم للأمام قليلاً ليتخطى كلمة ID المطبوعة
        
        run_num = p_id.add_run()
        run_num.text = emp_id
        run_num.font.size = Pt(10.5)
        run_num.font.bold = True
        run_num.font.name = 'Arial'
        run_num.font.color.rgb = RGBColor(255, 0, 0)

    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# واجهة مستخدم ستريمليت Streamlit
st.set_page_config(page_title="Xceed Print Generator", layout="centered")
st.title("🪪 طباعة كروت موظفي Xceed الآلية (8 كروت بالشريحة)")
st.write("قم برفع ملف البيانات والصور وصورة الاستمبة المجمعة ليقوم البرنامج بتوزيعها تلقائياً بالخطوط والألوان الصحيحة وبأبعاد متناسقة.")

excel_file = st.file_uploader("1. اختر ملف بيانات الموظفين (Excel):", type=["xlsx", "xls"])
uploaded_images = st.file_uploader("2. اختر صور الموظفين معاً (اسم الصورة = ID):", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
template_image = st.file_uploader("3. اختر صورة الاستمبة الفارغة المجمعة (Template Image):", type=["jpg", "jpeg", "png"])

if excel_file and uploaded_images and template_image:
    if st.button("توليد كروت الطباعة المجمعة", type="primary"):
        with st.spinner("جاري التوسيط الدقيق للصور وتنسيق الأسماء الثنائية..."):
            try:
                result_pptx = generate_cards(excel_file, uploaded_images, template_image)
                st.success("تم توليد الملف بالتنسيق الموسط والمطابق تماماً للاستمبة المرفوعة! 🎉")
                
                st.download_button(
                    label="📥 تحميل ملف الباوربوينت المنظم للطباعة",
                    data=result_pptx,
                    file_name="Xceed_Cards_Perfect_Match.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            except Exception as e:
                st.error(f"حدث خطأ أثناء المعالجة: {e}")
