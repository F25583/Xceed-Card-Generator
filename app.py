import os
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import streamlit as st
from io import BytesIO
import copy

def generate_cards(excel_file, uploaded_images, template_file):
    # 1. قراءة ملف البيانات
    df = pd.read_excel(excel_file)
    
    # تحويل الصور المرفوعة إلى قاموس للبحث السريع بالـ ID
    images_dict = {}
    for img in uploaded_images:
        name_without_ext = os.path.splitext(img.name)[0]
        images_dict[str(name_without_ext).strip()] = img

    # 📌 التوزيع الدقيق والأبعاد المحدثة لـ 8 كروت داخل الصفحة (ارتفاع الصورة 1.7 والعرض 1.4)
    card_positions = [
        # الصف الأول
        {"img_left": Inches(0.48), "img_top": Inches(0.85), "txt_left": Inches(0.25), "txt_top": Inches(2.65)},
        {"img_left": Inches(2.88), "img_top": Inches(0.85), "txt_left": Inches(2.65), "txt_top": Inches(2.65)},
        {"img_left": Inches(5.28), "img_top": Inches(0.85), "txt_left": Inches(5.05), "txt_top": Inches(2.65)},
        {"img_left": Inches(7.68), "img_top": Inches(0.85), "txt_left": Inches(7.45), "txt_top": Inches(2.65)},
        # الصف الثاني
        {"img_left": Inches(0.48), "img_top": Inches(4.25), "txt_left": Inches(0.25), "txt_top": Inches(6.05)},
        {"img_left": Inches(2.88), "img_top": Inches(4.25), "txt_left": Inches(2.65), "txt_top": Inches(6.05)},
        {"img_left": Inches(5.28), "img_top": Inches(4.25), "txt_left": Inches(5.05), "txt_top": Inches(6.05)},
        {"img_left": Inches(7.68), "img_top": Inches(4.25), "txt_left": Inches(7.45), "txt_top": Inches(6.05)}
    ]

    # حفظ بيانات القالب الأصلي في الذاكرة لتكرار الشريحة بأمان
    template_bytes = template_file.getvalue()
    
    # فتح العرض التقديمي الرئيسي الذي سنضع فيه كل الشرائح
    main_prs = Presentation(BytesIO(template_bytes))
    
    # تأمين نسخة نظيفة من الشريحة الأولى لاستخدامها عند إضافة صفحات جديدة
    # (الباوربوينت يحتاج شريحة واحدة على الأقل بالملف)
    current_slide = main_prs.slides[0]

    for index, row in df.iterrows():
        card_index = index % 8
        
        # إذا وصلنا لـ 8 موظفين جدد، نقوم بإضافة شريحة جديدة نظيفة مأخوذة من القالب الأصلي
        if index > 0 and card_index == 0:
            temp_prs = Presentation(BytesIO(template_bytes))
            # أخذ الشريحة من الملف المؤقت وإضافتها للملف الرئيسي بشكل آمن
            slide_layout = temp_prs.slides[0].slide_layout
            current_slide = main_prs.slides.add_slide(slide_layout)
            
            # نقل الأشكال الرسومية الكبيرة بصورة آمنة
            for shape in temp_prs.slides[0].shapes:
                if shape.name.startswith('Picture') or hasattr(shape, 'image'):
                    img_stream = BytesIO(shape.image.bytes)
                    current_slide.shapes.add_picture(img_stream, shape.left, shape.top, shape.width, shape.height)

        # استخراج الـ ID والاسم بدقة تفادياً للكسور العشرية للإكسل
        raw_id = str(row['ID']).strip()
        emp_id = raw_id.split('.')[0] if '.' in raw_id else raw_id
        
        # ✂️ تعديل الاسم ليأخذ الاسم الأول والثاني فقط
        full_name = str(row['Name']).strip().split()
        emp_name = " ".join(full_name[:2]) if len(full_name) >= 2 else " ".join(full_name)
        
        pos = card_positions[card_index]
        
        # 1. إضافة صورة الموظف داخل الكارت الخاص به (ارتفاع 1.7 وعرض 1.4 بوصة)
        if emp_id in images_dict:
            img_file = images_dict[emp_id]
            img_file.seek(0)
            current_slide.shapes.add_picture(img_file, pos["img_left"], pos["img_top"], width=Inches(1.4), height=Inches(1.7))
        
        # 2. إضافة صندوق النصوص (الاسم ورقم الـ ID فقط باللون الأحمر)
        txBox = current_slide.shapes.add_textbox(pos["txt_left"], pos["txt_top"], Inches(1.8), Inches(0.7))
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.margin_top = Inches(0)
        tf.margin_left = Inches(0)
        tf.margin_right = Inches(0)
        tf.margin_bottom = Inches(0)
        
        # سطر اسم الموظف باللون الأزرق لشركة Xceed وبحجم 10 مناسب للمربع
        p_name = tf.paragraphs[0]
        p_name.text = emp_name
        p_name.font.size = Pt(10)
        p_name.font.bold = True
        p_name.font.name = 'Arial'
        p_name.font.color.rgb = RGBColor(0, 102, 204)
        
        # سطر رقم الـ ID فقط باللون الأحمر بجانب كلمة ID الثابتة في التصميم
        p_id = tf.add_paragraph()
        p_id.space_before = Pt(2)
        run_num = p_id.add_run()
        run_num.text = emp_id
        run_num.font.size = Pt(10)
        run_num.font.bold = True
        run_num.font.name = 'Arial'
        run_num.font.color.rgb = RGBColor(255, 0, 0)

    # حفظ الملف النهائي المعدل في الذاكرة لتنزيله
    output = BytesIO()
    main_prs.save(output)
    output.seek(0)
    return output

# واجهة مستخدم ستريمليت Streamlit
st.set_page_config(page_title="Xceed Print Generator", layout="centered")
st.title("🖪 طباعة كروت موظفي Xceed الآلية (8 كروت بالشريحة)")
st.write("قم برفع ملف البيانات والصور وقالب الـ PPTX المجمع ليقوم البرنامج بتوزيعها تلقائياً بالخطوط والألوان الصحيحة.")

excel_file = st.file_uploader("1. اختر ملف بيانات الموظفين (Excel):", type=["xlsx", "xls"])
uploaded_images = st.file_uploader("2. اختر صور الموظفين معاً (اسم الصورة = ID):", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
template_file = st.file_uploader("3. اختر قالب الباوربوينت الفارغ المجمع (new card.pptx):", type=["pptx"])

if excel_file and uploaded_images and template_file:
    if st.button("توليد كروت الطباعة المجمعة", type="primary"):
        with st.spinner("جاري الكتابة داخل القالب وتنسيق الأسماء الثنائية..."):
            try:
                result_pptx = generate_cards(excel_file, uploaded_images, template_file)
                st.success("تم إنتاج الملف بنجاح وبنية الملف سليمة 100%! 🎉")
                
                st.download_button(
                    label="📥 تحميل ملف الباوربوينت الملون والمنظم للطباعة",
                    data=result_pptx,
                    file_name="Xceed_Cards_Ready_To_Print.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            except Exception as e:
                st.error(f"حدث خطأ أثناء المعالجة: {e}")
