import os
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import streamlit as st
from io import BytesIO

def generate_cards(excel_file, uploaded_images, template_image_file):
    # 1. قراءة ملف البيانات
    df = pd.read_excel(excel_file)
    
    # 2. إنشاء ملف باوربوينت جديد تماماً من الصفر
    prs = Presentation()
    # ضبط أبعاد الصفحة لتكون مقاس العرض الافتراضي القياسي (Widescreen 13.333 × 7.5 بوصة)
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6] # شريحة فارغة تماماً بدون أي عناصر تلقائية
    
    # تحويل الصور المرفوعة إلى قاموس للبحث السريع بالـ ID
    images_dict = {}
    for img in uploaded_images:
        name_without_ext = os.path.splitext(img.name)[0]
        images_dict[str(name_without_ext).strip()] = img

    # 📌 الإحداثيات الموزونة بدقة بالغة ليسقط كل موظف داخل كارت الاستمبة المرفقة تماماً:
    card_positions = [
        # الصف العلوي (الأربعة كروت من اليسار إلى اليمين)
        {"img_left": Inches(0.92), "img_top": Inches(0.72), "txt_left": Inches(0.40), "txt_top": Inches(2.62)},
        {"img_left": Inches(4.00), "img_top": Inches(0.72), "txt_left": Inches(3.48), "txt_top": Inches(2.62)},
        {"img_left": Inches(7.08), "img_top": Inches(0.72), "txt_left": Inches(6.56), "txt_top": Inches(2.62)},
        {"img_left": Inches(10.16), "img_top": Inches(0.72), "txt_left": Inches(9.64), "txt_top": Inches(2.62)},
        # الصف السفلي (الأربعة كروت من اليسار إلى اليمين)
        {"img_left": Inches(0.92), "img_top": Inches(4.28), "txt_left": Inches(0.40), "txt_top": Inches(6.18)},
        {"img_left": Inches(4.00), "img_top": Inches(4.28), "txt_left": Inches(3.48), "txt_top": Inches(6.18)},
        {"img_left": Inches(7.08), "img_top": Inches(4.28), "txt_left": Inches(6.56), "txt_top": Inches(6.18)},
        {"img_left": Inches(10.16), "img_top": Inches(4.28), "txt_left": Inches(9.64), "txt_top": Inches(6.18)}
    ]

    current_slide = None

    for index, row in df.iterrows():
        card_index = index % 8
        
        # إنشاء شريحة جديدة وتركيب صورة الاستمبة كخلفية كاملة عند البداية وكل 8 موظفين
        if card_index == 0:
            current_slide = prs.slides.add_slide(blank_layout)
            template_image_file.seek(0)
            current_slide.shapes.add_picture(template_image_file, Inches(0), Inches(0), width=prs.slide_width, height=prs.slide_height)

        # استخراج البيانات ومعالجة الـ ID لتفادي مشكلة الكسور العشرية للإكسيل
        raw_id = str(row['ID']).strip()
        emp_id = raw_id.split('.')[0] if '.' in raw_id else raw_id
        
        # ✂️ تعديل الاسم ليأخذ الاسم الأول والثاني فقط
        full_name = str(row['Name']).strip().split()
        emp_name = " ".join(full_name[:2]) if len(full_name) >= 2 else " ".join(full_name)
        
        pos = card_positions[card_index]
        
        # 1. إضافة صورة الموظف داخل كارت الاستمبة (المقاس المضبوط: الارتفاع 1.7 والعرض 1.4 بوصة)
        if emp_id in images_dict:
            img_file = images_dict[emp_id]
            img_file.seek(0)
            current_slide.shapes.add_picture(img_file, pos["img_left"], pos["img_top"], width=Inches(1.4), height=Inches(1.7))
        
        # 2. إضافة النص المنسق (الاسم باللون الأزرق ورقم الـ ID باللون الأحمر)
        txBox = current_slide.shapes.add_textbox(pos["txt_left"], pos["txt_top"], Inches(2.4), Inches(0.8))
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.margin_top = Inches(0)
        tf.margin_left = Inches(0)
        tf.margin_right = Inches(0)
        tf.margin_bottom = Inches(0)
        
        # سطر اسم الموظف (أول وثاني اسم) متمركز باللون الأزرق لـ Xceed
        p_name = tf.paragraphs[0]
        p_name.text = emp_name
        p_name.font.size = Pt(11)
        p_name.font.bold = True
        p_name.font.name = 'Arial'
        p_name.font.color.rgb = RGBColor(0, 102, 204)  # لون أزرق مطابق للشعار
        
        # سطر رقم الـ ID فقط باللون الأحمر (يزحف تلقائياً بعد مسافة كلمة :ID الثابتة بالخلفية)
        p_id = tf.add_paragraph()
        p_id.space_before = Pt(4)
        run_indent = p_id.add_run()
        run_indent.text = "        " # مسافة بيضاء ذكية لترحيل الرقم ليأتي بعد كلمة :ID المطبوعة في صورتك تماماً
        
        run_num = p_id.add_run()
        run_num.text = emp_id
        run_num.font.size = Pt(11)
        run_num.font.bold = True
        run_num.font.name = 'Arial'
        run_num.font.color.rgb = RGBColor(255, 0, 0)  # أرقام الـ ID باللون الأحمر

    # حفظ ملف الباوربوينت النهائي الجديد في الذاكرة لتنزيله عبر المتصفح
    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# واجهة مستخدم ستريمليت Streamlit
st.set_page_config(page_title="Xceed Print Generator", layout="centered")
st.title("🪪 طباعة كروت موظفي Xceed الآلية (8 كروت بالشريحة)")
st.write("قم برفع ملف البيانات والصور وصورة الاستمبة المجمعة ليقوم البرنامج بتوزيعها تلقائياً بالخطوط والألوان الصحيحة.")

excel_file = st.file_uploader("1. اختر ملف بيانات الموظفين (Excel):", type=["xlsx", "xls"])
uploaded_images = st.file_uploader("2. اختر صور الموظفين معاً (اسم الصورة = ID):", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
template_image = st.file_uploader("3. اختر صورة الاستمبة الفارغة المجمعة (Template Image):", type=["jpg", "jpeg", "png"])

if excel_file and uploaded_images and template_image:
    if st.button("توليد كروت الطباعة المجمعة", type="primary"):
        with st.spinner("جاري دمج الصور وتنسيق الأسماء الثنائية على صورة الاستمبة..."):
            try:
                result_pptx = generate_cards(excel_file, uploaded_images, template_image)
                st.success("تم توليد الملف بنجاح وبأعلى استقرارية وبدون أخطاء! 🎉")
                
                st.download_button(
                    label="📥 تحميل ملف الباوربوينت الملون والمنظم للطباعة",
                    data=result_pptx,
                    file_name="Xceed_Cards_Ready_To_Print.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            except Exception as e:
                st.error(f"حدث خطأ أثناء المعالجة: {e}")
