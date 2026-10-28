import os
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import streamlit as st
from io import BytesIO

def generate_cards(excel_file, uploaded_images, template_file):
    df = pd.read_excel(excel_file)
    prs = Presentation(template_file)
    
    images_dict = {}
    for img in uploaded_images:
        name_without_ext = os.path.splitext(img.name)[0]
        images_dict[str(name_without_ext).strip()] = img

    # 📌 التوزيع الدقيق والجديد بالملي لـ 8 كروت داخل الصفحة بناءً على القالب المرفق:
    # تم حساب المسافات لضمان سقوط الصورة والنص داخل الإطار الأزرق لكل كارت
    card_positions = [
        # الصف الأول (من اليسار إلى اليمين)
        {"img_left": Inches(0.48), "img_top": Inches(0.95), "txt_left": Inches(0.25), "txt_top": Inches(2.7)},
        {"img_left": Inches(2.88), "img_top": Inches(0.95), "txt_left": Inches(2.65), "txt_top": Inches(2.7)},
        {"img_left": Inches(5.28), "img_top": Inches(0.95), "txt_left": Inches(5.05), "txt_top": Inches(2.7)},
        {"img_left": Inches(7.68), "img_top": Inches(0.95), "txt_left": Inches(7.45), "txt_top": Inches(2.7)},
        # الصف الثاني (من اليسار إلى اليمين)
        {"img_left": Inches(0.48), "img_top": Inches(4.35), "txt_left": Inches(0.25), "txt_top": Inches(6.1)},
        {"img_left": Inches(2.88), "img_top": Inches(4.35), "txt_left": Inches(2.65), "txt_top": Inches(6.1)},
        {"img_left": Inches(5.28), "img_top": Inches(4.35), "txt_left": Inches(5.05), "txt_top": Inches(6.1)},
        {"img_left": Inches(7.68), "img_top": Inches(4.35), "txt_left": Inches(7.45), "txt_top": Inches(6.1)}
    ]

    template_bytes = template_file.getvalue()
    
    # تنظيف وتفريغ الشرائح الافتراضية
    for i in range(len(prs.slides)-1, -1, -1):
        rId = prs.slides._sldIdLst[i].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[i]

    current_slide = None

    for index, row in df.iterrows():
        card_index = index % 8
        
        if card_index == 0:
            # فتح نسخة نظيفة ومستقلة لكل شريحة مجمعة تحتوي على الـ 8 كروت كخلفية ثابتة
            prs_temp = Presentation(BytesIO(template_bytes))
            slide_layout = prs_temp.slide_layouts[0]
            current_slide = prs.slides.add_slide(slide_layout)
            
            # نسخ جميع التصميمات الهندسية والخلفيات الأصلية لضمان ظهور الكروت الزرقاء الفارغة
            for shape in prs_temp.slides[0].shapes:
                try:
                    if shape.name.startswith('Picture') or hasattr(shape, 'image'):
                        img_stream = BytesIO(shape.image.bytes)
                        current_slide.shapes.add_picture(img_stream, shape.left, shape.top, shape.width, shape.height)
                    elif shape.has_table or shape.has_chart:
                        continue
                    else:
                        # نسخ الأشكال التلقائية الأخرى إن وجدت
                        el = shape.element
                        import copy
                        new_el = copy.deepcopy(el)
                        current_slide.shapes._spTree.append(new_el)
                except:
                    continue

        raw_id = str(row['ID']).strip()
        emp_id = raw_id.split('.')[0] if '.' in raw_id else raw_id
        emp_name = str(row['Name']).strip()
        
        pos = card_positions[card_index]
        
        # 1. إسقاط صورة الموظف وضبطها لتتوسط الجزء العلوي من الكارت (العرض 1.3 والارتفاع 1.6 بوصة)
        if emp_id in images_dict:
            img_file = images_dict[emp_id]
            img_file.seek(0)
            current_slide.shapes.add_picture(img_file, pos["img_left"], pos["img_top"], width=Inches(1.3), height=Inches(1.6))
        
        # 2. إسقاط وتنسيق النصوص (الاسم والـ ID فقط باللون الأحمر) لتتطابق مع البيانات
        txBox = current_slide.shapes.add_textbox(pos["txt_left"], pos["txt_top"], Inches(1.8), Inches(0.7))
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.margin_top = Inches(0)
        tf.margin_left = Inches(0)
        tf.margin_right = Inches(0)
        tf.margin_bottom = Inches(0)
        
        # فقرة الاسم (اللون الأزرق لـ Xceed وحجم خط 10 ليناسب المساحة)
        p_name = tf.paragraphs[0]
        p_name.text = emp_name
        p_name.font.size = Pt(10)
        p_name.font.bold = True
        p_name.font.name = 'Arial'
        p_name.font.color.rgb = RGBColor(0, 102, 204)
        
        # فقرة رقم الـ ID (باللون الأحمر الزاهي وحجم 10 ليسقط مباشرة بعد كلمة ID المطبوعة في قالبك)
        p_id = tf.add_paragraph()
        p_id.space_before = Pt(4)
        run_num = p_id.add_run()
        run_num.text = emp_id
        run_num.font.size = Pt(10)
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
st.write("قم برفع ملف البيانات والصور وقالب الـ PPTX المجمع ليقوم البرنامج بتوزيعها تلقائياً بالخطوط والألوان الصحيحة.")

excel_file = st.file_uploader("1. اختر ملف بيانات الموظفين (Excel):", type=["xlsx", "xls"])
uploaded_images = st.file_uploader("2. اختر صور الموظفين معاً (اسم الصورة = ID):", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
template_file = st.file_uploader("3. اختر قالب الباوربوينت الفارغ المجمع (new card.pptx):", type=["pptx"])

if excel_file and uploaded_images and template_file:
    if st.button("توليد كروت الطباعة المجمعة", type="primary"):
        with st.spinner("جاري تنسيق الألوان وتوزيع الموظفين داخل شرائح الطباعة..."):
            try:
                result_pptx = generate_cards(excel_file, uploaded_images, template_file)
                st.success("تم تنسيق وتوزيع كافة الموظفين وجاهز للتحميل والطباعة فوراً! 🎉")
                
                st.download_button(
                    label="📥 تحميل ملف الباوربوينت الملون والمنظم للطباعة",
                    data=result_pptx,
                    file_name="Xceed_Cards_Ready_To_Print.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            except Exception as e:
                st.error(f"حدث خطأ أثناء المعالجة: {e}")
