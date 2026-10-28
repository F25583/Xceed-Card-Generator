import os
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import streamlit as st
from io import BytesIO

def generate_cards(excel_file, uploaded_images, template_file):
    # 1. قراءة ملف البيانات
    df = pd.read_excel(excel_file)
    
    # 2. فتح قالب الباوربوينت المجمع (الذي يحتوي على 8 كروت)
    prs = Presentation(template_file)
    blank_slide_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
    
    # تحويل الصور المرفوعة إلى قاموس للبحث السريع بالـ ID
    images_dict = {}
    for img in uploaded_images:
        name_without_ext = os.path.splitext(img.name)[0]
        images_dict[str(name_without_ext).strip()] = img

    # تحديد إحداثيات وأماكن الكروت الثمانية بدقة داخل الشريحة (بالبوصة Inches)
    card_positions = [
        # الصف الأول: كارت 1، 2، 3، 4
        {"img_left": Inches(0.48), "img_top": Inches(0.85), "txt_left": Inches(0.2), "txt_top": Inches(2.2)},
        {"img_left": Inches(2.88), "img_top": Inches(0.85), "txt_left": Inches(2.6), "txt_top": Inches(2.2)},
        {"img_left": Inches(5.28), "img_top": Inches(0.85), "txt_left": Inches(5.0), "txt_top": Inches(2.2)},
        {"img_left": Inches(7.68), "img_top": Inches(0.85), "txt_left": Inches(7.4), "txt_top": Inches(2.2)},
        # الصف الثاني: كارت 5، 6، 7، 8
        {"img_left": Inches(0.48), "img_top": Inches(3.95), "txt_left": Inches(0.2), "txt_top": Inches(5.3)},
        {"img_left": Inches(2.88), "img_top": Inches(3.95), "txt_left": Inches(2.6), "txt_top": Inches(5.3)},
        {"img_left": Inches(5.28), "img_top": Inches(3.95), "txt_left": Inches(5.0), "txt_top": Inches(5.3)},
        {"img_left": Inches(7.68), "img_top": Inches(3.95), "txt_left": Inches(7.4), "txt_top": Inches(5.3)}
    ]

    # حذف الشريحة الأولى الافتراضية الفارغة تماماً للبدء من جديد بالنسخ النظيفة
    for i in range(len(prs.slides)-1, -1, -1):
        rId = prs.slides._sldIdLst[i].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[i]

    current_slide = None

    for index, row in df.iterrows():
        # تحديد موقع الكارت الحالي داخل الشريحة (من 0 إلى 7)
        card_index = index % 8
        
        # إنشاء شريحة جديدة مجمعة عند البداية أو كل 8 موظفين
        if card_index == 0:
            prs_temp = Presentation(template_file)
            current_slide = prs.slides.add_slide(blank_slide_layout)
            # نسخ الخلفية والتصميم الأساسي للـ 8 كروت من القالب
            for shape in prs_temp.slides[0].shapes:
                if shape.has_picture:
                    current_slide.shapes.add_picture(template_file, shape.left, shape.top, shape.width, shape.height)

        emp_id = str(row['ID']).strip().split('.')[0]
        emp_name = str(row['Name']).strip()
        
        pos = card_positions[card_index]
        
        # 1. إضافة صورة الموظف داخل الكارت الخاص به
        if emp_id in images_dict:
            img_file = images_dict[emp_id]
            img_file.seek(0)
            current_slide.shapes.add_picture(img_file, pos["img_left"], pos["img_top"], width=Inches(1.4), height=Inches(1.7))
        
        # 2. إضافة النص المنسق (الاسم ورقم الـ ID فقط باللون الأحمر)
        txBox = current_slide.shapes.add_textbox(pos["txt_left"], pos["txt_top"], Inches(2.0), Inches(0.8))
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.margin_top = Inches(0)
        tf.margin_bottom = Inches(0)
        
        # سطر اسم الموظف باللون الأزرق لشركة Xceed
        p_name = tf.paragraphs[0]
        p_name.text = emp_name
        p_name.font.size = Pt(12)
        p_name.font.bold = True
        p_name.font.name = 'Arial'
        p_name.font.color.rgb = RGBColor(0, 102, 204)  # لون أزرق مطابق للشعار
        
        # سطر رقم الـ ID فقط (باللون الأحمر) ليأتي بجانب كلمة ID الموجودة في قالبك
        p_id = tf.add_paragraph()
        run_num = p_id.add_run()
        run_num.text = emp_id
        run_num.font.size = Pt(11)
        run_num.font.bold = True
        run_num.font.name = 'Arial'
        run_num.font.color.rgb = RGBColor(255, 0, 0)  # أرقام الـ ID باللون الأحمر تماماً

    # حفظ الملف في الذاكرة للتحميل عبر المتصفح
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
