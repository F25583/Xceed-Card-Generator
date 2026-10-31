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
    df = pd.read_excel(excel_file)
    
    template_image_file.seek(0)
    pil_img = Image.open(template_image_file)
    img_width_px, img_height_px = pil_img.size
    aspect_ratio = img_width_px / img_height_px
    
    prs = Presentation()
    prs.slide_height = Inches(7.5) 
    prs.slide_width = Inches(7.5 * aspect_ratio) 
    
    blank_layout = prs.slide_layouts if len(prs.slide_layouts) > 6 else prs.slide_layouts
    
    images_dict = {}
    for img in uploaded_images:
        name_without_ext = os.path.splitext(img.name)
        images_dict[str(name_without_ext).strip()] = img

    card_positions = [
        # ─────── الصف العلوي (الأسماء عند 2.75 والـ ID عند 3.06) ───────
        {"img_left": Inches(0.77), "img_top": Inches(1.01), "name_left": Inches(0.56), "name_top": Inches(2.75), "id_left": Inches(1.23), "id_top": Inches(3.06)}, 
        {"img_left": Inches(3.11), "img_top": Inches(1.01), "name_left": Inches(2.86), "name_top": Inches(2.75), "id_left": Inches(3.60), "id_top": Inches(3.06)}, 
        {"img_left": Inches(5.45), "img_top": Inches(1.01), "name_left": Inches(5.20), "name_top": Inches(2.75), "id_left": Inches(5.94), "id_top": Inches(3.06)}, 
        {"img_left": Inches(7.79), "img_top": Inches(1.01), "name_left": Inches(7.54), "name_top": Inches(2.75), "id_left": Inches(8.28), "id_top": Inches(3.06)}, 
        
        # ─────── الصف السفلي (الأسماء عند 6.40 والـ ID عند 6.71) ───────
        {"img_left": Inches(0.77), "img_top": Inches(4.66), "name_left": Inches(0.56), "name_top": Inches(6.40), "id_left": Inches(1.24), "id_top": Inches(6.71)}, 
        {"img_left": Inches(3.15), "img_top": Inches(4.66), "name_left": Inches(2.90), "name_top": Inches(6.40), "id_left": Inches(3.59), "id_top": Inches(6.71)}, 
        {"img_left": Inches(5.53), "img_top": Inches(4.66), "name_left": Inches(5.28), "name_top": Inches(6.40), "id_left": Inches(5.97), "id_top": Inches(6.71)}, 
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
        emp_id = raw_id.split('.') if '.' in raw_id else raw_id
        
        full_name = str(row['Name']).strip().split()
        emp_name = " ".join(full_name[:2]) if len(full_name) >= 2 else " ".join(full_name)
        
        pos = card_positions[card_index]
        
        if emp_id in images_dict:
            img_file = images_dict[emp_id]
            img_file.seek(0)
            current_slide.shapes.add_picture(img_file, pos["img_left"], pos["img_top"], width=Inches(1.22), height=Inches(1.70))
        
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
# ==========================================
# 🎨 واجهة مستخدم ستريمليت الاحترافية والمعدلة بالكامل
# ==========================================

# رابط مباشر ومستقر للوجو لمنع حدوث كسر في الصورة
LOGO_URL = "https://githubusercontent.com"

st.markdown(f"""
    <style>
    /* محاذاة العناوين الرئيسية وتقليل الفراغات المزعجة */
    .main-title {{
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        color: #0b7bc1;
        text-align: center;
        font-weight: 700;
        margin-top: 10px;
        margin-bottom: 5px;
        direction: rtl;
    }}
    .sub-title {{
        text-align: center;
        color: #666666;
        font-size: 15px;
        margin-bottom: 30px;
        direction: rtl;
    }}
    /* تنسيق صناديق رفع الملفات بشكل منسق وجذاب */
    .stFileUploader {{
        background-color: #fcfdfe;
        border: 2px dashed #bcdcfa;
        border-radius: 12px;
        padding: 15px;
        margin-bottom: 20px;
    }}
    /* جعل أزرار التشغيل والتحميل تحتل مساحة واضحة وبلون هوية الشركة */
    div.stButton > button:first-child {{
        background-color: #0b7bc1;
        color: white;
        border-radius: 8px;
        width: 100%;
        font-size: 18px;
        font-weight: bold;
        padding: 12px;
        border: none;
        box-shadow: 0px 4px 6px rgba(11,123,193,0.15);
        transition: 0.3s ease;
    }}
    div.stButton > button:first-child:hover {{
        background-color: #085e94;
        box-shadow: 0px 4px 12px rgba(11,123,193,0.3);
    }}
    </style>
""", unsafe_allow_html=True)

# ─── شريط جانبي منسق للمساعدة (Sidebar) ───
with st.sidebar:
    try:
        st.image(LOGO_URL, width=140)
    except:
        st.markdown("### 🪪 Xceed System")
    st.markdown("<h5 style='text-align: center; color: #555; margin-top: -10px;'>لوحة التحكم الإدارية</h5>", unsafe_allow_html=True)
    st.divider()
    st.markdown("### 🛠️ دليل تشغيل المنظومة")
    st.info("""
    1. **بيانات الموظفين:** ملف Excel يحتوي على الأعمدة الأساسية `ID` و `Name`.
    2. **الصور الشخصية:** رفع الصور معاً على أن يكون اسم الصورة مطابقاً للـ ID تماماً.
    3. **قالب الخلفية:** رفع استمبة الكروت المجمعة الفارغة.
    """)
    st.divider()
    st.caption("تطوير برمي ذكي © 2026 Xceed Print System")

# ─── الواجهة الرئيسية للتطبيق ───
# إظهار اللوجو الرئيسي في المنتصف بشكل صحيح
col_logo1, col_logo2, col_logo3 = st.columns([1, 2, 1])
with col_logo2:
    try:
        st.image(LOGO_URL, use_container_width=True)
    except:
        st.write("")

st.markdown('<h1 class="main-title">منظومة طباعة كروت موظفي Xceed الآلية</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">توليد تلقائي متوازن هندسياً بمعدل 8 كروت مجمعة لكل شريحة طباعة</p>', unsafe_allow_html=True)

st.subheader("📁 لوحة رفع الملفات المستندية")

with st.container():
    excel_file = st.file_uploader("📊 1. قاعدة بيانات الموظفين (ملف Excel):", type=["xlsx", "xls"])
    uploaded_images = st.file_uploader("📸 2. الصور الشخصية للموظفين (يمكن رفع ملفات متعددة معاً):", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
    template_image = st.file_uploader("🖼️ 3. صورة الاستمبة الفارغة المجمعة (Template Background):", type=["jpg", "jpeg", "png"])

st.divider()

if excel_file and uploaded_images and template_image:
    c1, c2 = st.columns(2)
    with c1:
        st.metric(label="عدد الصور المرفوعة", value=f"{len(uploaded_images)} صورة")
    with c2:
        st.success("جميع المستندات جاهزة تماماً للمطابقة!")
        
    st.write("") 
    
    if st.button("⚡ بدء التوليد الفوري والمطابقة الموزونة", type="primary"):
        with st.spinner("جاري صب البيانات وتطبيق هندسة تعويض الترحيل المحدثة..."):
            try:
                result_pptx = generate_cards(excel_file, uploaded_images, template_image)
                st.balloons()
                st.success("🎉 تم توليد الملف بنجاح وتطابق تام 100% مع أبعاد كروتك!")
                
                st.download_button(
                    label="📥 تحميل ملف الباوربوينت النهائي المصدر للطباعة",
                    data=result_pptx,
                    file_name="Xceed_Final_Print_Ready.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            except Exception as e:
                st.error(f"⚠️ حدث خطأ أثناء المعالجة: {e}")
else:
    st.warning("📥 يرجى رفع كافة الملفات الثلاثة الموضحة أعلاه لتفعيل زر التوليد الآلي.")
