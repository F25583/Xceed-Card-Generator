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
    
    # التأكد من جلب تخطيط فارغ تماماً بنسبة 100%
    blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]
    
    images_dict = {}
    for img in uploaded_images:
        name_without_ext = os.path.splitext(img.name)[0]
        images_dict[str(name_without_ext).strip()] = img

    # 📌 المقاسات الصارمة النهائية الخاصة بمسطرتك لمنع الترحيل
    card_positions = [
        {"img_left": Inches(0.77), "img_top": Inches(1.01), "name_left": Inches(0.56), "name_top": Inches(2.75), "id_left": Inches(1.23), "id_top": Inches(3.06)}, 
        {"img_left": Inches(3.11), "img_top": Inches(1.01), "name_left": Inches(2.86), "name_top": Inches(2.75), "id_left": Inches(3.60), "id_top": Inches(3.06)}, 
        {"img_left": Inches(5.45), "img_top": Inches(1.01), "name_left": Inches(5.20), "name_top": Inches(2.75), "id_left": Inches(5.94), "id_top": Inches(3.06)}, 
        {"img_left": Inches(7.79), "img_top": Inches(1.01), "name_left": Inches(7.54), "name_top": Inches(2.75), "id_left": Inches(8.28), "id_top": Inches(3.06)}, 
        
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
            # 🛡️ السطر الحاسم لإجبار الباوربوينت على وضع خلفية الاستمبة في كل شريحة جديدة
            template_image_file.seek(0)
            current_slide.shapes.add_picture(template_image_file, Inches(0), Inches(0), width=prs.slide_width, height=prs.slide_height)

        raw_id = str(row['ID']).strip()
        emp_id = raw_id.split('.')[0] if '.' in raw_id else raw_id
        
        full_name = str(row['Name']).strip().split()
        emp_name = " ".join(full_name[:2]) if len(full_name) >= 2 else " ".join(full_name)
        
        pos = card_positions[card_index]
        
        # إسقاط صورة الموظف
        if emp_id in images_dict:
            img_file = images_dict[emp_id]
            img_file.seek(0)
            current_slide.shapes.add_picture(img_file, pos["img_left"], pos["img_top"], width=Inches(1.22), height=Inches(1.70))
        
        # إضافة مربع الاسم الأزرق
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
        
        # إضافة مربع الـ ID الأحمر
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
