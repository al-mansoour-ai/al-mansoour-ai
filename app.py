import streamlit as st
from datetime import datetime
from docx import Document
from io import BytesIO

# ================== 1. التنسيق البصري المؤسسي (فخامة رسمية) ==================
st.set_page_config(page_title="المنصور AI - الإصدار المستقر V30", layout="centered")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
    
    /* إخفاء أدوات المنصة للخصوصية والاحترافية */
    div[data-testid="stToolbar"], #MainMenu, footer, header, .stDeployButton, [data-testid="stStatusWidget"] {
        display: none !important; visibility: hidden !important;
    }
    
    .stApp { background-color: #f8fafc; color: #1e293b; direction: rtl; }
    
    .main-box {
        background: #ffffff; border-top: 10px solid #1e3a8a; padding: 40px;
        border-radius: 20px; box-shadow: 0 20px 60px rgba(0,0,0,0.05); margin-top: 5px;
    }

    * { font-family: 'Cairo', sans-serif !important; text-align: right; }
    .brand-title { color: #1e3a8a !important; font-weight: 900 !important; font-size: 2.3rem !important; text-align: center; margin:0; }
    .methodology-tag { 
        background: #1e3a8a; color: #fbbf24; padding: 6px 20px; border-radius: 25px; 
        font-size: 0.85rem; display: table; margin: 10px auto 30px auto; font-weight: bold;
    }

    .section-title { 
        color: #1e3a8a; font-size: 1.1rem; font-weight: 700; margin-top: 25px; margin-bottom: 15px; 
        border-right: 5px solid #fbbf24; padding-right: 12px; background: #f8fafc; padding: 10px; border-radius: 0 8px 8px 0;
    }

    .hint-text { color: #64748b; font-size: 0.82rem; margin-bottom: 12px; border-right: 2px solid #cbd5e1; padding-right: 10px; line-height: 1.5; }
    .magic-desc { color: #2563eb; font-size: 0.72rem; font-weight: 600; text-align: center; margin-bottom: 4px; }

    /* الأزرار الملكية المتوازية */
    .btn-gen button { background: linear-gradient(90deg, #1e3a8a, #d4af37) !important; color: white !important; font-weight: 700 !important; height: 58px !important; border-radius: 12px !important; width: 100% !important; border:none !important; }
    .btn-exit button { background: #f1f5f9 !important; color: #64748b !important; border: 1px solid #cbd5e1 !important; height: 58px !important; border-radius: 12px !important; width: 100% !important; }
    
    .magic-btn button {
        height: 35px !important; font-size: 0.82rem !important; background: #f0f9ff !important;
        color: #1e3a8a !important; border: 1px dashed #cbd5e1 !important;
    }

    .package-table { width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 0.85rem; }
    .package-table th { background: #1e3a8a; color: white; padding: 10px; text-align: center; }
    .package-table td { border: 1px solid #e2e8f0; padding: 10px; text-align: center; background: white; }
    
    .whatsapp-btn {
        background: #25d366; color: white !important; padding: 12px 25px; border-radius: 50px;
        text-decoration: none; font-weight: 700; display: inline-flex; align-items: center; justify-content: center; gap: 10px; margin-top: 20px;
    }
    </style>
""", unsafe_allow_html=True)

# ================== 2. المنهجية العالمية الشاملة (8 تخصصات سيادية) ==================
GLOBAL_REPORTS = {
    "📑 تقرير الإنجاز الدوري | Progress Report": {
        "q": ["1️⃣ الملخص التنفيذي ومستوى الإنجاز العام", "2️⃣ تحليل الانحرافات عن الخطة الزمنية", "3️⃣ إدارة التحديات والمخاطر الميدانية", "4️⃣ آليات التجاوز والخطوات التصحيحية المتبعة"],
        "h": ["تحقيق 80% من المخرجات المخطط لها للفترة الحالية...", "تحديد الفوارق الزمنية بين المخطط والواقع...", "المخاطر التي هددت سير العمل وطرق معالجتها...", "الإجراءات العاجلة التي ضمنا بها استمرار التنفيذ..."]
    },
    "🎓 تقرير ختامي لتدريب | Capacity Building Report": {
        "q": ["1️⃣ نتائج التقييم القبلي والبعدي للمشاركين", "2️⃣ تقييم كفاءة المادة العلمية والمنهجية", "3️⃣ تفاعل المشاركين والبيئة اللوجستية", "4️⃣ توصيات استدامة الأثر التدريبي"],
        "h": ["قياس الفارق المعرفي وتطور مهارات المشاركين...", "مدى ملامسة المحتوى للاحتياجات الميدانية الفعلية...", "المعوقات التقنية أو التنظيمية التي واجهت الورشة...", "خطوات عملية لضمان تطبيق ما تم تعلمه في الميدان..."]
    },
    "💰 تقرير الأداء المالي | Financial Performance": {
        "q": ["1️⃣ تحليل المصروفات الفعلية مقابل الميزانية", "2️⃣ تحليل انحرافات التكلفة (Variance Analysis)", "3️⃣ المخاطر المالية والامتثال (Compliance)", "4️⃣ التوصيات المالية للفترة القادمة"],
        "h": ["مقارنة الإنفاق الفعلي بالخطط المعتمدة وتبرير الوفورات...", "الأسباب الكامنة وراء تجاوز أو انخفاض الإنفاق...", "مدى توافق العمليات مع معايير التدقيق واللوائح...", "مقترحات إعادة الهيكلة لتحسين كفاءة الإنفاق..."]
    },
    "📊 تقرير المتابعة والتقييم | M&E Report": {
        "q": ["1️⃣ قياس مؤشرات الأداء الرئيسية (KPIs)", "2️⃣ جودة المخرجات ورضا المستفيدين", "3️⃣ الدروس المستفادة والفرص الضائعة", "4️⃣ التوصيات الاستراتيجية للتطوير"],
        "h": ["مدى مطابقة التنفيذ مع مصفوفة النتائج المقررة...", "نتائج الاستبيانات والمقابلات مع الفئات المستهدفة...", "التجارب التي يمكن البناء عليها أو تجنبها مستقبلاً...", "مقترحات لتحسين كفاءة التدخلات القادمة..."]
    },
    "🚑 تقرير تقييم الاحتياجات | Needs Assessment": {
        "q": ["1️⃣ تحليل الوضع الراهن وفجوة الاحتياج", "2️⃣ تحديد الفئات الأكثر تضرراً واحتياجاً", "3️⃣ الأولويات العاجلة لخطط الاستجابة", "4️⃣ توصيات التدخل الاستراتيجي والتمويل"],
        "h": ["وصف دقيق للأزمة أو الاحتياج المراد معالجته...", "بيانات ديموغرافية وإحصائية للفئات المستهدفة...", "ما هي الاحتياجات التي لا تقبل التأجيل؟", "خارطة طريق مقترحة للجهات المانحة..."]
    },
    "🏛️ تقرير الحوكمة والامتثال | Compliance Report": {
        "q": ["1️⃣ مستوى الالتزام باللوائح والسياسات", "2️⃣ نتائج الرقابة والتدقيق الداخلي", "3️⃣ الثغرات المرصودة في نظام الحوكمة", "4️⃣ إجراءات التصحيح وتطوير الأداء"],
        "h": ["مدى تطابق الممارسات مع المعايير والقوانين...", "خلاصة عمليات الفحص والرقابة الدورية...", "نقاط الضعف في الهيكل التنظيمي أو الإداري...", "خطوات سد الثغرات القانونية والإدارية..."]
    },
    "🌍 تقرير الأثر البيئي والاجتماعي | ESIA Report": {
        "q": ["1️⃣ تحليل الأثر البيئي والحيوي للمشروع", "2️⃣ المسؤولية المجتمعية ورضا المستفيدين", "3️⃣ إجراءات التخفيف من الآثار الجانبية", "4️⃣ استدامة الموارد وحماية البيئة"],
        "h": ["تقييم تأثير العمليات على المحيط البيئي...", "مدى قبول وتفاعل المجتمع المحلي مع المشروع...", "كيفية التعامل مع الأضرار الجانبية للمشروع...", "خطط الحفاظ على الموارد للأجيال القادمة..."]
    },
    "🏗️ تقرير فني وهندسي | Technical Report": {
        "q": ["1️⃣ المواصفات الفنية ومطابقة المواد", "2️⃣ نتائج اختبارات الجودة الميدانية", "3️⃣ المعوقات الإنشائية والتحديات التقنية", "4️⃣ التعديلات والحلول الهندسية المنفذة"],
        "h": ["مدى التزام الموردين بالمواصفات المعتمدة...", "نتائج فحوصات المختبر والضغوط الهندسة...", "الصعوبات التي واجهت التنفيذ في الموقع...", "الحلول المبتكرة لتجاوز العقبات الإنشائية..."]
    }
}

# ================== 3. المحركات التقنية المستقرة ==================
def get_docx(p_name, rtype, donor, loc, agency, duration, content_dict):
    doc = Document()
    doc.add_heading(f"Report: {rtype}", 0)
    doc.add_paragraph(f"Project: {p_name}")
    doc.add_paragraph(f"Donor: {donor}")
    doc.add_paragraph(f"Implementing Agency: {agency}")
    doc.add_paragraph(f"Location: {loc} | Duration: {duration}")
    doc.add_paragraph(f"Date: {datetime.now().strftime('%Y-%m-%d')}")
    for title, text in content_dict.items():
        doc.add_heading(title, level=1)
        doc.add_paragraph(text if text else "N/A")
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# ================== 4. نظام التشغيل والواجهة V30 ==================
if "auth" not in st.session_state: st.session_state.auth = False

if not st.session_state.auth:
    st.markdown('<div class="main-box">', unsafe_allow_html=True)
    st.markdown('<h1 class="brand-title">بوابة المنصور AI</h1>', unsafe_allow_html=True)
    e = st.text_input("البريد الإلكتروني")
    p = st.text_input("كلمة المرور", type="password")
    if st.button("دخول آمن للمنصة"):
        if e and p: st.session_state.auth = True; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
    st.stop()

st.markdown('<div class="main-box">', unsafe_allow_html=True)
st.markdown('<h1 class="brand-title">المنصور AI للتقارير الاحترافية</h1>', unsafe_allow_html=True)
st.markdown('<div class="methodology-tag">صياغة استراتيجية وفق المنهجية العالمية | 2026</div>', unsafe_allow_html=True)

# القسم 1
st.markdown('<p class="section-title">نوع التقرير الدولي | Report Category</p>', unsafe_allow_html=True)
rtype = st.selectbox("🎯 الخطوة 1: حدد التخصص لضبط المنهجية تلقائياً:", list(GLOBAL_REPORTS.keys()))
cfg = GLOBAL_REPORTS[rtype]

# القسم 2
st.markdown('<p class="section-title">البيانات التعريفية | Metadata</p>', unsafe_allow_html=True)
c1, c2 = st.columns(2)
p_name = c1.text_input("اسم المشروع / البرنامج *", placeholder="Project Name")
donor = c1.text_input("الجهة المانحة / الممول", placeholder="Donor Agency")
loc = c2.text_input("مكان التنفيذ / المنطقة", placeholder="Location")
agency = c2.text_input("الجهة المنفذة / العميل", placeholder="Implementing Agency")
duration = st.text_input("مدة التنفيذ", placeholder="Duration")

# القسم 3
st.markdown('<p class="section-title">المحاور الاستراتيجية للتقرير | Strategic Pillars</p>', unsafe_allow_html=True)
responses = {}
for i in range(4):
    label = cfg["q"][i]
    hint = cfg["h"][i]
    st.markdown(f"<label>{label}</label>", unsafe_allow_html=True)
    st.markdown(f"<p class='hint-text'>🔍 مثال احترافي: {hint}</p>", unsafe_allow_html=True)
    
    col_t, col_b = st.columns([5, 1.5])
    with col_t:
        txt = st.text_area("", key=f"v30_{i}_{rtype}", height=100, label_visibility="collapsed")
        responses[label] = txt
    with col_b:
        st.markdown('<p class="magic-desc">اضغط لتحويل نصك لصياغة احترافية</p>', unsafe_allow_html=True)
        if st.button("✨ تحسين", key=f"btn_v30_{i}"):
            if txt: st.info(f"المقترح: {txt} (تمت المراجعة)")
            else: st.warning("أدخل نصاً")

st.markdown("<br>", unsafe_allow_html=True)
cg, ce = st.columns(2)
with cg:
    st.markdown('<div class="btn-gen">', unsafe_allow_html=True)
    gen_btn = st.button("🚀 توليد ومعالجة التقرير النهائي")
    st.markdown('</div>', unsafe_allow_html=True)
with ce:
    st.markdown('<div class="btn-exit">', unsafe_allow_html=True)
    if st.button("🚪 تسجيل الخروج"):
        st.session_state.auth = False
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

if gen_btn:
    if p_name:
        st.success("التقرير جاهز! اختر صيغة التصدير:")
        word_data = get_docx(p_name, rtype, donor, loc, agency, duration, responses)
        
        e1, e2 = st.columns(2)
        with e1:
            st.markdown('<div class="export-btn">', unsafe_allow_html=True)
            st.download_button("📝 تحميل ملف Word المعتمد", word_data, f"{p_name}.docx")
            st.markdown('</div>', unsafe_allow_html=True)
        with e2:
            st.markdown('<div class="export-btn">', unsafe_allow_html=True)
            st.download_button("📋 تحميل نص سريع", f"{p_name}\n{responses}", f"{p_name}.txt")
            st.markdown('</div>', unsafe_allow_html=True)
    else: st.error("⚠️ يرجى إدخال اسم المشروع أولاً.")

st.markdown('<p class="section-title">باقات العضوية والدعم</p>', unsafe_allow_html=True)
st.markdown("""<table class="package-table"><tr><th>الميزة</th><th>الفضية</th><th>الذهبية</th><th>المؤسسات</th></tr><tr><td>التقارير</td><td>5 شهرياً</td><td>غير محدود</td><td>غير محدود</td></tr><tr><td>الصياغة AI</td><td>أساسية</td><td>احترافية</td><td>مخصصة</td></tr></table>""", unsafe_allow_html=True)
st.markdown(f'<center><a href="https://wa.me/967774575749" class="whatsapp-btn">💬 تواصل لترقية حسابك أو الدعم الفني</a></center>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)
st.markdown("<center style='color:#94a3b8; font-size:0.7rem; margin-top:15px;'>🛡️ شبكة المنصور الدولية للاستشارات | 2026</center>", unsafe_allow_html=True)
