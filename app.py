import streamlit as st
from datetime import datetime
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO

# ================== 1. الهوية البصرية السيادية (V35) ==================
st.set_page_config(page_title="المنصور استراتيجي - V35", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
    div[data-testid="stToolbar"], #MainMenu, footer, header, .stDeployButton { display: none !important; }
    .stApp { background-color: #0c0c0c; color: #ffffff; direction: rtl; }
    .main-box { background: #151515; border: 1.5px solid #d4af37; padding: 35px; border-radius: 15px; box-shadow: 0 10px 50px rgba(0,0,0,0.8); }
    * { font-family: 'Cairo', sans-serif !important; text-align: right; }
    .brand-title { color: #d4af37 !important; font-weight: 900; font-size: 2.8rem; text-align: center; margin: 0; }
    .section-title { color: #d4af37; border-right: 6px solid #d4af37; padding-right: 15px; font-weight: 700; background: #1f1f1f; padding: 10px; margin: 20px 0; border-radius: 0 8px 8px 0; }
    .q-label { color: #fbbf24; font-weight: 600; margin-top: 15px; display: block; font-size: 0.95rem; }
    
    /* تنسيق التبويبات الذهبي */
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] { 
        background-color: #1a1a1a; border: 1px solid #333; color: #aaa; border-radius: 10px 10px 0 0; 
        padding: 12px 25px; font-weight: 600;
    }
    .stTabs [aria-selected="true"] { background-color: #d4af37 !important; color: #000 !important; font-weight: 900 !important; }
    
    .btn-gen button { 
        background: linear-gradient(90deg, #d4af37, #aa8a2e) !important; 
        color: #000 !important; font-weight: 900; height: 65px; font-size: 1.3rem; width: 100%; border-radius: 12px; border: none;
    }
    textarea { background-color: #222 !important; color: white !important; border: 1px solid #444 !important; border-radius: 8px !important; }
    </style>
""", unsafe_allow_html=True)

# ================== 2. محرك التصدير الماسي ==================
def generate_royal_report(p_name, rtype, donor, loc, agency, responses):
    doc = Document()
    doc.add_heading(f"تقرير سيادي معتمد: {rtype}", 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    table = doc.add_table(rows=2, cols=2)
    table.style = 'Table Grid'
    table.rows[0].cells[0].text = f"المشروع: {p_name}"
    table.rows[0].cells[1].text = f"الجهة المانحة: {donor}"
    table.rows[1].cells[0].text = f"الموقع: {loc}"
    table.rows[1].cells[1].text = f"تاريخ الإصدار: {datetime.now().strftime('%Y-%m-%d')}"
    
    for q, a in responses.items():
        doc.add_heading(q, level=1)
        doc.add_paragraph(a if a else "بيان لم يُدرج").alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# ================== 3. البوابة والدخول ==================
if "auth" not in st.session_state: st.session_state.auth = False

if not st.session_state.auth:
    st.markdown('<div class="main-box"><h1 class="brand-title">المنصور AI</h1>', unsafe_allow_html=True)
    if st.button("فتح البوابة السيادية"): st.session_state.auth = True; st.rerun()
    st.stop()

# ================== 4. الواجهة الكاملة (5 مسارات) ==================
st.markdown('<div class="main-box">', unsafe_allow_html=True)
st.markdown('<h1 class="brand-title">المنصور استراتيجي</h1>', unsafe_allow_html=True)
st.markdown('<p style="text-align:center; color:#d4af37; font-weight:700;">منصة الاستنطاق الاستراتيجي الشاملة | الإصدار الخماسي V35</p>', unsafe_allow_html=True)

# البيانات المشتركة
st.markdown('<p class="section-title">📍 البيانات التعريفية الموحدة</p>', unsafe_allow_html=True)
c1, c2 = st.columns(2)
p_name = c1.text_input("اسم المشروع")
donor = c1.text_input("الجهة المانحة")
loc = c2.text_input("الموقع الميداني")
agency = c2.text_input("الجهة المنفذة")

# نظام المسارات الخمسة (تم الإصلاح)
tab1, tab2, tab3, tab4, tab5 = st.tabs(["🛡️ الرقابة", "🎓 الأثر", "⚙️ العمليات", "📈 الاستراتيجية", "📢 العلاقات"])

final_responses = {}
report_type_label = ""

with tab1:
    st.markdown('<p class="section-title">مسار الرقابة والامتثال (Field Visit)</p>', unsafe_allow_html=True)
    q_audit = ["1. نسبة الإنجاز الفعلي مقابل المخطط؟", "2. مخالفات المواصفات الفنية المرصودة؟", "3. رصد أي مظاهر لهدر الموارد؟", "4. الخطر المحدق الذي يهدد استقرار العمل؟", "5. التوصية السيادية (استمرار، تعليق، إنذار)؟"]
    for i, q in enumerate(q_audit):
        st.markdown(f'<p class="q-label">{q}</p>', unsafe_allow_html=True)
        final_responses[q] = st.text_area("", key=f"aud_{i}", height=70, label_visibility="collapsed")
    report_type_label = "رقابة وامتثال - نزول ميداني"

with tab2:
    st.markdown('<p class="section-title">مسار الأثر (Impact - Kirkpatrick)</p>', unsafe_allow_html=True)
    q_impact = ["1. الفارق المعرفي المرصود بعد التدخل؟", "2. المهارة المكتسبة المحددة التي تمت ممارستها؟", "3. دليل تطبيق المهارة في بيئة العمل؟", "4. المعوقات الميدانية لنقل أثر التدريب؟"]
    for i, q in enumerate(q_impact):
        st.markdown(f'<p class="q-label">{q}</p>', unsafe_allow_html=True)
        final_responses[q] = st.text_area("", key=f"imp_{i}", height=70, label_visibility="collapsed")
    report_type_label = "قياس أثر - نموذج كيركباتريك"

with tab3:
    st.markdown('<p class="section-title">مسار العمليات (Operational Performance)</p>', unsafe_allow_html=True)
    q_ops = ["1. القرار الجوهري المتخذ في الاجتماع؟", "2. المسؤول المباشر عن التنفيذ؟", "3. الموعد النهائي القاطع للإنجاز؟", "4. الموارد اللوجستية المطلوبة فوراً؟"]
    for i, q in enumerate(q_ops):
        st.markdown(f'<p class="q-label">{q}</p>', unsafe_allow_html=True)
        final_responses[q] = st.text_area("", key=f"ops_{i}", height=70, label_visibility="collapsed")
    report_type_label = "عمليات وتشغيل - محضر قيادي"

with tab4:
    st.markdown('<p class="section-title">مسار الاستراتيجية (Strategic Risks)</p>', unsafe_allow_html=True)
    q_strat = ["1. التهديد المحتمل للأهداف الكبرى؟", "2. كلفة الخسارة المالية المباشرة المتوقعة؟", "3. الخطة البديلة (B) الجاهزة للتفعيل؟", "4. مستوى الأولوية (حرج، متوسط، منخفض)؟"]
    for i, q in enumerate(q_strat):
        st.markdown(f'<p class="q-label">{q}</p>', unsafe_allow_html=True)
        final_responses[q] = st.text_area("", key=f"str_{i}", height=70, label_visibility="collapsed")
    report_type_label = "استراتيجية - تحليل مخاطر"

with tab5:
    st.markdown('<p class="section-title">مسار العلاقات والظهور (Visibility)</p>', unsafe_allow_html=True)
    q_vis = ["1. الرسالة الاستراتيجية الثلاثية الموجهة؟", "2. أقوى اقتباس قيادي من الحدث؟", "3. صورة الإنجاز المراد ترسيخها لدى المانح؟", "4. الجمهور المستهدف الرئيسي؟"]
    for i, q in enumerate(q_vis):
        st.markdown(f'<p class="q-label">{q}</p>', unsafe_allow_html=True)
        final_responses[q] = st.text_area("", key=f"vis_{i}", height=70, label_visibility="collapsed")
    report_type_label = "علاقات عامة - تقرير إعلامي"

st.markdown("<br>", unsafe_allow_html=True)
if st.button("🚀 توليد التقرير السيادي النهائي"):
    if p_name and any(final_responses.values()):
        word_file = generate_royal_report(p_name, report_type_label, donor, loc, agency, final_responses)
        st.download_button("📥 تحميل المستند الماسي (Word)", word_file, file_name=f"Strategic_Report_{p_name}.docx")
    else: st.error("⚠️ يرجى إدخال اسم المشروع والبيانات المطلوبة.")

st.markdown('</div>', unsafe_allow_html=True)
