import streamlit as st
from datetime import datetime
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO

# ================== 1. الهوية البصرية السيادية ==================
st.set_page_config(page_title="المنصور استراتيجي - V34", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
    div[data-testid="stToolbar"], #MainMenu, footer, header, .stDeployButton { display: none !important; }
    .stApp { background-color: #0a0a0a; color: #ffffff; direction: rtl; }
    .main-box { background: #151515; border: 1.5px solid #d4af37; padding: 30px; border-radius: 15px; }
    * { font-family: 'Cairo', sans-serif !important; }
    .brand-title { color: #d4af37 !important; font-weight: 900; font-size: 2.8rem; text-align: center; margin: 0; }
    .section-title { color: #d4af37; border-right: 5px solid #d4af37; padding-right: 15px; font-weight: 700; background: #1f1f1f; padding: 10px; margin: 20px 0; }
    .q-label { color: #e5e7eb; font-weight: 600; margin-top: 10px; display: block; }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; background-color: #0a0a0a; }
    .stTabs [data-baseweb="tab"] { background-color: #1a1a1a; border: 1px solid #333; color: #888; border-radius: 8px 8px 0 0; padding: 10px 20px; }
    .stTabs [aria-selected="true"] { background-color: #d4af37 !important; color: #000 !important; font-weight: 900; }
    .btn-gen button { background: linear-gradient(90deg, #d4af37, #aa8a2e) !important; color: #000 !important; font-weight: 900; height: 60px; font-size: 1.2rem; width: 100%; border-radius: 12px; }
    </style>
""", unsafe_allow_html=True)

# ================== 2. محرك التصدير (المخرج الماسي) ==================
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

# ================== 3. بنية المسارات السيادية ==================
if "auth" not in st.session_state: st.session_state.auth = False

if not st.session_state.auth:
    st.markdown('<div class="main-box"><h1 class="brand-title">المنصور AI</h1>', unsafe_allow_html=True)
    if st.button("فتح البوابة السيادية"): st.session_state.auth = True; st.rerun()
    st.stop()

st.markdown('<div class="main-box">', unsafe_allow_html=True)
st.markdown('<h1 class="brand-title">المنصور استراتيجي</h1>', unsafe_allow_html=True)
st.markdown('<p style="text-align:center; color:#888;">منصة الاستنطاق الاستراتيجي الشاملة | V34</p>', unsafe_allow_html=True)

# البيانات المشتركة
st.markdown('<p class="section-title">📍 البيانات التعريفية الموحدة</p>', unsafe_allow_html=True)
col1, col2 = st.columns(2)
p_name = col1.text_input("اسم المشروع")
donor = col1.text_input("الجهة المانحة")
loc = col2.text_input("الموقع الجغرافي")
agency = col2.text_input("الجهة المنفذة")

# نظام المسارات (Tabs)
tab1, tab2, tab3, tab4, tab5 = st.tabs(["🛡️ الرقابة", "🎓 الأثر", "⚙️ العمليات", "📈 الاستراتيجية", "📢 العلاقات"])

def render_questions(q_list, key_prefix):
    res = {}
    for i, q in enumerate(q_list):
        st.markdown(f'<p class="q-label">{q}</p>', unsafe_allow_html=True)
        ans = st.text_area("", key=f"{key_prefix}_{i}", height=70, label_visibility="collapsed")
        res[q] = ans
    return res

with tab1:
    st.markdown('<p class="section-title">مسار الرقابة والامتثال (تقرير نزول ميداني)</p>', unsafe_allow_html=True)
    q_audit = ["نسبة الإنجاز الفعلي؟", "مخالفات المواصفات؟", "هدر الموارد؟", "الخطر المحدق؟", "التوصية السيادية؟"]
    responses = render_questions(q_audit, "audit")

with tab2:
    st.markdown('<p class="section-title">مسار الأثر (تقرير التدريب - كيركباتريك)</p>', unsafe_allow_html=True)
    q_impact = ["الفارق المعرفي المرصود؟", "المهارة المكتسبة المحددة؟", "دليل تطبيق المهارة ميدانياً؟", "معوقات نقل الأثر؟"]
    responses = render_questions(q_impact, "impact")

with tab3:
    st.markdown('<p class="section-title">مسار العمليات (محضر اجتماع قيادي)</p>', unsafe_allow_html=True)
    q_ops = ["القرار الجوهري المتخذ؟", "المسؤول عن التنفيذ؟", "الموعد النهائي القاطع؟", "الموارد المطلوبة؟"]
    responses = render_questions(q_ops, "ops")

with tab4:
    st.markdown('<p class="section-title">مسار الاستراتيجية (تقرير مخاطر)</p>', unsafe_allow_html=True)
    q_strat = ["التهديد المحتمل؟", "كلفة الخسارة المالية؟", "الخطة (ب) البديلة؟", "مستوى الأولوية؟"]
    responses = render_questions(q_strat, "strat")

with tab5:
    st.markdown('<p class="section-title">مسار العلاقات (تقرير إعلامي)</p>', unsafe_allow_html=True)
    q_vis = ["الرسالة الثلاثية الموجهة؟", "أهم اقتباس قيادي؟", "صورة الإنجاز المراد ترسيخها؟", "الجمهور المستهدف؟"]
    responses = render_questions(q_vis, "vis")

st.markdown("<br>", unsafe_allow_html=True)
if st.button("🚀 توليد التقرير السيادي النهائي"):
    if p_name:
        word_file = generate_royal_report(p_name, "تقرير استراتيجي شامل", donor, loc, agency, responses)
        st.download_button("📥 تحميل المستند الماسي (Word)", word_file, file_name=f"Mansour_Report_{datetime.now().strftime('%H%M%S')}.docx")
    else: st.error("⚠️ يرجى إدخال اسم المشروع")

st.markdown('</div>', unsafe_allow_html=True)
