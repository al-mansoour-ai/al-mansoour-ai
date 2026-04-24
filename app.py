
import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document

# 1. إعدادات الصفحة
st.set_page_config(page_title="المنصور AI - استشارات", layout="centered")

# 2. تطبيق التصميم (بدون نصوص ظاهرة)
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700;900&display=swap');
    html, body, [class*="st-"] { font-family: 'Cairo', sans-serif; direction: rtl; text-align: right; }
    .stApp { background-color: #f8fafc; }
    #MainMenu, footer, header { visibility: hidden; }
    .main-box {
        background: white; border-top: 10px solid #1e3a8a; 
        padding: 30px; border-radius: 15px; 
        box-shadow: 0 10px 25px rgba(0,0,0,0.05); margin-top: -50px;
    }
    .brand-title { color: #1e3a8a; font-weight: 900; font-size: 2rem; text-align: center; margin-bottom: 5px; }
    .tagline { background: #1e3a8a; color: #fbbf24; padding: 5px 15px; border-radius: 20px; font-size: 0.8rem; display: table; margin: 0 auto 25px auto; font-weight: bold; }
    .stButton>button {
        background: linear-gradient(90deg, #1e3a8a, #d4af37) !important;
        color: white !important; font-weight: bold !important; height: 50px !important; border-radius: 10px !important; border: none !important; width: 100%;
    }
</style>
""", unsafe_allow_html=True)

# 3. الهيكل الرئيسي للمنصة
st.markdown('<div class="main-box">', unsafe_allow_html=True)
st.markdown('<h1 class="brand-title">المنصور AI للتقارير الاستراتيجية</h1>', unsafe_allow_html=True)
st.markdown('<div class="tagline">إصدار 2026 | معايير IBCS & ISO 2145</div>', unsafe_allow_html=True)

# تفعيل الذكاء الاصطناعي
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")
except:
    st.error("⚠️ يرجى التأكد من ضبط API Key في Secrets")

# الأقسام
tab1, tab2 = st.tabs(["📝 توليد تقرير", "✨ تحسين لغوي"])

with tab1:
    rtype = st.selectbox("🎯 نوع التخصص:", ["تقرير إنجاز دوري", "دراسة جدوى فنية", "تقرير متابعة وتقييم"])
    p_name = st.text_input("اسم المشروع / الجهة")
    raw_data = st.text_area("البيانات الخام (نقاط مختصرة):", height=150)
    
    if st.button("🚀 صياغة واحتراف"):
        if raw_data and p_name:
            with st.spinner("جاري التحليل الاستراتيجي..."):
                prompt = f"صغ لي {rtype} احترافي للمشروع {p_name}. استخدم معايير ISO 2145، الصوت النشط، وجمل قصيرة. البيانات: {raw_data}"
                response = model.generate_content(prompt)
                st.markdown("---")
                st.markdown(response.text)
                st.session_state['out'] = response.text
        else: st.warning("أكمل البيانات أولاً")

with tab2:
    st.info("تحسين النصوص وفق نموذج Cypress Media (إيجاز، صوت نشط، قوة)")
    txt_fix = st.text_area("ألصق النص هنا:")
    if st.button("✨ تحسين الآن"):
        if txt_fix:
            res = model.generate_content(f"حول هذا النص لصوت نشط ولهجة قيادية موجزة: {txt_fix}")
            st.success(res.text)

# التصدير
if 'out' in st.session_state:
    doc = Document()
    doc.add_heading(f"تقرير: {p_name}", 0)
    doc.add_paragraph(st.session_state['out'])
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    st.download_button("💾 تحميل ملف Word المعتمد", bio, "Mansour_AI_Report.docx")

st.markdown('</div>', unsafe_allow_html=True)
st.markdown("<center style='color:#94a3b8; font-size:0.7rem; margin-top:20px;'>🛡️ شبكة المنصور الدولية للاستشارات | 2026</center>", unsafe_allow_html=True)
