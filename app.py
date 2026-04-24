import streamlit as st
import google.generativeai as genai
import os
from datetime import datetime
from docx import Document
from io import BytesIO

# 1. التنسيق البصري المؤسسي (فخامة رسمية - كحلي وذهبي)
st.set_page_config(page_title="منصة المنصور الاستراتيجية V30", layout="centered")

st.markdown("""
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;700;900&display=swap" rel="stylesheet">
    <style>
    /* إخفاء أدوات المنصة للخصوصية */
    #MainMenu, footer, header {visibility: hidden;}
    
    html, body, [class*="st-"] { font-family: 'Cairo', sans-serif; direction: rtl; text-align: right; }
    .stApp { background-color: #f8fafc; }
    
    .main-box {
        background: white; border-top: 10px solid #1e3a8a; 
        padding: 40px; border-radius: 20px; 
        box-shadow: 0 20px 60px rgba(0,0,0,0.05);
    }
    .brand-title { color: #1e3a8a; font-weight: 900; font-size: 2.3rem; text-align: center; margin:0; }
    .methodology-tag { 
        background: #1e3a8a; color: #fbbf24; padding: 6px 20px; border-radius: 25px; 
        font-size: 0.85rem; display: table; margin: 10px auto 30px auto; font-weight: bold;
    }
    .section-title { 
        color: #1e3a8a; font-size: 1.1rem; font-weight: 700; margin-top: 25px; 
        border-right: 5px solid #fbbf24; padding-right: 12px; background: #f8fafc; padding: 10px;
    }
    .stButton>button {
        background: linear-gradient(90deg, #1e3a8a, #d4af37) !important;
        color: white !important; font-weight: 700 !important; height: 55px !important; border-radius: 12px !important;
    }
    </style>
""", unsafe_allow_html=True)

# 2. ربط المحرك (Gemini)
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")
except Exception as e:
    st.error("⚠️ يرجى التأكد من إضافة GEMINI_API_KEY في إعدادات Secrets")

# 3. محتوى المنهجية المطور (من دليلك الاستراتيجي)
st.markdown('<div class="main-box">', unsafe_allow_html=True)
st.markdown('<h1 class="brand-title">المنصور AI للتقارير الاستراتيجية</h1>', unsafe_allow_html=True)
st.markdown('<div class="methodology-tag">إصدار 2026 | معايير IBCS & ISO 2145</div>', unsafe_allow_html=True)

tab1, tab2, tab3 = st.tabs(["📝 توليد تقرير", "🔬 دراسة جدوى", "✨ تحسين لغوي"])

with tab1:
    st.markdown('<p class="section-title">إعداد التقرير الدوري والمؤسسي</p>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    rtype = c1.selectbox("نوع التقرير", ["تقرير إنجاز", "تقرير تحليلي", "تقرير مقارن"])
    level = c2.select_slider("مستوى الصياغة", ["رسمي", "احترافي", "قيادي رفيع"])
    
    p_name = st.text_input("اسم المشروع / الجهة")
    raw_data = st.text_area("أدخل البيانات الخام (حتى لو كانت غير مرتبة)", height=150)
    
    if st.button("🚀 صياغة التقرير الاستراتيجي"):
        if raw_data and p_name:
            with st.spinner("جاري تطبيق معايير الكتابة التقنية (Active Voice)..."):
                prompt = f"بصفتك خبير استشاري، صغ لي {rtype} بمستوى {level} للمشروع {p_name}. التزم بنظام ISO 2145، استخدم الصوت النشط، عبارات إيجابية، وجمل قصيرة. أضف توصيات عملية. البيانات: {raw_data}"
                response = model.generate_content(prompt)
                st.session_state['report_text'] = response.text
                st.markdown("---")
                st.markdown(response.text)
        else: st.warning("أدخل البيانات الأساسية أولاً.")

with tab2:
    st.markdown('<p class="section-title">حسابات الجدوى الفنية والقدرة</p>', unsafe_allow_html=True)
    col_a, col_b = st.columns(2)
    rate = col_a.number_input("معدل الإنتاج (وحدة/ساعة)", value=10.0)
    hours = col_b.number_input("ساعات العمل يومياً", value=8)
    days = st.number_input("أيام العمل سنوياً", value=300)
    
    if st.button("📊 تحليل القدرة والجدوى"):
        capacity = rate * hours * days
        st.info(f"القدرة المركبة المحسوبة: {capacity:,} وحدة/سنة")
        with st.spinner("جاري توليد تحليل الجدوى..."):
            f_prompt = f"قدم تحليل جدوى فنية لمشروع قدرته الإنتاجية {capacity} سنوياً، مع تحليل حساسية سريع للمخاطر."
            res = model.generate_content(f_prompt)
            st.write(res.text)

with tab3:
    st.markdown('<p class="section-title">محرك التدقيق (Active Voice)</p>', unsafe_allow_html=True)
    text_to_fix = st.text_area("ألصق النص الضعيف لتحويله لأسلوب قيادي موجز:")
    if st.button("✨ تحسين النص"):
        if text_to_fix:
            fix_prompt = f"أعد صياغة النص بأسلوب Cypress Media (إيجاز، صوت نشط، جمل مثبتة): {text_to_fix}"
            res = model.generate_content(fix_prompt)
            st.success(res.text)

# 4. التصدير لـ Word
if 'report_text' in st.session_state:
    doc = Document()
    doc.add_heading(f"تقرير منصة المنصور: {p_name}", 0)
    doc.add_paragraph(st.session_state['report_text'])
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    st.download_button("💾 تحميل المستند (Word)", bio, f"Mansour_Report.docx")

st.markdown('</div>', unsafe_allow_html=True)
st.markdown("<center style='color:#94a3b8; font-size:0.7rem; margin-top:20px;'>🛡️ شبكة المنصور الدولية للاستشارات | 2026</center>", unsafe_allow_html=True)
