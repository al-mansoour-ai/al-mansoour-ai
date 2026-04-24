import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document
from datetime import datetime

# 1. التنسيق البصري الفخم (بدون تداخل)
st.set_page_config(page_title="منصة المنصور الاستراتيجية", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
    * { font-family: 'Cairo', sans-serif !important; direction: rtl; text-align: right; }
    .stApp { background-color: #f8fafc; }
    
    /* منع تداخل النصوص في القوائم */
    .stSelectbox, .stTextInput, .stTextArea { margin-bottom: 15px !important; }
    
    /* تصميم البطاقات */
    .report-card { 
        background: white; border-radius: 12px; padding: 25px; 
        border-right: 8px solid #1e3a8a; box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        margin-bottom: 20px;
    }
    
    .brand-header { background: #1e3a8a; color: white; padding: 20px; border-radius: 15px; text-align: center; margin-bottom: 30px; }
    .brand-title { font-weight: 900; font-size: 2rem; margin: 0; }
    
    .q-label { color: #1e3a8a; font-weight: 800; font-size: 1.1rem; margin-bottom: 10px; display: block; }
    
    /* الأزرار */
    .stButton>button { 
        background: linear-gradient(90deg, #1e3a8a, #152e6d) !important; color: white !important; 
        font-weight: 700 !important; border-radius: 10px !important; width: 100%; height: 50px;
    }
    .magic-btn button { 
        background: #fff9db !important; color: #856404 !important; border: 1px dashed #fab005 !important; 
        height: 35px !important; font-size: 0.8rem !important; width: auto !important;
    }
</style>
""", unsafe_allow_html=True)

# تفعيل Gemini
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")
except:
    st.error("⚠️ يرجى ضبط المفتاح في الإعدادات")

# بنك الـ 12 تخصصاً كاملاً
STRATEGY_BANK = {
    "📑 تقرير إنجاز دوري": ["الملخص التنفيذي", "تحليل الأنشطة", "إدارة الانحرافات", "الموارد والميزانية", "التحديات والحلول", "الخطة القادمة"],
    "💰 دراسة جدوى استثمارية": ["تحليل الفجوة", "المتطلبات الفنية", "النمذجة المالية", "تحليل الحساسية", "تحليل SWOT", "التوصية النهائية"],
    "🎓 تقرير ختامي لتدريب": ["الأهداف والمنهجية", "نتائج القبلي والبعدي", "تقييم المدرب", "اللوجستيات", "استدامة الأثر"],
    "🔍 تقرير متابعة وتقييم (M&E)": ["قياس KPIs", "جودة المخرجات", "رضا المستفيدين", "الدروس المستفادة", "توصيات التطوير"],
    "🚑 تقرير تقييم احتياجات": ["وصف الاحتياج", "الفئات المتضررة", "الأولويات العاجلة", "الفجوة المتاحة", "خارطة التدخل"],
    "🏛️ تقرير حوكمة وامتثال": ["الالتزام باللوائح", "نتائج الرقابة", "الثغرات المرصودة", "إجراءات التصحيح"],
    "💰 تقرير أداء مالي": ["بيان المصروفات", "تحليل الانحرافات", "التدفق النقدي", "كفاءة الإنفاق"],
    "🏗️ تقرير فني وهندسي": ["المواصفات الفنية", "اختبارات الجودة", "المعوقات والحلول", "السلامة المهنية"],
    "🌍 تقرير أثر بيئي": ["الأثر الحيوي", "المسؤولية المجتمعية", "إجراءات التخفيف", "خطة الاستدامة"],
    "📝 تقرير تحليل مناقصات": ["التقييم الفني", "التقييم المالي", "مخاطر الموردين", "توصية الترسية"],
    "⚠️ تقرير إدارة مخاطر": ["سجل المخاطر", "الاحتمالية والأثر", "خطط الاستجابة", "مسؤوليات التحكم"],
    "🌟 تقرير استراتيجي سنوي": ["الرؤية والمنجز العام", "تحليل الأداء السنوي", "الوضع المالي", "أهداف العام القادم"]
}

# --- القائمة الجانبية (Sidebar) لنظام الرفع والاختيار ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3135/3135715.png", width=80)
    st.markdown("### الإعدادات والوثائق")
    rtype = st.selectbox("🎯 نوع التقرير:", list(STRATEGY_BANK.keys()))
    st.write("---")
    uploaded_file = st.file_uploader("📂 ارفع وثيقة للتحليل (اختياري)", type=['pdf', 'docx', 'txt'])
    if uploaded_file:
        st.success("تم رفع الوثيقة بنجاح")

# --- الواجهة الرئيسية ---
st.markdown('<div class="brand-header"><h1 class="brand-title">منصة المنصور الاستراتيجية AI</h1><p>نظام الصياغة والتحليل المؤسسي المتقدم</p></div>', unsafe_allow_html=True)

# 1. البيانات الأساسية
with st.container():
    st.markdown('<div class="report-card">', unsafe_allow_html=True)
    st.markdown('<span class="q-label">🛡️ أولاً: بيانات الغلاف الرسمي</span>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    p_name = c1.text_input("عنوان التقرير (اسم المشروع) *")
    p_ref = c2.text_input("الرقم المرجعي (Ref No.)")
    p_agency = c1.text_input("الجهة المُعِدّة")
    p_donor = c2.text_input("الجهة المستلمة")
    p_loc = c1.text_input("المكان")
    p_date = c2.text_input("التاريخ", value=datetime.now().strftime('%Y-%m-%d'))
    st.markdown('</div>', unsafe_allow_html=True)

# 2. المقدمة والشكر
with st.expander("🤝 خطاب الإرسال والشكر"):
    p_thanks = st.text_area("أدخل كلمة الشكر أو المقدمة هنا...")

# 3. محاور التقرير (مع زر التحسين)
st.markdown(f"### 🔍 محاور {rtype}")
user_responses = {}
for i, pillar in enumerate(STRATEGY_BANK[rtype]):
    st.markdown('<div class="report-card">', unsafe_allow_html=True)
    st.markdown(f'<span class="q-label">{pillar}</span>', unsafe_allow_html=True)
    txt = st.text_area("أدخل التفاصيل...", key=f"v12_{i}", height=100, label_visibility="collapsed")
    
    c_btn, _ = st.columns([1, 3])
    with c_btn:
        st.markdown('<div class="magic-btn">', unsafe_allow_html=True)
        if st.button(f"✨ تحسين الصياغة", key=f"btn12_{i}"):
            if txt:
                with st.spinner("جاري المعالجة..."):
                    res = model.generate_content(f"صغ هذا المحور بأسلوب استشاري رفيع: {txt}")
                    st.info(res.text)
            else: st.warning("أدخل نصاً")
        st.markdown('</div>', unsafe_allow_html=True)
    user_responses[pillar] = txt
    st.markdown('</div>', unsafe_allow_html=True)

# 4. الخاتمة والاعتماد
with st.container():
    st.markdown('<div class="report-card">', unsafe_allow_html=True)
    st.markdown('<span class="q-label">🖊️ الاعتماد والتوقيعات</span>', unsafe_allow_html=True)
    v1, v2, v3 = st.columns(3)
    p_pre = v1.text_input("إعداد:")
    p_rev = v2.text_input("مراجعة:")
    p_app = v3.text_input("اعتماد:")
    st.markdown('</div>', unsafe_allow_html=True)

# التوليد
if st.button("🚀 إصدار التقرير النهائي"):
    if p_name and any(user_responses.values()):
        with st.spinner("جاري صياغة الوثيقة السيادية..."):
            all_txt = "\n".join([f"{k}: {v}" for k, v in user_responses.items() if v])
            full_prompt = f"صغ تقريراً استراتيجياً لـ {p_name}. الجهة: {p_agency}. المحاور: {all_txt}. التوقيعات: {p_pre}, {p_rev}, {p_app}."
            res = model.generate_content(full_prompt)
            st.markdown(res.text)
            st.session_state['v12_out'] = res.text
    else: st.warning("يرجى تعبئة البيانات.")

if 'v12_out' in st.session_state:
    doc = Document()
    doc.add_heading(p_name, 0)
    doc.add_paragraph(st.session_state['v12_out'])
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    st.download_button("💾 تحميل ملف Word", bio, f"{p_name}.docx")

st.markdown('<center><a href="https://wa.me/967774575749" style="text-decoration:none; color:#25d366; font-weight:bold;">💬 الدعم الفني: 774575749</a></center>', unsafe_allow_html=True)
