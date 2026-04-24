import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document
from datetime import datetime

# 1. الهندسة البصرية الفخمة (Premium UI/UX)
st.set_page_config(page_title="منصة المنصور AI - النظام المتكامل", layout="centered")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
    
    /* توحيد الخط والاتجاه */
    * { font-family: 'Cairo', sans-serif !important; direction: rtl; text-align: right; }
    .stApp { background-color: #f1f5f9; }
    #MainMenu, footer, header { visibility: hidden; }

    /* البطاقة الرئيسية */
    .main-card { 
        background: white; border-top: 10px solid #1e3a8a; 
        padding: 35px; border-radius: 15px; 
        box-shadow: 0 15px 35px rgba(0,0,0,0.05); margin-top: -50px; 
    }
    
    .brand-title { color: #1e3a8a; font-weight: 900; font-size: 2.2rem; text-align: center; margin: 0; }
    .brand-subtitle { color: #c5a059; text-align: center; font-weight: 700; font-size: 0.9rem; margin-bottom: 30px; }

    /* العناوين الفرعية الرشيقة */
    .section-header { 
        background: #1e3a8a; color: white; padding: 10px 20px; 
        border-radius: 8px; font-weight: 700; font-size: 1.1rem; 
        margin: 25px 0 15px 0; display: inline-block;
    }
    
    .q-label { color: #1e293b; font-weight: 800; border-right: 5px solid #c5a059; padding-right: 12px; margin-top: 20px; display: block; }
    .hint-box { color: #64748b; font-size: 0.8rem; background: #f8fafc; padding: 10px; border-radius: 6px; border: 1px solid #e2e8f0; margin-bottom: 10px; }

    /* الأزرار الاحترافية */
    .stButton>button { 
        background: linear-gradient(135deg, #1e3a8a 0%, #152e6d 100%) !important; color: white !important; 
        font-weight: 700 !important; height: 55px !important; border-radius: 12px !important; 
        border: none !important; width: 100%; transition: 0.3s;
    }
    
    .magic-btn button { 
        background: #fdfaf3 !important; color: #856404 !important; border: 1px dashed #c5a059 !important; 
        height: 35px !important; font-size: 0.8rem !important; margin-top: 5px; width: auto !important;
    }

    .whatsapp-btn { background: #25d366; color: white !important; padding: 12px 25px; border-radius: 50px; text-decoration: none; font-weight: 700; display: inline-flex; align-items: center; gap: 10px; margin-top: 20px; }
</style>
""", unsafe_allow_html=True)

# تفعيل Gemini
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")
except:
    st.error("⚠️ يرجى ضبط المفتاح في الإعدادات")

# بنك التقارير (الـ 12 تخصصاً كاملاً)
STRATEGY_BANK = {
    "📑 تقرير إنجاز دوري": ["الملخص التنفيذي للأداء", "تحليل الأنشطة والمنجزات", "إدارة الانحرافات", "التحديات والحلول", "الأولويات القادمة"],
    "💰 دراسة جدوى": ["الاحتياج السوقي", "المتطلبات الفنية", "النمذجة المالية", "تحليل المخاطر", "توصية الاستثمار"],
    "🎓 تقرير ختامي لتدريب": ["أهداف المنهجية", "نتائج القبلي والبعدي", "تقييم المدرب واللوجستيات", "توصيات الاستدامة"],
    "🔍 تقرير متابعة وتقييم (M&E)": ["قياس KPIs", "جودة المخرجات", "رضا المستفيدين", "الدروس المستفادة"],
    "🏛️ حوكمة وامتثال": ["الالتزام باللوائح", "نتائج التدقيق", "الثغرات المرصودة", "خطة التحسين"],
    "🏗️ تقرير فني وهندسي": ["المواصفات والمواد", "اختبارات الجودة", "المعوقات والحلول", "السلامة المهنية"]
    # (يمكنك إضافة البقية بنفس الطريقة)
}

st.markdown('<div class="main-card">', unsafe_allow_html=True)
st.markdown('<h1 class="brand-title">منصة المنصور الاستراتيجية AI</h1>', unsafe_allow_html=True)
st.markdown('<p class="brand-subtitle">PREMIUM EXECUTIVE REPORTING SYSTEM V10</p>', unsafe_allow_html=True)

# الخطوة 1: الإعداد والرفع
st.markdown('<div class="section-header">📁 الخطوة 1: الإعداد ورفع الوثائق</div>', unsafe_allow_html=True)
rtype = st.selectbox("🎯 حدد تخصص التقرير:", list(STRATEGY_BANK.keys()))
uploaded_file = st.file_uploader("هل لديك مسودة أو وثيقة تريد من الذكاء الاصطناعي تحليلها؟ (اختياري)", type=['pdf', 'docx', 'txt'])

# الخطوة 2: البيانات التعريفية
st.markdown('<div class="section-header">🛡️ الخطوة 2: الغلاف والبيانات الرسمية</div>', unsafe_allow_html=True)
p_name = st.text_input("اسم المشروع / النشاط *")
p_agency = st.text_input("الجهة المُعِدّة")
p_donor = st.text_input("الجهة الموجه إليها")
p_loc = st.text_input("المكان والنطاق الجغرافي")

# الخطوة 3: المحاور (مع زر التحسين)
st.markdown('<div class="section-header">🔍 الخطوة 3: صلب التقرير والتحليل</div>', unsafe_allow_html=True)
user_responses = {}
for i, pillar in enumerate(STRATEGY_BANK[rtype]):
    st.markdown(f'<span class="q-label">{pillar}</span>', unsafe_allow_html=True)
    st.markdown(f'<div class="hint-box">💡 مثال: ابدأ بذكر أهم النتائج المحققة في {pillar} بشكل موجز...</div>', unsafe_allow_html=True)
    txt = st.text_area("", key=f"v10_{i}", height=120, label_visibility="collapsed")
    
    if st.button(f"✨ تحسين {pillar}", key=f"btn10_{i}"):
        if txt:
            with st.spinner("جاري الصياغة..."):
                res = model.generate_content(f"صغ هذا المحور بأسلوب استشاري رفيع: {txt}")
                st.code(res.text)
        else: st.warning("أدخل نصاً")
    user_responses[pillar] = txt

# الخطوة 4: التخصيص
st.markdown('<div class="section-header">➕ الخطوة 4: إضافات مخصصة</div>', unsafe_allow_html=True)
st.info("إذا كان لديك بيانات أخرى وتريد إضافتها، اضغط على الزر وخصص قسماً جديداً:")
if 'v10_extra' not in st.session_state: st.session_state.v10_extra = []
new_sec = st.text_input("اسم القسم الجديد:")
if st.button("خصص قسم الآن"):
    if new_sec: st.session_state.v10_extra.append(new_sec); st.rerun()

for ex in st.session_state.v10_extra:
    user_responses[ex] = st.text_area(f"بيانات {ex}...", key=f"ex10_{ex}")

# الخطوة 5: التوقيعات
st.markdown('<div class="section-header">🖊️ الخطوة 5: الاعتماد والتوقيع</div>', unsafe_allow_html=True)
p_pre = st.text_input("أعده:")
p_rev = st.text_input("راجعه:")
p_app = st.text_input("اعتمده:")

st.write("---")
if st.button("🚀 معالجة وتوليد التقرير الاستراتيجي الشامل"):
    if p_name and any(user_responses.values()):
        with st.spinner("جاري صهر البيانات وفق معايير الجودة العالمية..."):
            summary = "\n".join([f"- {k}: {v}" for k, v in user_responses.items() if v])
            prompt = f"أنت مستشار دولي. صغ تقريراً استراتيجياً متكاملاً. المشروع: {p_name}. الجهة: {p_agency}. الموجه لـ: {p_donor}. المكان: {p_loc}. المعطيات: {summary}. التوقيعات: {p_pre}، {p_rev}، {p_app}."
            res = model.generate_content(prompt)
            st.markdown(res.text)
            st.session_state['v10_out'] = res.text
    else: st.warning("يرجى ملء البيانات.")

if 'v10_out' in st.session_state:
    doc = Document()
    doc.add_heading(f"Report: {p_name}", 0)
    doc.add_paragraph(st.session_state['v10_out'])
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    st.download_button("💾 تحميل الوثيقة المعتمدة (Word)", bio, f"{p_name}.docx")

st.markdown('<center><a href="https://wa.me/967774575749" class="whatsapp-btn">💬 تواصل معنا: 774575749</a></center>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)
