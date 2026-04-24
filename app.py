import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document

# 1. التنسيق البصري الفخم (مستوحى من الفيديو)
st.set_page_config(page_title="منصة المنصور AI - التوليد الاستراتيجي", layout="centered")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700;900&display=swap');
    html, body, [class*="st-"] { font-family: 'Cairo', sans-serif !important; direction: rtl; text-align: right; }
    .stApp { background-color: #f4f7f9; }
    #MainMenu, footer, header { visibility: hidden; }
    
    .main-card { background: white; border-top: 10px solid #1e3a8a; padding: 40px; border-radius: 15px; box-shadow: 0 10px 30px rgba(0,0,0,0.05); margin-top: -60px; }
    .brand-title { color: #1e3a8a; font-weight: 900; font-size: 2.2rem; text-align: center; }
    .step-header { background: #f8fafc; padding: 15px; border-right: 5px solid #d4af37; font-weight: 800; color: #1e3a8a; margin-bottom: 20px; border-radius: 0 8px 8px 0; }
    
    .stButton>button { background: linear-gradient(90deg, #1e3a8a, #d4af37) !important; color: white !important; font-weight: 700 !important; height: 55px !important; border-radius: 12px !important; }
    .magic-btn button { background: #fdfaf3 !important; color: #856404 !important; border: 1px dashed #d4af37 !important; height: 35px !important; font-size: 0.8rem !important; }
    
    .whatsapp-fixed { position: fixed; bottom: 20px; left: 20px; background: #25d366; color: white !important; padding: 10px 20px; border-radius: 50px; text-decoration: none; font-weight: bold; z-index: 1000; box-shadow: 0 4px 10px rgba(0,0,0,0.1); }
</style>
""", unsafe_allow_html=True)

# تفعيل Gemini
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")
except:
    st.error("⚠️ يرجى ضبط المفتاح في الإعدادات")

st.markdown('<div class="main-card">', unsafe_allow_html=True)
st.markdown('<h1 class="brand-title">منصة المنصور AI للتقارير</h1>', unsafe_allow_html=True)
st.markdown('<p style="text-align:center; color:#d4af37; font-weight:bold;">نظام الصياغة المتقدم وفق معايير المنظمات الدولية</p>', unsafe_allow_html=True)

# نظام التبويبات (الخطوات) كما في الفيديو
tabs = st.tabs(["1️⃣ المعلومات الأساسية", "2️⃣ التنفيذ والحضور", "3️⃣ الأهداف والإنجازات", "4️⃣ التحديات والدروس"])

data = {}

with tabs[0]:
    st.markdown('<div class="step-header">بيانات المشروع والجهات</div>', unsafe_allow_html=True)
    data['p_name'] = st.text_input("اسم المشروع العام *", placeholder="مثال: برنامج التمكين الرقمي والحوكمة")
    data['p_agency'] = st.text_input("الجهة المنفذة *", placeholder="المؤسسة أو الجهة المسؤولة عن التنفيذ")
    data['p_donor'] = st.text_input("الشريك الممول", placeholder="المنظمات أو الجهات الممولة للمشروع")
    data['p_budget'] = st.text_input("الميزانية الإجمالية (اختياري)", placeholder="مثال: 500,000 دولار أمريكي")

with tabs[1]:
    st.markdown('<div class="step-header">الموقع والتاريخ والحضور</div>', unsafe_allow_html=True)
    data['p_activity'] = st.text_input("عنوان النشاط / الفعالية", placeholder="مثال: ورشة العمل المتقدمة في الإدارة")
    data['p_loc'] = st.text_input("مكان التنفيذ", placeholder="مثال: مسقط - فندق جراند هرمز")
    c1, c2 = st.columns(2)
    data['p_total'] = c1.number_input("إجمالي عدد المشاركين", value=0)
    data['p_males'] = c2.number_input("عدد الذكور", value=0)
    data['p_females'] = st.number_input("عدد الإناث", value=0)

with tabs[2]:
    st.markdown('<div class="step-header">الأهداف والمؤشرات الرئيسية</div>', unsafe_allow_html=True)
    data['p_goals'] = st.text_area("أهداف المشروع الرئيسية", placeholder="1. تعزيز الكفاءات...\n2. إنشاء إطار عمل...")
    data['p_achievements'] = st.text_area("الإنجازات الرئيسية المحققة", placeholder="تأهيل 120 كادراً بنسبة نجاح 95%...")
    data['p_kpis'] = st.text_area("مؤشرات الأداء (KPIs)", placeholder="معدل الحضور: 92%، درجة الرضا: 4.7/5...")

with tabs[3]:
    st.markdown('<div class="step-header">العقبات والإجراءات المتخذة</div>', unsafe_allow_html=True)
    data['p_challenges'] = st.text_area("التحديات التي واجهتكم", placeholder="ضعف تغطية الإنترنت في المناطق البعيدة...")
    data['p_solutions'] = st.text_area("الإجراءات المتخذة للحل", placeholder="توفير مودم خارجي عالي السرعة...")
    data['p_lessons'] = st.text_area("الدروس المستفادة والتوصيات", placeholder="أهمية التنسيق اللوجستي المسبق...")

st.write("---")
if st.button("🚀 إنشاء التقرير الاستراتيجي الآن"):
    if data['p_name'] and data['p_agency']:
        with st.spinner("جاري تحليل البيانات وصياغة الوثيقة النهائية..."):
            context = "\n".join([f"{k}: {v}" for k, v in data.items() if v])
            prompt = f"أنت خبير استشاري دولي. صغ تقريراً استراتيجياً متكاملاً بناءً على هذه البيانات: {context}. استخدم ترقيم ISO 2145، صوت نشط، نبرة قيادية، وأضف ملخصاً تنفيذياً وخاتمة رسمية مع مساحات للتوقيع."
            res = model.generate_content(prompt)
            st.markdown("### 📄 المعاينة النهائية للتقرير:")
            st.markdown(res.text)
            st.session_state['v7_final'] = res.text
    else:
        st.warning("الرجاء إدخال اسم المشروع والجهة المنفذة على الأقل.")

if 'v7_final' in st.session_state:
    doc = Document()
    doc.add_heading(f"Report: {data['p_name']}", 0)
    doc.add_paragraph(st.session_state['v7_final'])
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    st.download_button("💾 تحميل ملف Word المعتمد", bio, f"{data['p_name']}.docx")

# زر الواتساب الثابت
st.markdown(f'<a href="https://wa.me/967774575749" class="whatsapp-fixed">💬 تواصل معنا: 774575749</a>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)
