import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document
from datetime import datetime
import PyPDF2

# --- 1. الهندسة البصرية التنفيذية (Premium UX) ---
st.set_page_config(page_title="منصة المنصور الاستراتيجية", layout="centered")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
    
    /* ضبط الاتجاه العام وإجبار المحاذاة لليمين */
    html, body, [data-testid="stAppViewContainer"] {
        direction: rtl;
        text-align: right;
        background: #0f172a; /* خلفية كحلي عميق جداً للتباين */
    }

    * { font-family: 'Cairo', sans-serif !important; }

    /* إجبار نصوص الـ Labels والـ Placeholders على اليمين */
    .stTextInput label, .stTextArea label, .stSelectbox label, .stMultiSelect label {
        text-align: right !important;
        display: block !important;
        color: #C5A059 !important; /* لون ذهبي خافت للعناوين الفرعية */
        font-weight: 700 !important;
        font-size: 0.95rem !important;
    }

    /* تحسين مظهر حقول الإدخال والتباين */
    input, textarea, .stSelectbox div {
        text-align: right !important;
        direction: rtl !important;
        background-color: rgba(30, 41, 59, 0.5) !important;
        color: #ffffff !important;
        border: 1px solid rgba(197, 160, 89, 0.2) !important;
        font-size: 1rem !important;
    }

    /* حاوية العنوان الرئيسي - رصينة ومتزنة */
    .brand-header {
        text-align: center;
        padding: 30px 10px;
        background: linear-gradient(180deg, rgba(30, 41, 59, 0.8) 0%, rgba(15, 23, 42, 0) 100%);
        border-bottom: 2px solid rgba(197, 160, 89, 0.1);
        margin-bottom: 30px;
    }
    .brand-title { color: #C5A059; font-weight: 900; font-size: 2.2rem; margin-bottom: 5px; }
    .brand-subtitle { color: #94a3b8; font-weight: 400; font-size: 0.9rem; }

    /* نظام البطاقات - وضوح تام وتوازن */
    .module-card { 
        background: rgba(30, 41, 59, 0.4); 
        border-right: 4px solid #C5A059; 
        padding: 20px; 
        border-radius: 10px; 
        margin-bottom: 25px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.2);
    }
    
    .module-label { 
        color: #f8fafc; font-weight: 800; font-size: 1.1rem; 
        margin-bottom: 15px; border-bottom: 1px solid rgba(197, 160, 89, 0.2);
        padding-bottom: 10px; display: block;
    }

    /* التلميحات - تباين عالي للقراءة */
    .hint-box {
        background: rgba(197, 160, 89, 0.1);
        border: 1px solid rgba(197, 160, 89, 0.3);
        color: #e2e8f0;
        padding: 12px;
        border-radius: 8px;
        font-size: 0.85rem;
        margin-bottom: 15px;
        line-height: 1.6;
    }

    /* زر التوليد الملكي */
    .stButton>button { 
        background: linear-gradient(90deg, #C5A059 0%, #9a7b41 100%) !important; 
        color: #0f172a !important; font-weight: 900 !important; height: 50px !important; 
        border-radius: 8px !important; border: none !important; width: 100%;
        margin-top: 20px;
    }

    /* إخفاء شعارات Streamlit المزعجة */
    #MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# --- 2. الإعدادات البرمجية ---
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")
except:
    st.error("⚠️ يرجى ضبط مفتاح API في الإعدادات")

# بنك المعرفة (التخصصات مع الأمثلة المدمجة)
REPORTS_BANK = {
    "📑 تقرير إنجاز دوري (إداري)": [
        ("الملخص التنفيذي للأداء", "اكتب ملخصاً عاماً.. مثال: تم تحقيق 90% من المستهدفات الربعية بنجاح."),
        ("تحليل الأنشطة المنفذة", "سرد المنجزات.. مثال: إتمام تدريب الفريق، وتطوير النظام التقني."),
        ("التحديات والمعالجات", "المعوقات.. مثال: واجهنا نقصاً في الموارد وتمت الاستعانة بشريك خارجي.")
    ],
    "🎓 تقرير برنامج تدريبي ختامي": [
        ("المنهجية وبيانات المدرب", "وصف التدريب.. مثال: تم استخدام منهجية Master Coach التفاعلية."),
        ("تحليل مستوى الاستيعاب", "النتائج.. مثال: أظهر التقييم البعدي تحسناً بنسبة 70% في المهارات."),
        ("التوصيات الختامية", "المستقبل.. مثال: ضرورة عقد جلسات تنشيطية كل شهرين.")
    ],
    "🔍 متابعة وتقييم (M&E)": [
        ("مؤشرات الأداء KPI", "الأرقام.. مثال: تحقيق نسبة وصول 100% للفئات المستهدفة."),
        ("الجودة والدروس المستفادة", "التقييم.. مثال: أظهرت الزيارات الميدانية رضى المستفيدين بنسبة 85%.")
    ]
    # يمكن إضافة بقية الـ 12 تخصصاً هنا بنفس النمط
}

# --- 3. الواجهة ---
st.markdown("""
<div class="brand-header">
    <h1 class="brand-title">منصة المنصور الاستراتيجية</h1>
    <p class="brand-subtitle">نظام الصياغة والتحليل المؤسسي الشامل - الإصدار المحدث</p>
</div>
""", unsafe_allow_html=True)

# الخطوة 1
st.markdown('<div class="module-card">', unsafe_allow_html=True)
st.markdown('<span class="module-label">📁 الخطوة 1: المراجع والنوع</span>', unsafe_allow_html=True)
rtype = st.selectbox("حدد تخصص التقرير:", list(REPORTS_BANK.keys()))
up_file = st.file_uploader("ارفق ملفات PDF أو Word (اختياري)", type=['pdf', 'docx', 'txt'])
st.markdown('</div>', unsafe_allow_html=True)

# الخطوة 2
st.markdown('<div class="module-card">', unsafe_allow_html=True)
st.markdown('<span class="module-label">🛡️ الخطوة 2: البيانات الرسمية</span>', unsafe_allow_html=True)
p_name = st.text_input("عنوان المشروع / التقرير *")
col1, col2 = st.columns(2)
with col1:
    p_agency = st.text_input("الجهة المُعِدّة")
    p_loc = st.text_input("مكان التنفيذ")
with col2:
    p_donor = st.text_input("الجهة الموجه إليها")
    p_date = st.text_input("التاريخ", value=datetime.now().strftime('%Y-%m-%d'))
st.markdown('</div>', unsafe_allow_html=True)

# الخطوة 3
st.markdown(f"### 🔍 تفاصيل {rtype}")
user_ans = {}
for i, (pillar, hint) in enumerate(REPORTS_BANK[rtype]):
    st.markdown('<div class="module-card">', unsafe_allow_html=True)
    st.markdown(f'<span class="module-label">{pillar}</span>', unsafe_allow_html=True)
    st.markdown(f'<div class="hint-box">💡 {hint}</div>', unsafe_allow_html=True)
    txt = st.text_area("أدخل البيانات الخام هنا:", key=f"v25_{i}", height=120)
    user_ans[pillar] = txt
    st.markdown('</div>', unsafe_allow_html=True)

# الخطوة 4
st.markdown('<div class="module-card">', unsafe_allow_html=True)
st.markdown('<span class="module-label">🖊️ الخطوة 4: الاعتماد</span>', unsafe_allow_html=True)
c1, c2, c3 = st.columns(3)
with c1: p_pre = st.text_input("إعداد")
with c2: p_rev = st.text_input("مراجعة")
with c3: p_app = st.text_input("اعتماد")

if st.button("🚀 توليد التقرير الاستراتيجي الشامل"):
    if p_name and any(user_ans.values()):
        with st.spinner("جاري المعالجة بصفتي خبير استشاري..."):
            summary = "\n".join([f"محور {k}: {v}" for k, v in user_ans.items() if v])
            prompt = f"""
            بصفتك كبير مستشارين في التخطيط الاستراتيجي، صغ تقريراً احترافياً لـ ({p_name}).
            البيانات: {summary}
            الجهة: {p_agency} | الموجه إليه: {p_donor}
            التواقيع: {p_pre}, {p_rev}, {p_app}
            اجعل اللغة قوية، مهنية، ومنظمة في فقرات واضحة.
            """
            res = model.generate_content(prompt)
            st.session_state['out_v25'] = res.text
            st.markdown("---")
            st.markdown(f'<div style="background: white; color: #1e293b; padding: 30px; border-radius: 8px; border-right: 10px solid #C5A059; line-height: 1.8; text-align: right;">{res.text}</div>', unsafe_allow_html=True)
    else:
        st.warning("يرجى ملء البيانات الأساسية.")
st.markdown('</div>', unsafe_allow_html=True)

st.markdown(f'<center><a href="https://wa.me/967774575749" style="text-decoration:none; color:#C5A059; font-weight:900;">💬 دعم المنصور الفني: 774575749</a></center>', unsafe_allow_html=True)
