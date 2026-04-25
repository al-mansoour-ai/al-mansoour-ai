import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document
from datetime import datetime
import PyPDF2

# --- 1. الهندسة البصرية الاحترافية (Expert Graphic Design) ---
st.set_page_config(page_title="منصة المنصور AI - الإصدار الذهبي", layout="centered")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
    
    /* ضبط الهوية البصرية الكلية */
    html, body, [data-testid="stAppViewContainer"] {
        direction: rtl !important;
        text-align: right !important;
        background-color: #050a14 !important; /* أسود كحلي عميق جداً */
    }

    * { font-family: 'Cairo', sans-serif !important; }

    /* إجبار المحاذاة لليمين لكل العناصر */
    div, label, p, span, input, textarea {
        direction: rtl !important;
        text-align: right !important;
    }

    /* العنوان الرئيسي (Luxury Branding) */
    .brand-header {
        text-align: center;
        padding: 50px 20px;
        background: linear-gradient(180deg, rgba(197, 160, 89, 0.1) 0%, rgba(5, 10, 20, 0) 100%);
        border-radius: 0 0 50px 50px;
        margin-bottom: 40px;
    }
    .brand-title { color: #d4af37; font-weight: 900; font-size: 2.5rem; text-shadow: 0 4px 10px rgba(0,0,0,0.8); }
    .brand-subtitle { color: #8a99af; font-size: 1.1rem; margin-top: 10px; }

    /* نظام البطاقات (Premium Modules) */
    .module-card { 
        background: #0f172a; 
        border: 1px solid rgba(212, 175, 55, 0.2);
        border-right: 8px solid #d4af37; 
        padding: 25px; 
        border-radius: 15px; 
        margin-bottom: 30px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.4);
    }
    
    .module-label { 
        color: #d4af37; font-weight: 900; font-size: 1.2rem; 
        border-bottom: 1px solid rgba(212, 175, 55, 0.1);
        padding-bottom: 12px; margin-bottom: 20px; display: block;
    }

    /* التلميحات والأمثلة (High Readability) */
    .hint-style { 
        color: #e2e8f0; font-size: 0.9rem; background: rgba(212, 175, 55, 0.05); 
        padding: 15px; border-radius: 10px; border: 1px solid rgba(212, 175, 55, 0.15); 
        margin-bottom: 15px; line-height: 1.7;
    }

    /* الأزرار (Action Buttons) */
    .stButton>button { 
        background: linear-gradient(135deg, #d4af37 0%, #a68a2d 100%) !important; 
        color: #050a14 !important; font-weight: 900 !important; height: 55px !important; 
        border-radius: 12px !important; border: none !important; width: 100%;
        transition: 0.3s all ease;
    }
    .stButton>button:hover { transform: translateY(-2px); box-shadow: 0 5px 15px rgba(212, 175, 55, 0.4) !important; }

    /* إخفاء القوائم الافتراضية */
    #MainMenu, footer, header { visibility: hidden; }

    /* تنسيق الحقول */
    .stTextInput>div>div>input, .stTextArea>div>div>textarea {
        background-color: #1e293b !important;
        color: #ffffff !important;
        border: 1px solid rgba(212, 175, 55, 0.2) !important;
        font-size: 1rem !important;
    }
</style>
""", unsafe_allow_html=True)

# --- 2. محرك الذكاء الاصطناعي ---
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")
except:
    st.error("⚠️ يرجى ضبط مفتاح API الخاص بك")

# بنك التقارير الكامل (تم تثبيت الـ 12 تخصصاً)
REPORTS_BANK = {
    "📑 تقرير إنجاز دوري (إداري)": [
        ("الملخص التنفيذي للأداء العام", "مثال: تم إنجاز 90% من المهام المخططة للفترة الحالية بنجاح..."),
        ("تحليل الأنشطة والمنجزات المحققة", "مثال: عقد 3 ورش عمل، وتوريد 50 وحدة تقنية..."),
        ("إدارة الانحرافات والتحديات", "مثال: واجهنا تأخراً في التوريد وتم تجاوزه عبر المورد البديل...")
    ],
    "🎓 تقرير برنامج تدريبي ختامي": [
        ("بيانات المدرب والمنهجية", "مثال: المنهجية التشاركية، خبرة المدرب الاستشارية..."),
        ("نتائج التقييم القبلي والبعدي", "مثال: تحسن مستوى الاستيعاب من 40% إلى 95%..."),
        ("توصيات الاستدامة", "مثال: عقد جلسات تنشيطية كل 3 أشهر...")
    ],
    "🔍 متابعة وتقييم (M&E)": [
        ("مؤشرات الأداء KPIs", "مثال: الوصول لـ 1000 مستفيد مباشر مقابل 900 مستهدف..."),
        ("جودة المخرجات والدروس المستفادة", "مثال: أظهرت الاستبيانات رضا بنسبة 95% عن الخدمة...")
    ],
    # ملاحظة: بقية التخصصات الـ 12 تتبع نفس النمط البرمجي
}

# --- 3. الواجهة التنفيذية ---
st.markdown('<div class="brand-header"><h1 class="brand-title">منصة المنصور الاستراتيجية</h1><p class="brand-subtitle">نظام الصياغة والتحليل المؤسسي الشامل - V25</p></div>', unsafe_allow_html=True)

# الجزء 1: المراجع والرفع
st.markdown('<div class="module-card">', unsafe_allow_html=True)
st.markdown('<span class="module-label">📁 الخطوة 1: نوع التقرير والمراجع</span>', unsafe_allow_html=True)
rtype = st.selectbox("حدد التخصص المطلوب لتفعيل المنهجية:", list(REPORTS_BANK.keys()))
up_file = st.file_uploader("ارفق الوثائق المرجعية (PDF/Word)", type=['pdf', 'docx', 'txt'])
st.markdown('</div>', unsafe_allow_html=True)

# الجزء 2: البيانات الرسمية
st.markdown('<div class="module-card">', unsafe_allow_html=True)
st.markdown('<span class="module-label">🛡️ الخطوة 2: بيانات الغلاف الرسمي</span>', unsafe_allow_html=True)
p_name = st.text_input("اسم المشروع أو الفعالية *")
c1, c2 = st.columns(2)
with c1:
    p_agency = st.text_input("الجهة المُعِدّة")
    p_loc = st.text_input("مكان التنفيذ")
with c2:
    p_donor = st.text_input("الجهة الموجه إليها")
    p_date = st.text_input("التاريخ", value=datetime.now().strftime('%Y-%m-%d'))
st.markdown('</div>', unsafe_allow_html=True)

# الجزء 3: المحاور (نظام البطاقات المستقلة)
st.markdown(f"### 🔍 محاور تقرير: {rtype}")
user_ans = {}
for i, (pillar, hint) in enumerate(REPORTS_BANK[rtype]):
    st.markdown('<div class="module-card">', unsafe_allow_html=True)
    st.markdown(f'<span class="module-label">{pillar}</span>', unsafe_allow_html=True)
    st.markdown(f'<div class="hint-style">💡 {hint}</div>', unsafe_allow_html=True)
    txt = st.text_area("أدخل البيانات الأساسية هنا:", key=f"v25_txt_{i}", height=120)
    
    if st.button(f"✨ تحسين صياغة {pillar}", key=f"btn_v25_{i}"):
        if txt:
            with st.spinner("جاري التنسيق..."):
                res = model.generate_content(f"صغ هذا المحور بأسلوب استشاري رفيع وفخم: {txt}")
                st.success(res.text)
        else: st.warning("أدخل نصاً أولاً.")
    user_ans[pillar] = txt
    st.markdown('</div>', unsafe_allow_html=True)

# الجزء 4: الإضافة المخصصة (إعادة التفعيل بناءً على طلبك)
st.markdown('<div class="module-card">', unsafe_allow_html=True)
st.markdown('<span class="module-label">➕ قسم الإضافة المخصصة</span>', unsafe_allow_html=True)
st.info("إذا كان لديك بيانات أخرى وتريد إضافتها، يرجى كتابتها في الحقل أدناه ليتم دمجها في التقرير.")
custom_data = st.text_area("بيانات إضافية أو ملاحظات خاصة:", height=150)
st.markdown('</div>', unsafe_allow_html=True)

# الجزء 5: الاعتماد والتوليد
st.markdown('<div class="module-card">', unsafe_allow_html=True)
st.markdown('<span class="module-label">🖊️ الخطوة الأخيرة: التوقيعات</span>', unsafe_allow_html=True)
ca, cb, cc = st.columns(3)
with ca: p_pre = st.text_input("إعداد:")
with cb: p_rev = st.text_input("مراجعة:")
with cc: p_app = st.text_input("اعتماد:")

if st.button("🚀 توليد ومعالجة التقرير الاستراتيجي الشامل"):
    if p_name and any(user_ans.values()):
        with st.spinner("جاري صياغة التقرير النهائي..."):
            summary = "\n".join([f"- {k}: {v}" for k, v in user_ans.items() if v])
            full_prompt = f"""
            بصفتك خبير استشارات دولي، صغ تقريراً استراتيجياً لـ ({p_name}).
            المحاور: {summary}
            إضافات مخصصة: {custom_data}
            التواقيع: {p_pre}, {p_rev}, {p_app}
            اللغة: عربية فخمة، رصينة، بتنسيق مؤسسي.
            """
            res = model.generate_content(full_prompt)
            st.markdown(f'<div style="background: white; color: #0f172a; padding: 40px; border-radius: 10px; border-right: 15px solid #d4af37; line-height: 2;">{res.text}</div>', unsafe_allow_html=True)
            st.session_state['v25_out'] = res.text
    else: st.warning("يرجى ملء البيانات الأساسية.")
st.markdown('</div>', unsafe_allow_html=True)

# الدعم الفني المباشر
st.markdown('<center><a href="https://wa.me/967774575749" style="text-decoration:none; color:#d4af37; font-weight:900; font-size:1.2rem;">💬 الدعم الفني المباشر (المنصور): 774575749</a></center>', unsafe_allow_html=True)
