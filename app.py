import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document

# 1. الهوية البصرية الرسمية (Premium Corporate Style)
st.set_page_config(page_title="منصة المنصور AI - الإصدار السيادي", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
    html, body, [class*="st-"] { font-family: 'Cairo', sans-serif; direction: rtl; text-align: right; }
    
    /* الخلفية والألوان الرسمية */
    .stApp { background-color: #f4f7fa; }
    #MainMenu, footer, header { visibility: hidden; }
    
    .main-box { 
        background: white; border-top: 8px solid #1e3a8a; 
        padding: 40px; border-radius: 12px; 
        box-shadow: 0 4px 20px rgba(0,0,0,0.08); margin-top: -60px; 
    }
    
    .brand-title { color: #1e3a8a; font-weight: 900; font-size: 2.5rem; text-align: center; }
    .tagline { color: #d4af37; text-align: center; font-weight: 600; margin-bottom: 30px; letter-spacing: 1px; }
    
    /* تنسيق الأقسام والأسئلة */
    .section-header { 
        background: #1e3a8a; color: white; padding: 12px 20px; 
        border-radius: 8px; font-weight: bold; margin: 25px 0 15px 0;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.1);
    }
    
    .q-label { color: #334155; font-weight: 700; margin-bottom: 8px; display: block; border-right: 4px solid #d4af37; padding-right: 10px; }
    
    /* الأزرار الرسمية */
    .stButton>button { 
        background: #1e3a8a !important; color: #ffffff !important; 
        font-weight: bold !important; height: 55px !important; 
        border-radius: 8px !important; border: none !important; width: 100%;
        transition: 0.3s;
    }
    .stButton>button:hover { background: #152e6d !important; border-bottom: 4px solid #d4af37 !important; }
    
    .custom-section { background: #f8fafc; border: 1px dashed #cbd5e1; padding: 15px; border-radius: 8px; margin-top: 15px; }
</style>
""", unsafe_allow_html=True)

# 2. بنك الأسئلة الاستراتيجية العميق (8 تخصصات عالمية)
STRATEGIC_BANK = {
    "📊 تقرير إنجاز دوري (إدارة عليا)": [
        "الملخص التنفيذي للأداء العام", "مؤشرات الإنجاز مقابل الجدول الزمني", 
        "تحليل الموارد المستهلكة (بشرية/مالية)", "العوائق التي أثرت على سرعة التنفيذ",
        "الإجراءات التصحيحية المتخذة", "خطة العمل للأسبوع/الشهر القادم"
    ],
    "🏗️ تقرير فني وهندسي (Technical)": [
        "مطابقة المواصفات الفنية والمواد", "نتائج اختبارات الجودة الميدانية", 
        "تقرير السلامة المهنية والموقع", "المعوقات الإنشائية والحلول الهندسية",
        "نسبة الإنجاز المادي مقابل المالي", "ملاحظات الاستشاري الفني"
    ],
    "💰 دراسة جدوى اقتصادية (Feasibility)": [
        "تحليل الفجوة السوقية والاحتياج", "الدراسة الفنية (الآلات/العمالة/المكان)", 
        "النموذج المالي وتوقعات الإيرادات", "حساب نقطة التعادل (Break-even Point)",
        "تحليل المخاطر (SWOT Analysis)", "الاستنتاج النهائي وقرار الاستثمار"
    ],
    "🎓 تقرير ختامي لبرنامج تدريبي": [
        "الأهداف المعرفية والمهارية للبرنامج", "تحليل التقييم القبلي والبعدي للمشاركين", 
        "تقييم كفاءة المدرب والمادة العلمية", "تفاعل المتدربين والبيئة اللوجستية",
        "أبرز قصص النجاح أو التغيير المرصودة", "توصيات استدامة الأثر المهني"
    ],
    "🔍 تقرير متابعة وتقييم (M&E)": [
        "مصفوفة النتائج والمؤشرات المحققة", "جودة المخرجات ومدى مطابقتها للمعايير", 
        "تحليل التغذية الراجعة من المستفيدين", "الدروس المستفادة والفرص الضائعة",
        "الكفاءة في استخدام الموارد المتاحة", "توصيات استراتيجية للمراحل القادمة"
    ],
    "🏛️ تقرير الحوكمة والامتثال (Compliance)": [
        "مدى الالتزام باللوائح والسياسات الداخلية", "نتائج الرقابة والتدقيق الدوري", 
        "الثغرات المرصودة في نظام الرقابة", "حالات عدم الامتثال والإجراءات المتخذة",
        "مقترحات تطوير الهيكل التنظيمي", "خطة تحسين مستوى الشفافية"
    ],
    "🚑 تقرير تقييم الاحتياجات (Needs Assessment)": [
        "وصف دقيق للأزمة أو المشكلة الراهنة", "تحديد الفئات الأكثر تضرراً (ديموغرافياً)", 
        "تحليل الموارد المتاحة والفجوة القائمة", "الأولويات العاجلة للتدخل الإنساني/التنموي",
        "العوائق المحتملة لعمليات الاستجابة", "خارطة طريق مقترحة للتمويل والتدخل"
    ],
    "💰 تقرير الأداء المالي (Financial)": [
        "بيان الدخل والمصروفات الفعلية", "تحليل الانحرافات المالية عن الميزانية", 
        "حالة التدفق النقدي والسيولة", "الالتزامات والديون (إن وجدت)",
        "التوصيات لتقليل الهدر المالي", "التنبؤات المالية للفترة القادمة"
    ]
}

# 3. بناء الواجهة
st.markdown('<div class="main-box">', unsafe_allow_html=True)
st.markdown('<h1 class="brand-title">منصة المنصور الاستراتيجية AI</h1>', unsafe_allow_html=True)
st.markdown('<p class="tagline">PREMIUM EXECUTIVE REPORTING SYSTEM</p>', unsafe_allow_html=True)

# تفعيل Gemini
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")
except:
    st.error("⚠️ يرجى ضبط المفتاح في الإعدادات")

# اختيار التخصص
report_type = st.selectbox("🎯 اختر تخصص التقرير (ستتغير الأسئلة بناءً عليه):", list(STRATEGIC_BANK.keys()))

st.markdown('<div class="section-header">📋 أولاً: المعلومات العامة والتعريفية</div>', unsafe_allow_html=True)
col_a, col_b = st.columns(2)
p_name = col_a.text_input("اسم المشروع / المهمة *")
p_agency = col_b.text_input("الجهة المنفذة / العميل")
p_target = col_a.text_input("الجهة الموجه إليها التقرير")
p_loc = col_b.text_input("مكان التنفيذ / التاريخ")

# عرض الأسئلة المتغيرة
st.markdown(f'<div class="section-header">🔍 ثانياً: المحاور الاستراتيجية لـ {report_type}</div>', unsafe_allow_html=True)
user_responses = {}
for pillar in STRATEGIC_BANK[report_type]:
    st.markdown(f'<span class="q-label">{pillar}</span>', unsafe_allow_html=True)
    user_responses[pillar] = st.text_area(f"أدخل البيانات الخاصة بـ {pillar}...", key=pillar, height=120)

# ميزة التخصيص
if 'extra' not in st.session_state: st.session_state.extra = []
st.markdown('<div class="section-header">➕ ثالثاً: أقسام إضافية مخصصة (اختياري)</div>', unsafe_allow_html=True)
new_sec = st.text_input("أضف قسماً خاصاً بك (مثال: الصور الميدانية، الملحق القانوني):")
if st.button("إضافة هذا القسم للفورم"):
    if new_sec: st.session_state.extra.append(new_sec); st.rerun()

for ex in st.session_state.extra:
    st.markdown(f'<div class="custom-section"><b>⭐ قسم مخصص: {ex}</b>', unsafe_allow_html=True)
    user_responses[ex] = st.text_area(f"أدخل تفاصيل {ex}...", key=f"ex_{ex}")
    if st.button(f"حذف {ex}"): st.session_state.extra.remove(ex); st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# التوليد
st.write("---")
if st.button("🚀 توليد التقرير الاستراتيجي الشامل"):
    if p_name and any(user_responses.values()):
        with st.spinner("جاري صياغة التقرير بنبرة قيادية (Executive Tone)..."):
            data_summary = "\n".join([f"- {k}: {v}" for k, v in user_responses.items() if v])
            prompt = f"""
            بصفتك مستشاراً دولياً، صغ تقريراً من نوع {report_type} للمشروع {p_name}.
            المعلومات: الجهة {p_agency}، المكان {p_loc}، الموجه لـ {p_target}.
            المعطيات التفصيلية:
            {data_summary}
            
            المعايير:
            - استخدم لغة عربية فصحى رفيعة، صوت نشط، وإيجاز استراتيجي.
            - اتبع ترقيم ISO 2145 (1. ثم 1.1).
            - صغ ملخصاً تنفيذياً في البداية وتوصيات في النهاية.
            """
            response = model.generate_content(prompt)
            st.markdown("### 📄 التقرير النهائي:")
            st.success("تم التوليد بنجاح.")
            st.markdown(response.text)
            st.session_state['final'] = response.text
    else: st.warning("يرجى ملء البيانات.")

# تحميل Word
if 'final' in st.session_state:
    doc = Document()
    doc.add_heading(f"تقرير: {p_name}", 0)
    doc.add_paragraph(st.session_state['final'])
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    st.download_button("💾 تحميل ملف Word المعتمد", bio, f"{p_name}.docx")

st.markdown('</div>', unsafe_allow_html=True)
st.markdown("<center style='color:#64748b; font-size:0.7rem; margin-top:20px;'>🛡️ شبكة المنصور الدولية للاستشارات | إدارة التميز المؤسسي | 2026</center>", unsafe_allow_html=True)
