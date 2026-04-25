import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document
from datetime import datetime

# 1. الهندسة البصرية المتقدمة (منع التداخل نهائياً)
st.set_page_config(page_title="منصة المنصور AI - التميز المؤسسي", layout="centered")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
    
    /* توحيد الخط والاتجاه */
    * { font-family: 'Cairo', sans-serif !important; direction: rtl; text-align: right; }
    .stApp { background-color: #f1f5f9; }
    #MainMenu, footer, header { visibility: hidden; }

    /* حل مشكلة تداخل الرفع (CSS التطهيري) */
    [data-testid="stFileUploader"] { 
        padding: 20px; 
        background-color: #ffffff; 
        border: 2px dashed #1e3a8a; 
        border-radius: 15px;
    }
    /* إخفاء النصوص المتداخلة واستبدالها بنص نظيف */
    [data-testid="stFileUploader"] section > button {
        display: none !important;
    }
    [data-testid="stFileUploader"] section::before {
        content: "📥 اضغط هنا لرفع الملف المرجعي";
        color: #1e3a8a;
        font-weight: bold;
        display: block;
        text-align: center;
        padding: 10px;
        cursor: pointer;
    }

    /* تصميم البطاقات المؤسسية */
    .report-card { 
        background: white; border-top: 10px solid #1e3a8a; 
        padding: 30px; border-radius: 15px; 
        box-shadow: 0 10px 30px rgba(0,0,0,0.05); margin-bottom: 35px;
    }
    
    .brand-title { color: #1e3a8a; font-weight: 900; font-size: 2.3rem; text-align: center; }
    .section-title { background: #1e3a8a; color: white; padding: 12px 20px; border-radius: 8px; font-weight: 700; margin: 25px 0; font-size: 1.1rem; }
    .q-label { color: #1e293b; font-weight: 800; border-right: 5px solid #d4af37; padding-right: 12px; margin-top: 25px; display: block; }
    .hint-box { color: #64748b; font-size: 0.85rem; background: #fffbeb; padding: 12px; border-radius: 8px; border: 1px solid #fef3c7; margin: 10px 0; line-height: 1.6; }

    /* الأزرار الملكية */
    .stButton>button { 
        background: linear-gradient(135deg, #1e3a8a 0%, #152e6d 100%) !important; color: white !important; 
        font-weight: 700 !important; height: 55px !important; border-radius: 10px !important; width: 100%; border: none;
    }
</style>
""", unsafe_allow_html=True)

# تفعيل Gemini
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")
except:
    st.error("⚠️ يرجى التأكد من ضبط مفتاح GEMINI_API_KEY")

# --- بنك المعرفة الشامل (الـ 12 تخصصاً كما تعهدنا) ---
REPORTS_DB = {
    "📑 تقرير إنجاز دوري (إداري)": [
        ("الملخص التنفيذي للأداء", "وصف شامل لمستوى التقدم مقابل الأهداف المخططة..."),
        ("المخرجات والأنشطة المنفذة", "قائمة مفصلة بكل نشاط تم تنفيذه، مكانه، والفئة المستهدفة..."),
        ("تحليل الانحرافات والتحديات", "أسباب أي تأخير ميداني والإجراء التصحيحي المتخذ..."),
        ("خطة العمل القادمة", "تحديد الأولويات القصوى للمرحلة التالية...")
    ],
    "💰 دراسة جدوى استثمارية": [
        ("تحليل السوق والمنافسة", "حجم الطلب، الفجوة السوقية، ونقاط قوة المنافسين..."),
        ("النمذجة المالية والربحية", "تقدير رأس المال، التدفقات النقدية، وفترة استرداد رأس المال..."),
        ("تحليل SWOT والتوصية", "توليف نقاط القوة والفرص لاتخاذ قرار الاستثمار النهائي...")
    ],
    "🎓 تقرير ختامي لبرنامج تدريبي": [
        ("بيانات المدرب والمنهجية", "اسم المدرب، خبرته، والأساليب المستخدمة..."),
        ("إحصائيات الحضور والجندر", "إجمالي المشاركين، عدد الذكور، عدد الإناث..."),
        ("تحليل الأثر المعرفي (Pre/Post)", "مقارنة نتائج الاختبارات قبل التدريب وبعده..."),
        ("توصيات الاستدامة", "مقترحات لضمان تطبيق المهارات الجديدة...")
    ],
    "🔍 متابعة وتقييم (M&E)": [
        ("مؤشرات الأداء KPIs", "تحليل رقمي دقيق للمؤشرات مقابل المستهدفات..."),
        ("رضا المستفيدين والدروس", "خلاصة التغذية الراجعة وأثر التدخل في حياتهم...")
    ],
    "🏗️ هندسي وفني": [("مطابقة المواصفات", "نتائج الفحوصات للمواد..."), ("سير العمل الميداني", "نسبة الإنجاز الفعلي مقابل المخطط...")],
    "🏛️ حوكمة وامتثال": [("الالتزام باللوائح", "تطابق الممارسات مع القوانين..."), ("نتائج التدقيق", "رصد الثغرات وإجراءات التصحيح...")],
    "💰 أداء مالي": [("تحليل الإيرادات والمصروفات", "بيان مالي للتدفقات النقدية..."), ("انحراف الميزانية", "تحليل الفروقات المالية...")],
    "🚑 تقييم احتياجات": [("وصف الاحتياج الراهن", "تحليل الوضع الميداني والفجوة..."), ("الأولويات العاجلة", "التدخلات التي لا تقبل التأجيل...")],
    "🌍 أثر بيئي واجتماعي": [("تحليل الأثر الحيوي", "تأثير المشروع على البيئة..."), ("خطط التخفيف", "إجراءات تقليل الأضرار الجانبية...")],
    "📝 تحليل مناقصات": [("التقييم الفني والمالي", "مقارنة عروض الموردين..."), ("توصية الترسية", "مبررات اختيار المورد الفائز...")],
    "⚠️ إدارة مخاطر": [("سجل المخاطر المحدث", "المخاطر الجديدة وكيفية التعامل معها..."), ("خطط الاستجابة", "فعالية خطط الطوارئ...")],
    "🌟 استراتيجي سنوي": [("إنجازات العام", "حصاد الأهداف الكبرى والمنجز العام..."), ("التوجهات القادمة", "خارطة طريق العام القادم...")]
}

# --- بناء الواجهة ---
st.markdown('<div class="report-card">', unsafe_allow_html=True)
st.markdown('<h1 class="brand-title">منصة المنصور الاستراتيجية AI</h1>', unsafe_allow_html=True)
st.markdown('<p style="text-align:center; color:#c5a059; font-weight:700;">نظام الصياغة والتحليل المؤسسي الشامل - V22</p>', unsafe_allow_html=True)

# 1. الإعداد والرفع (التنسيق الجديد)
st.markdown('<div class="section-title">🎯 الخطوة الأولى: نوع التقرير والمراجع</div>', unsafe_allow_html=True)
rtype = st.selectbox("اختر تخصص التقرير المطلوب:", list(REPORTS_DB.keys()))
up_file = st.file_uploader("", type=['pdf', 'docx', 'txt', 'jpg', 'png'])

# 2. بيانات الغلاف
st.markdown('<div class="section-title">🛡️ الخطوة الثانية: بيانات الغلاف الرسمي</div>', unsafe_allow_html=True)
p_name = st.text_input("اسم المشروع / البرنامج التدريبي *")
p_agency = st.text_input("الجهة المُعِدّة (المؤسسة)")
p_donor = st.text_input("الجهة الموجه إليها")
p_loc = st.text_input("مكان التنفيذ")
p_date = st.text_input("التاريخ", value=datetime.now().strftime('%Y-%m-%d'))

# 3. صلب التقرير
st.markdown(f"### 🔍 الخطوة الثالثة: محاور {rtype}")
user_ans = {}
for i, (pillar, hint) in enumerate(REPORTS_DB[rtype]):
    st.markdown('<div class="report-card">', unsafe_allow_html=True)
    st.markdown(f'<span class="q-label">{pillar}</span>', unsafe_allow_html=True)
    st.markdown(f'<div class="hint-box">📝 {hint}</div>', unsafe_allow_html=True)
    txt = st.text_area("", key=f"v22_{i}", height=120, label_visibility="collapsed")
    
    if st.button(f"✨ تحسين صياغة {pillar}", key=f"btn_v22_{i}"):
        if txt:
            with st.spinner("جاري الصياغة..."):
                res = model.generate_content(f"صغ هذا المحور بأسلوب استشاري رفيع: {txt}")
                st.success(res.text)
        else: st.warning("أدخل نصاً أولاً.")
    user_ans[pillar] = txt
    st.markdown('</div>', unsafe_allow_html=True)

# 4. التخصيص
if 'extra_v22' not in st.session_state: st.session_state.extra_v22 = []
st.markdown('<div class="section-title">➕ الخطوة الرابعة: إضافات مخصصة</div>', unsafe_allow_html=True)
new_sec = st.text_input("اسم القسم الإضافي الجديد:")
if st.button("أنشئ القسم الآن"):
    if new_sec and new_sec not in st.session_state.extra_v22:
        st.session_state.extra_v22.append(new_sec); st.rerun()
for ex in st.session_state.extra_v22:
    user_ans[ex] = st.text_area(f"بيانات {ex}...", key=f"ex22_{ex}")

# 5. التوقيعات
st.markdown('<div class="section-title">🖊️ الخطوة الخامسة: هيكل الاعتماد</div>', unsafe_allow_html=True)
p_pre = st.text_input("إعداد:")
p_rev = st.text_input("مراجعة:")
p_app = st.text_input("اعتماد:")

if st.button("🚀 توليد التقرير الاستراتيجي الشامل"):
    if p_name:
        with st.spinner("جاري التوليد..."):
            summary = "\n".join([f"- {k}: {v}" for k, v in user_ans.items() if v])
            prompt = f"صغ تقريراً استراتيجياً لـ {p_name}. المحاور: {summary}. التوقيعات: {p_pre}, {p_rev}, {p_app}."
            res = model.generate_content(prompt)
            st.markdown(res.text)
            st.session_state['v22_out'] = res.text
    else: st.warning("يرجى ملء اسم المشروع.")

if 'v22_out' in st.session_state:
    doc = Document()
    doc.add_heading(p_name, 0)
    doc.add_paragraph(st.session_state['v22_out'])
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    st.download_button("💾 تحميل ملف Word المعتمد", bio, f"{p_name}.docx")

st.markdown('<center><a href="https://wa.me/967774575749" style="text-decoration:none; color:#25d366; font-weight:bold;">💬 الدعم الفني: 774575749</a></center>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)
