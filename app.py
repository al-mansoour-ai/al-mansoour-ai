import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document

# 1. التنسيق والجماليات
st.set_page_config(page_title="منصة المنصور AI - التخصيص الكامل", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700;900&display=swap');
    html, body, [class*="st-"] { font-family: 'Cairo', sans-serif; direction: rtl; text-align: right; }
    .stApp { background-color: #f8fafc; }
    .main-box { background: white; border-top: 10px solid #1e3a8a; padding: 40px; border-radius: 20px; box-shadow: 0 15px 35px rgba(0,0,0,0.1); margin-top: -50px; }
    .brand-title { color: #1e3a8a; font-weight: 900; font-size: 2.2rem; text-align: center; }
    .custom-section { background: #fffbeb; border: 1px dashed #d4af37; padding: 15px; border-radius: 10px; margin-top: 10px; }
    .stButton>button { background: linear-gradient(90deg, #1e3a8a, #d4af37) !important; color: white !important; font-weight: bold !important; border-radius: 12px !important; }
    .add-btn button { background: #10b981 !important; border: none !important; color: white !important; }
</style>
""", unsafe_allow_html=True)

# 2. قاعدة بيانات التقارير الأساسية
REPORT_DATABASE = {
    "📑 تقرير إنجاز إداري": ["الملخص التنفيذي", "الأنشطة المنفذة", "تحليل التحديات", "الخطوات القادمة"],
    "🎓 تقرير برنامج تدريبي": ["الأهداف التدريبية", "تحليل نتائج المتدربين", "تقييم المحتوى والمدرب", "توصيات الاستدامة"],
    "🔬 دراسة جدوى": ["فكرة المشروع", "الدراسة السوقية", "المتطلبات والتقديرات", "تحليل المخاطر"],
    "📊 تقرير متابعة وتقييم (M&E)": ["مؤشرات الأداء المحققة", "جودة المخرجات", "تغذية راجعة من المستفيدين", "الدروس المستفادة"]
}

# 3. واجهة المستخدم
st.markdown('<div class="main-box">', unsafe_allow_html=True)
st.markdown('<h1 class="brand-title">منصة المنصور AI للتقارير</h1>', unsafe_allow_html=True)

# تفعيل Gemini
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")
except:
    st.error("⚠️ يرجى ضبط المفتاح في Secrets")

# اختيار النوع الأساسي
report_type = st.selectbox("🎯 اختر نوع التقرير الأساسي:", list(REPORT_DATABASE.keys()))
p_name = st.text_input("اسم المشروع / الجهة المنفذة *")

st.write("---")
st.write("### 📝 محاور التقرير (جاهزة ومخصصة)")

# عرض الأسئلة الأساسية بناءً على النوع المختار
user_data = {}
for pillar in REPORT_DATABASE[report_type]:
    st.markdown(f"**🔹 {pillar}**")
    user_data[pillar] = st.text_area(f"أدخل بيانات {pillar}...", key=pillar, height=100)

# --- ميزة التخصيص (إضافة أقسام جديدة) ---
st.write("---")
st.markdown("### ➕ هل تريد إضافة أقسام إضافية خاصة؟")
if 'custom_sections' not in st.session_state:
    st.session_state['custom_sections'] = []

# خانة لإضافة اسم القسم الجديد
new_section_name = st.text_input("اكتب اسم القسم الجديد هنا (مثلاً: المرفقات، الميزانية التفصيلية، إحصائيات الصور):")
if st.button("➕ إضافة هذا القسم للفورم"):
    if new_section_name and new_section_name not in st.session_state['custom_sections']:
        st.session_state['custom_sections'].append(new_section_name)
        st.rerun()

# عرض الأقسام التي أضافها العميل
for custom in st.session_state['custom_sections']:
    st.markdown(f'<div class="custom-section">', unsafe_allow_html=True)
    st.markdown(f"**⭐ قسم مخصص: {custom}**")
    user_data[custom] = st.text_area(f"أدخل تفاصيل {custom}...", key=f"custom_{custom}", height=100)
    if st.button(f"🗑️ حذف قسم {custom}"):
        st.session_state['custom_sections'].remove(custom)
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# --- التوليد النهائي ---
st.write("---")
if st.button("🚀 صياغة التقرير الشامل (بما في ذلك الأقسام المخصصة)"):
    if p_name and any(user_data.values()):
        with st.spinner("جاري صهر البيانات وصياغتها استراتيجياً..."):
            # تحويل البيانات لنص مفهوم للذكاء الاصطناعي
            data_string = "\n".join([f"القسم ({k}): {v}" for k, v in user_data.items() if v])
            
            prompt = f"""
            بصفتك مستشاراً عالمياً، صغ تقريراً من نوع {report_type} للمشروع {p_name}.
            يجب أن يتضمن التقرير كافة الأقسام التالية ويعالجها باحترافية:
            {data_string}
            
            المعايير المطلوبة:
            - لغة قيادية، إيجاز، وصوت نشط.
            - ترقيم دولي ISO 2145.
            - دمج الأقسام المخصصة بسلاسة في سياق التقرير.
            """
            
            response = model.generate_content(prompt)
            st.markdown("### 📄 التقرير النهائي:")
            st.info(response.text)
            st.session_state['report_final'] = response.text
    else:
        st.warning("يرجى تعبئة البيانات.")

# التصدير لـ Word
if 'report_final' in st.session_state:
    doc = Document()
    doc.add_heading(f"تقرير: {p_name}", 0)
    doc.add_paragraph(st.session_state['report_final'])
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    st.download_button("💾 تحميل ملف Word المعتمد", bio, f"{p_name}.docx")

st.markdown('</div>', unsafe_allow_html=True)
st.markdown("<center style='color:#94a3b8; font-size:0.7rem; margin-top:20px;'>🛡️ شبكة المنصور الدولية للاستشارات | 2026</center>", unsafe_allow_html=True)
