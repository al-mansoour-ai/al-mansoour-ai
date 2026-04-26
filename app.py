import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import datetime

# 1. المعمارية البصرية السيادية
st.set_page_config(page_title="المنصور الاستراتيجية", layout="wide")
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700&display=swap');
    #MainMenu, footer, header {visibility: hidden;}
    .stApp { background-color: #0c0c0c !important; }
    h1, h2, h3, h4, p, span, div, label, li { font-family: 'Cairo', sans-serif !important; text-align: right !important; direction: rtl !important; color: #ffffff; }
    h1, h2, h3 { color: #D4AF37 !important; }
    input, textarea, div[role="listbox"], .stSelectbox > div { background-color: #1a1a1a !important; border: 1px solid #D4AF37 !important; color: #ffffff !important; text-align: right !important; }
    .stButton > button { background-color: #D4AF37 !important; color: #0c0c0c !important; font-weight: 700 !important; width: 100% !important; }
</style>
""", unsafe_allow_html=True)

# 2. الهياكل المنهجية العالمية (الخمسة المسارات)
reports_db = {
    "مسار الرقابة (ISO 19011)": {
        "تقرير النزول الميداني": ["نسبة الإنجاز مقارنة بالمخطط:", "حالات عدم المطابقة الفنية:", "مسببات الانحراف (Root Causes):", "مؤشرات الهدر المالي:"],
        "تقرير تفتيش الامتثال": ["المعيار القانوني المرجعي:", "المخالفات المرصودة بالأدلة:", "الأثر المترتب على المخالفة:"]
    },
    "مسار الأثر (Kirkpatrick Model)": {
        "تقرير قياس الأثر": ["التحول الملموس في الفئة المستهدفة:", "مؤشرات النجاح الرقمية (KPIs):", "الاستدامة والعائد المجتمعي:"]
    },
    "مسار الاستراتيجية (Risk Management)": {
        "دراسة الجدوى والمخاطر": ["الفرصة السوقية المستهدفة:", "أخطر التهديدات وخطة الطوارئ:", "فترة استرداد رأس المال:"]
    },
    "مسار العمليات (Lean Management)": {
        "تقرير الإنجاز الدوري": ["المستهدفات المحققة:", "كفاءة الموازنة التشغيلية:", "الفجوات التنفيذية الحالية:"]
    },
    "مسار العلاقات (Visibility)": {
        "تقرير التغطية الإعلامية": ["حجم الوصول الرقمي والميداني:", "قائمة الشركاء والمؤثرين:", "تحليل انطباعات الجمهور:"]
    }
}

# 3. إدارة التوثيق والاشتراك
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/c/c4/Gold_Star_Solid.svg/1024px-Gold_Star_Solid.svg.png", width=60)
    st.header("بوابـة الاشتـراك")
    # ملاحظة: في النسخة الاحترافية يتم ربط هذا بقاعدة بيانات خارجية
    st.info("نظام الأرصدة المسبقة الدفع")
    auth_code = st.text_input("كود التفعيل السيادي:", type="password")
    
st.title("المنصور الاستراتيجية")
st.markdown("### بناء الوثيقة المعتمدة")

# الطبقة الإدارية الإلزامية (الغلاف)
col1, col2 = st.columns(2)
with col1:
    org = st.text_input("الجهة المصدرة للتقرير:")
    proj = st.text_input("اسم المشروع / المهمة:")
with col2:
    zone = st.text_input("النطاق الجغرافي:")
    user = st.text_input("مُعد الوثيقة (الاسم والمنصب):")

st.markdown("---")

# الطبقة التحليلية (المنهجية)
p_choice = st.selectbox("المسار المنهجي:", list(reports_db.keys()))
r_choice = st.selectbox("نوع الوثيقة:", list(reports_db[p_choice].keys()))

answers = {}
for q in reports_db[p_choice][r_choice]:
    answers[q] = st.text_area(q)

st.markdown("---")

# الطبقة الاعتمادية (الختامية)
recs = st.text_area("التوصيات والمقترحات الاستراتيجية للإدارة العليا:")
apps = st.text_input("الملاحق المرفقة (شواهد، صور، كشوفات):")

# 4. محرك التوليد الصارم
if st.button("اعتماد وتوليد الوثيقة"):
    # التحقق المنهجي
    if not (org and proj and user and auth_code):
        st.error("خطأ إداري: يجب استكمال بيانات الغلاف وكود التفعيل.")
    else:
        try:
            genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
            model = genai.GenerativeModel('gemini-1.5-pro')
            
            data_feed = "\n".join([f"{k}: {v}" for k, v in answers.items() if v])
            prompt = f"""
            أنت مستشار توثيق عالمي. صغ '{r_choice}' لجهة '{org}' بخصوص '{proj}'.
            الهيكل الإلزامي:
            1. غلاف التقرير (الجهة، المشروع، الموقع: {zone}، المعد، التاريخ).
            2. الملخص التنفيذي.
            3. التحليل المنهجي بناءً على: {data_feed}.
            4. التوصيات الاستراتيجية: {recs}.
            5. الملاحق: {apps}.
            اللغة: رسمية، رصينة، تعتمد الأرقام.
            """
            
            with st.spinner("جاري المعالجة المنهجية..."):
                response = model.generate_content(prompt)
                st.success("تم الاعتماد")
                st.info(response.text)
                
                # إنشاء ملف Word احترافي
                doc = Document()
                doc.add_heading(f"{org} - {r_choice}", 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
                doc.add_paragraph(f"المشروع: {proj}\nالنطاق الجغرافي: {zone}\nإعداد: {user}").alignment = WD_ALIGN_PARAGRAPH.RIGHT
                for p in response.text.split('\n'):
                    if p.strip():
                        para = doc.add_paragraph(p.strip())
                        para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                
                bio = io.BytesIO()
                doc.save(bio)
                st.download_button("تحميل الوثيقة المعتمدة (Word)", bio.getvalue(), file_name=f"Report_{proj}.docx")
        except Exception as e:
            st.error(f"عطل في المحرك: {e}")
