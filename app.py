import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import datetime

# 1. المعمارية البصرية (إلغاء الجوانب لضمان ثبات الجوال)
st.set_page_config(page_title="المنصور الاستراتيجية", layout="wide")
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700&display=swap');
    #MainMenu, footer, header {visibility: hidden;}
    .stApp { background-color: #0c0c0c !important; }
    h1, h2, h3, h4, p, span, div, label, li { font-family: 'Cairo', sans-serif !important; text-align: right !important; direction: rtl !important; color: #ffffff; }
    h1, h2, h3 { color: #D4AF37 !important; margin-bottom: 0px; }
    input, textarea, div[role="listbox"], .stSelectbox > div { background-color: #1a1a1a !important; border: 1px solid #D4AF37 !important; color: #ffffff !important; text-align: right !important; }
    .stButton > button { background-color: #D4AF37 !important; color: #0c0c0c !important; font-weight: 700 !important; width: 100% !important; padding: 15px !important; }
    .status-box { border: 2px solid #D4AF37; padding: 20px; border-radius: 10px; background-color: #111; margin-bottom: 25px; }
</style>
""", unsafe_allow_html=True)

# 2. قاعدة البيانات العالمية (المسارات + الفروع + الأمثلة)
reports_db = {
    "مسار الرقابة (ISO 19011)": {
        "تقرير النزول الميداني": [
            ("نسبة الإنجاز مقارنة بالمخطط:", "مثال: المخطط 60%، المنفذ فعلياً 40%"),
            ("حالات عدم المطابقة الفنية:", "مثال: قطر الأنابيب 4 إنش بدلاً من 6 إنش المعتمدة"),
            ("مسببات الانحراف الجذرية:", "مثال: تأخر توريد المواد بسبب أزمة الوقود"),
            ("مؤشرات الهدر المالي:", "مثال: بقاء المعدات مستأجرة دون عمل لمدة 10 أيام")
        ],
        "تقرير تفتيش الامتثال": [
            ("المعيار القانوني المرجعي:", "مثال: مادة السلامة المهنية في قانون العمل"),
            ("المخالفات المرصودة بالأدلة:", "مثال: عدم ارتداء خوذ السلامة في منطقة الرافعة")
        ]
    },
    "مسار الأثر (Kirkpatrick Model)": {
        "تقرير تقييم أثر التدريب": [
            ("التحول الملموس في الأداء:", "مثال: انخفاض زمن المعاملة من ساعة إلى 20 دقيقة"),
            ("مؤشرات النجاح الرقمية (KPIs):", "مثال: وصول الخدمة لـ 500 أسرة إضافية شهرياً")
        ]
    },
    "مسار الاستراتيجية": {
        "دراسة الجدوى والمخاطر": [
            ("الفرصة السوقية المستهدفة:", "مثال: سد فجوة توريد الطاقة المتجددة للمناطق الريفية"),
            ("أهم 3 مخاطر وخطة العلاج:", "مثال: تقلب سعر الصرف (العلاج: الشراء المسبق للأصول)")
        ]
    }
}

# 3. إدارة الاشتراك والرصيد (في صدر الواجهة)
st.title("المنصور الاستراتيجية")
with st.container():
    st.markdown('<div class="status-box">', unsafe_allow_html=True)
    st.subheader("💳 بوابة شحن الباقات")
    col_sub1, col_sub2 = st.columns([2, 1])
    with col_sub1:
        auth_code = st.text_input("أدخل كود تفعيل الباقة (لفتح التوليد):", type="password")
    with col_sub2:
        st.write("لطلب الأكواد: 774575749")
        st.write("الباقات: [منجز (3) - خبير (10) - سيادي (VIP)]")
    st.markdown('</div>', unsafe_allow_html=True)

# 4. بناء الوثيقة (الطبقة الإدارية)
st.markdown("### 🏛️ أولاً: البيانات الإدارية (غلاف الوثيقة)")
col_meta1, col_meta2 = st.columns(2)
with col_meta1:
    org = st.text_input("الجهة المصدرة للتقرير:", placeholder="مثال: مؤسسة شباب اليمن")
    proj = st.text_input("اسم المشروع / المهمة:", placeholder="مثال: مشروع الاستجابة الطارئة")
with col_meta2:
    zone = st.text_input("النطاق الجغرافي:", placeholder="مثال: اليمن - محافظة مأرب")
    user = st.text_input("مُعد الوثيقة (الاسم والمنصب):")

st.markdown("---")

# 5. الطبقة المنهجية (المسارات والفرعية والأمثلة)
st.markdown("### 🔍 ثانياً: الاستنطاق المنهجي (المسارات العالمية)")
p_choice = st.selectbox("حدد المسار الرئيسي:", list(reports_db.keys()))
r_choice = st.selectbox("حدد التقرير الفرعي:", list(reports_db[p_choice].keys()))

st.info(f"نظام الاستنطاق مفعل لـ: {r_choice}")

answers = {}
# عرض الأسئلة مع الأمثلة التوضيحية (Placeholders)
for q_text, q_hint in reports_db[p_choice][r_choice]:
    answers[q_text] = st.text_area(q_text, placeholder=f"إرشاد: {q_hint}")

st.markdown("---")

# 6. الطبقة الاعتمادية
st.markdown("### 📝 ثالثاً: الخواتيم والاعتماد")
recs = st.text_area("التوصيات والمقترحات الاستراتيجية الختامية:")
apps = st.text_input("الملاحق المرفقة (شواهد، صور، روابط):")

# 7. المحرك التنفيذي
if st.button("اعتماد وتوليد الوثيقة السيادية"):
    # التحقق من كود التفعيل والبيانات
    if not auth_code:
        st.error("⚠️ يجب إدخال كود تفعيل الباقة لتشغيل المحرك الذكي.")
    elif not (org and proj and user):
        st.warning("⚠️ يرجى استكمال بيانات غلاف الوثيقة (الجهة، المشروع، المعد).")
    else:
        try:
            genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
            model = genai.GenerativeModel('gemini-1.5-pro')
            
            data_feed = "\n".join([f"{k} {v}" for k, v in answers.items() if v])
            prompt = f"بصفتك مستشاراً عالمياً، صغ تقرير {r_choice} لجهة {org} حول {proj}. النطاق: {zone}. البيانات الميدانية: {data_feed}. التوصيات: {recs}. الملاحق: {apps}."
            
            with st.spinner("جاري المعالجة المنهجية وصياغة الملف..."):
                response = model.generate_content(prompt)
                st.success("تم الاعتماد بنجاح")
                st.markdown("#### معاينة الوثيقة:")
                st.info(response.text)
                
                # تصدير Word
                doc = Document()
                doc.add_heading(f"{org} - {r_choice}", 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
                doc.add_paragraph(f"المشروع: {proj} | الموقع: {zone}\nالتاريخ: {datetime.date.today()}\nإعداد: {user}").alignment = WD_ALIGN_PARAGRAPH.RIGHT
                for p in response.text.split('\n'):
                    if p.strip():
                        para = doc.add_paragraph(p.strip())
                        para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                
                bio = io.BytesIO()
                doc.save(bio)
                st.download_button("تحميل الوثيقة الرسمية (Word)", bio.getvalue(), file_name=f"Report_{proj}.docx")
        except Exception as e:
            st.error(f"عطل في المحرك: {e}")
