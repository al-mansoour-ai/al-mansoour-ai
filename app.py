import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import datetime
import random

# ==========================================
# 1. الإعدادات والواجهة السيادية
# ==========================================
st.set_page_config(page_title="المنصور الاستراتيجية | المخرجات العالمية", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700&display=swap');
    #MainMenu, footer, header {visibility: hidden;}
    .stApp { background-color: #0c0c0c !important; }
    h1, h2, h3, h4, p, span, div, label, li { font-family: 'Cairo', sans-serif !important; text-align: right !important; direction: rtl !important; color: #ffffff; }
    h1, h2, h3, h4 { color: #D4AF37 !important; }
    div[data-baseweb="input"] > div, div[data-baseweb="textarea"] > div, div[data-baseweb="select"] > div {
        background-color: #1a1a1a !important; border: 1px solid #D4AF37 !important;
    }
    input, textarea, div[role="listbox"] { color: #ffffff !important; text-align: right !important; direction: rtl !important; }
    .stButton > button { background-color: #D4AF37 !important; color: #0c0c0c !important; font-weight: 700 !important; width: 100% !important; border: none !important; padding: 15px !important; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. قاعدة البيانات المنهجية (الخمسة المسارات)
# ==========================================
reports_tree = {
    "مسار الرقابة والامتثال": {
        "تقرير النزول الميداني (ISO 19011)": [
            ("نسبة الإنجاز الفعلي مقارنة بالمستهدف:", "مثال: المخطط 70%، المنفذ فعلياً 45%"),
            ("حالات عدم المطابقة الفنية المرصودة:", "مثال: استخدام مواد بناء غير مطابقة للمواصفات الفنية"),
            ("مسببات الانحراف الجذرية (Root Causes):", "مثال: تأخر صرف المستخلصات أدى لتوقف العمالة"),
            ("مؤشرات الهدر في الموارد:", "مثال: وجود فائض من الإسمنت معرض للتلف بسبب سوء التخزين")
        ],
        "تقرير تفتيش الالتزام القانوني": [
            ("المعيار القانوني محل التدقيق:", "مثال: لائحة السلامة المهنية بوزارة الأشغال"),
            ("نقاط الضعف في الامتثال:", "مثال: عدم تجديد تراخيص العمل لبعض الكوادر")
        ]
    },
    "مسار الأثر": {
        "تقرير تقييم مشروع (Kirkpatrick)": [
            ("التحول الملموس في واقع الفئة المستهدفة:", "مثال: تحسن مستوى الدخل للأسر المستفيدة بنسبة 30%"),
            ("مؤشرات النجاح الاستراتيجية (KPIs):", "مثال: استدامة وصول المياه لـ 500 فرد يومياً دون انقطاع")
        ]
    },
    "مسار العمليات": {
        "تقرير الإنجاز الدوري": [
            ("المستهدفات التشغيلية المحققة:", "مثال: تم إنجاز 5 ورش تدريبية من أصل 6"),
            ("كفاءة استهلاك الموازنة التشغيلية:", "مثال: تم صرف 80% من الميزانية مع تحقيق 100% من الأهداف")
        ]
    },
    "مسار الاستراتيجية": {
        "دراسة جدوى ومخاطر": [
            ("الفرصة السوقية والميزة التنافسية:", "مثال: الفجوة في سوق الأدوية التخصصية في اليمن"),
            ("أخطر التهديدات وخطة الطوارئ:", "مثال: خطر إغلاق المنافذ (الخطة: توفير مخزون استراتيجي لـ 6 أشهر)")
        ]
    },
    "مسار العلاقات": {
        "تقرير التغطية الإعلامية": [
            ("حجم الوصول والانطباع العام:", "مثال: 50 ألف مشاهدة وتفاعل إيجابي من الجمهور بنسبة 90%"),
            ("قائمة الشركاء والمؤثرين المشاركين:", "مثال: مشاركة 10 وكالات إخبارية محلية ودولية")
        ]
    }
}

# ==========================================
# 3. إدارة الرصيد والباقات
# ==========================================
if "user_balance" not in st.session_state:
    st.session_state.user_balance = 0

with st.sidebar:
    st.header("💳 محفظة الباقات")
    st.write(f"الرصيد الحالي: **{st.session_state.user_balance} تقرير**")
    activation_code = st.text_input("أدخل رمز شحن الباقة:", type="password")
    if st.button("تفعيل الرصيد"):
        valid_codes = {"MANSOUR_3": 3, "EXPERT_10": 10, "STRATEGIC_VIP": 100}
        if activation_code in valid_codes:
            st.session_state.user_balance += valid_codes[activation_code]
            st.success(f"تم شحن {valid_codes[activation_code]} تقارير")
        else:
            st.error("الرمز غير صالح")
    st.markdown("---")
    st.markdown("**لطلب شحن الرصيد:**\n774575749 (كريمي/ون كاش)")

# ==========================================
# 4. بناء الوثيقة (الطبقات الثلاث)
# ==========================================
st.title("المنصور الاستراتيجية")
st.markdown("#### أولاً: الإطار المؤسسي")
col1, col2 = st.columns(2)
with col1:
    entity = st.text_input("اسم الجهة:")
    project = st.text_input("اسم المشروع:")
with col2:
    loc = st.text_input("النطاق الجغرافي:")
    author = st.text_input("مُعد التقرير:")

st.markdown("---")
st.markdown("#### ثانياً: الاستنطاق التحليلي")
pillar = st.selectbox("المسار المنهجي:", list(reports_tree.keys()))
rtype = st.selectbox("نوع الوثيقة:", list(reports_tree[pillar].keys()))

answers = {}
for q, ex in reports_tree[pillar][rtype]:
    answers[q] = st.text_area(q, placeholder=f"إرشاد: {ex}")

st.markdown("---")
st.markdown("#### ثالثاً: التوصيات والملاحق")
recoms = st.text_area("التوصيات الاستراتيجية الختامية:")
apps = st.text_input("الملاحق المرفقة (شواهد، صور، كشوفات):")

# ==========================================
# 5. محرك التوليد والخصم
# ==========================================
if st.button("توليد الوثيقة المعتمدة"):
    if st.session_state.user_balance <= 0:
        st.error("⚠️ رصيدك 0. يرجى شحن باقة للاستمرار.")
    elif not (entity and project and author):
        st.warning("⚠️ يرجى استكمال البيانات الإدارية.")
    else:
        try:
            api_key = st.secrets["GEMINI_API_KEY"]
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-1.5-pro')
            
            context = "\n".join([f"- {k}: {v}" for k, v in answers.items() if v])
            prompt = f"بصفتك مستشاراً عالمياً، صغ {rtype} بأسلوب سيادي رسمي لجهة {entity} بخصوص {project}. البيانات: {context}. التوصيات: {recoms}. الملاحق: {apps}."
            
            with st.spinner("جاري التوليد والخصم من الرصيد..."):
                response = model.generate_content(prompt)
                st.session_state.user_balance -= 1
                st.success(f"تم التوليد! الرصيد المتبقي: {st.session_state.user_balance}")
                st.info(response.text)
                
                # ملف Word
                doc = Document()
                doc.add_heading(rtype, 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
                doc.add_paragraph(f"الجهة: {entity}\nالمشروع: {project}\nإعداد: {author}").alignment = WD_ALIGN_PARAGRAPH.RIGHT
                for line in response.text.split('\n'):
                    if line.strip():
                        p = doc.add_paragraph(line.strip())
                        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                bio = io.BytesIO()
                doc.save(bio)
                st.download_button("تحميل الملف المعتمد", bio.getvalue(), file_name=f"{rtype}.docx")
        except Exception as e:
            st.error(f"خطأ: {e}")
