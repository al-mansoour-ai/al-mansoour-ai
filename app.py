import streamlit as st
import google.generativeai as genai
import os

# ===== إعداد الصفحة =====
st.set_page_config(page_title="المنصور AI", layout="centered")

# ===== تصميم =====
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');

html, body {
    direction: rtl;
    background: #f8fafc;
}

* { font-family: 'Cairo', sans-serif !important; }

.title {
    text-align:center;
    font-size:30px;
    font-weight:900;
    color:#1e3a8a;
}

.card {
    background:white;
    padding:20px;
    border-radius:12px;
    margin-bottom:15px;
    border-right:5px solid #d4af37;
}

.hint {
    font-size:13px;
    color:#64748b;
}

button {
    background: linear-gradient(90deg,#1e3a8a,#d4af37)!important;
    color:white!important;
    height:50px!important;
    border-radius:10px!important;
}
</style>
""", unsafe_allow_html=True)

# ===== API =====
api_key = os.getenv("GEMINI_API_KEY")
if not api_key:
    st.error("⚠️ ضع مفتاح API")
    st.stop()

genai.configure(api_key=api_key)
model = genai.GenerativeModel("gemini-1.5-flash")

# ===== أنواع التقارير =====
REPORTS = {
    "📊 مشروع": [
        ("الهدف العام", "مثال: تحسين الوصول للخدمات"),
        ("أهم الأنشطة", "مثال: تنفيذ ورش تدريب"),
        ("النتائج", "مثال: زيادة الكفاءة بنسبة 30%"),
        ("التحديات", "مثال: تأخر الإمدادات"),
        ("الحلول", "مثال: التعاقد مع مورد بديل")
    ],
    "🎓 تدريب": [
        ("هدف التدريب", "تطوير المهارات"),
        ("الفئة المستهدفة", "موظفين / طلاب"),
        ("مستوى التفاعل", "مرتفع / متوسط"),
        ("نتائج التعلم", "تحسن واضح"),
        ("الأثر", "تطبيق فعلي للمهارات")
    ]
}

# ===== واجهة =====
st.markdown("<div class='title'>🚀 المنصور AI</div>", unsafe_allow_html=True)

rtype = st.selectbox("نوع التقرير", list(REPORTS.keys()))
project = st.text_input("اسم المشروع")

answers = {}

# ===== الأسئلة =====
for i, (q, hint) in enumerate(REPORTS[rtype]):
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.markdown(f"**{q}**")
    st.markdown(f"<div class='hint'>💡 {hint}</div>", unsafe_allow_html=True)
    answers[q] = st.text_area("", key=i)
    st.markdown("</div>", unsafe_allow_html=True)

# ===== توليد (بدون إظهار البرومبت) =====
def generate():
    data = "\n".join([f"{k}: {v}" for k,v in answers.items() if v])

    text = f"""
    تقرير رسمي احترافي:

    المشروع: {project}
    النوع: {rtype}

    {data}

    اكتب التقرير بأسلوب رسمي احترافي مع:
    مقدمة - تحليل - نتائج - توصيات
    """

    return model.generate_content(text).text

# ===== زر =====
if st.button("🚀 إنشاء التقرير"):
    if not project:
        st.warning("أدخل اسم المشروع")
    else:
        with st.spinner("جاري إعداد التقرير..."):
            result = generate()

        st.success("تم إنشاء التقرير")

        st.markdown(f"""
        <div style="
        background:white;
        padding:25px;
        border-radius:10px;
        line-height:2;
        border-right:8px solid #d4af37;">
        {result}
        </div>
        """, unsafe_allow_html=True)
