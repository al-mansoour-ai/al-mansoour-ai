import streamlit as st
import google.generativeai as genai
import os

# ================== إعداد الصفحة ==================
st.set_page_config(page_title="المنصور AI", layout="centered")

# ================== التصميم ==================
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
    font-size:32px;
    font-weight:900;
    color:#1e3a8a;
}

.card {
    background:white;
    padding:20px;
    border-radius:15px;
    margin-bottom:20px;
    border-right:6px solid #d4af37;
    box-shadow:0 5px 20px rgba(0,0,0,0.05);
}

.hint {
    background:#f1f5f9;
    padding:10px;
    border-radius:10px;
    font-size:14px;
    color:#475569;
}

button {
    background: linear-gradient(90deg,#1e3a8a,#d4af37)!important;
    color:white!important;
    border-radius:10px!important;
    height:50px!important;
    font-weight:bold!important;
}
</style>
""", unsafe_allow_html=True)

# ================== API ==================
api_key = os.getenv("GEMINI_API_KEY")
if not api_key:
    st.error("⚠️ ضع مفتاح GEMINI_API_KEY في secrets")
    st.stop()

genai.configure(api_key=api_key)
model = genai.GenerativeModel("gemini-1.5-flash")

# ================== بنك التقارير ==================
REPORT_ENGINE = {

    "📊 تقرير أداء مشروع (PMBOK)": [
        ("الملخص التنفيذي",
         "مثال: تم تنفيذ 90% من أنشطة المشروع، مع تحقيق نتائج ملموسة في تحسين جودة الخدمات..."),

        ("نطاق المشروع والأنشطة",
         "مثال: شمل المشروع التدريب، التوريد، والمتابعة الميدانية في ثلاث مناطق..."),

        ("تحليل الأداء مقابل الخطة",
         "مثال: نسبة الإنجاز 85% مقابل 90% المخطط بسبب تأخر التوريد..."),

        ("إدارة المخاطر والتحديات",
         "مثال: تم التعامل مع تأخر الإمدادات عبر مورد بديل وخطة طوارئ..."),

        ("جودة المخرجات ورضا المستفيدين",
         "مثال: نسبة رضا المستفيدين بلغت 92% حسب الاستبيانات..."),

        ("الدروس المستفادة",
         "مثال: أهمية التخطيط اللوجستي المبكر وتقليل الاعتماد على مورد واحد..."),

        ("التوصيات الاستراتيجية",
         "مثال: التوسع في المشروع مع تحسين أنظمة المتابعة الرقمية...")
    ],

    "🎓 تقرير تدريب (Kirkpatrick)": [
        ("أهداف التدريب",
         "مثال: تطوير مهارات القيادة وإدارة المشاريع لدى المشاركين..."),

        ("المنهجية التدريبية",
         "مثال: تدريب تفاعلي + تمارين عملية + دراسات حالة..."),

        ("تحليل التفاعل",
         "مثال: أظهر المشاركون تفاعلاً عالياً بنسبة 95%..."),

        ("تحليل التعلم",
         "مثال: ارتفع مستوى المعرفة من 40% إلى 90% بعد التدريب..."),

        ("تطبيق المهارات",
         "مثال: بدأ المشاركون بتطبيق الأدوات في بيئة العمل..."),

        ("الأثر النهائي",
         "مثال: تحسن الأداء المؤسسي وزيادة الإنتاجية..."),

        ("توصيات الاستدامة",
         "مثال: تنفيذ برامج متابعة بعد 3 أشهر...")
    ],

    "💰 تقرير مالي (IFRS)": [
        ("تحليل المصروفات",
         "مثال: تم صرف 80% من الميزانية على الأنشطة التشغيلية..."),

        ("مقارنة الميزانية",
         "مثال: يوجد انحراف بنسبة 10% عن المخطط..."),

        ("تحليل الانحرافات",
         "مثال: الزيادة ناتجة عن ارتفاع الأسعار والتضخم..."),

        ("الامتثال المالي",
         "مثال: جميع العمليات متوافقة مع السياسات المالية..."),

        ("مؤشرات الكفاءة",
         "مثال: انخفاض التكلفة لكل مستفيد بنسبة 15%..."),

        ("المخاطر المالية",
         "مثال: تقلب أسعار السوق يؤثر على التكاليف..."),

        ("التوصيات المالية",
         "مثال: تحسين إدارة الموارد وتخفيض النفقات التشغيلية...")
    ]
}

# ================== الواجهة ==================
st.markdown("<div class='title'>🚀 المنصور AI للتقارير الاحترافية</div>", unsafe_allow_html=True)

rtype = st.selectbox("اختر نوع التقرير", list(REPORT_ENGINE.keys()))
project_name = st.text_input("اسم المشروع")

answers = {}

# ================== عرض الأسئلة ==================
for i, (question, example) in enumerate(REPORT_ENGINE[rtype]):
    st.markdown("<div class='card'>", unsafe_allow_html=True)

    st.markdown(f"**{question}**")
    st.markdown(f"<div class='hint'>💡 {example}</div>", unsafe_allow_html=True)

    txt = st.text_area("", key=f"q_{i}")
    answers[question] = txt

    if st.button(f"✨ تحسين {question}", key=f"btn_{i}"):
        if txt:
            improved = model.generate_content(
                f"حوّل النص التالي إلى صياغة تقرير احترافي عالي المستوى:\n{txt}"
            )
            st.success(improved.text)
        else:
            st.warning("اكتب نص أولاً")

    st.markdown("</div>", unsafe_allow_html=True)

# ================== توليد التقرير ==================
if st.button("🚀 توليد التقرير النهائي"):
    if not project_name:
        st.warning("يرجى إدخال اسم المشروع")
    else:
        with st.spinner("جاري إعداد التقرير الاحترافي..."):

            content = "\n".join([
                f"{k}: {v}" for k, v in answers.items() if v
            ])

            prompt = f"""
            اكتب تقريراً احترافياً عالي المستوى.

            اسم المشروع: {project_name}
            نوع التقرير: {rtype}

            البيانات:
            {content}

            التعليمات:
            - لغة رسمية احترافية
            - تحليل عميق
            - عناوين واضحة
            - توصيات استراتيجية قوية
            """

            result = model.generate_content(prompt)

            st.success("تم إنشاء التقرير")

            st.markdown(f"""
            <div style="
            background:white;
            padding:30px;
            border-radius:12px;
            line-height:2;
            border-right:10px solid #d4af37;">
            {result.text}
            </div>
            """, unsafe_allow_html=True)
