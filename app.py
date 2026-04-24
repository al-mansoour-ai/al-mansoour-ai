import streamlit as st
import google.generativeai as genai
import openai
import os
from io import BytesIO
from docx import Document

# ===== إعداد الصفحة =====
st.set_page_config(page_title="المنصور AI", layout="centered")

# ===== التصميم (الهوية البصرية للمنصور) =====
st.markdown("""
<style>
.stApp {direction: rtl; background:#f8fafc;}
.title {text-align:center; font-size:35px; font-weight:bold; color:#1e3a8a; border-bottom: 2px solid #d4af37; padding-bottom:10px;}
button {
    background: linear-gradient(90deg,#1e3a8a,#d4af37)!important;
    color:white!important;
    font-weight:bold!important;
    border-radius:10px!important;
}
</style>
""", unsafe_allow_html=True)

# ===== جلب المفاتيح من الأسرار =====
gemini_key = os.getenv("GEMINI_API_KEY")
openai_key = os.getenv("OPENAI_API_KEY")

if gemini_key:
    genai.configure(api_key=gemini_key)
if openai_key:
    client_openai = openai.OpenAI(api_key=openai_key)

# ===== أنواع التقارير =====
REPORT_TYPES = {
    "📊 تقرير مشروع": ["ما الذي تم إنجازه؟","التحديات التي واجهتكم؟","نسبة الإنجاز الحالية؟","التوصيات المستقبلية؟"],
    "🎓 تقرير تدريب": ["ما هو هدف البرنامج التدريبي؟","عدد المشاركين؟","مستوى التفاعل؟","أهم المخرجات والنتائج؟"],
    "💰 تقرير مالي": ["ملخص المصاريف؟","مقارنة بالميزانية؟","الفروقات المالية؟","توصيات الاستدامة؟"]
}

st.markdown("<div class='title'>🚀 منصة المنصور AI الاستراتيجية</div>", unsafe_allow_html=True)
st.write("---")

# ===== واجهة الاختيارات =====
col1, col2 = st.columns(2)
with col1:
    plan = st.radio("إصدار المنصة", ["مجاني (Gemini)", "بريميوم (OpenAI)"])
with col2:
    report_type = st.selectbox("نوع التقرير المطلوبة", list(REPORT_TYPES.keys()))

project_name = st.text_input("اسم المشروع / المؤسسة")

# جمع الإجابات
answers = []
for i, q in enumerate(REPORT_TYPES[report_type]):
    answers.append(st.text_area(q, key=f"q_{i}"))

# ===== دالة التوليد الذكي =====
def generate_report():
    prompt = f"اكتب تقرير احترافي رفيع المستوى لـ {project_name} نوعه {report_type}:\n"
    for q, a in zip(REPORT_TYPES[report_type], answers):
        prompt += f"{q}: {a}\n"
    prompt += "\nالمطلوب: صياغة رسمية، مقدمة، تحليل دقيق، توصيات استراتيجية."

    if "بريميوم" in plan and openai_key:
        response = client_openai.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}]
        )
        return response.choices[0].message.content
    else:
        model = genai.GenerativeModel("gemini-1.5-flash")
        return model.generate_content(prompt).text

# ===== معالجة ملف Word =====
def create_word(text):
    doc = Document()
    doc.add_heading(f"تقرير: {project_name}", 0)
    doc.add_paragraph(text)
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ===== زر التشغيل =====
if st.button("🚀 توليد وتحميل التقرير الآن"):
    if not project_name or not all(answers):
        st.warning("يرجى إكمال جميع الحقول لضمان جودة التقرير.")
    else:
        with st.spinner("جاري صياغة التقرير بذكاء المنصور..."):
            final_report = generate_report()
            st.success("تم توليد التقرير بنجاح!")
            st.markdown("### نص التقرير المقترح:")
            st.info(final_report)
            
            st.download_button(
                label="📄 تحميل التقرير كملف Word",
                data=create_word(final_report),
                file_name=f"AlMansour_Report_{project_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
