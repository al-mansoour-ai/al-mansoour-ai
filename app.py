import streamlit as st
import google.generativeai as genai
import os
from io import BytesIO
from docx import Document
from datetime import datetime
from reportlab.platypus import SimpleDocTemplate, Paragraph
from reportlab.lib.styles import getSampleStyleSheet

# ===== إعداد =====
st.set_page_config(page_title="المنصور AI SaaS", layout="centered")

# ===== API =====
api_key = st.secrets.get("GEMINI_API_KEY")
if not api_key:
    st.error("ضع API KEY في Secrets")
    st.stop()

genai.configure(api_key=api_key)
model = genai.GenerativeModel("gemini-1.5-flash")

# ===== نظام المستخدم =====
if "user" not in st.session_state:
    st.session_state.user = None

if "usage" not in st.session_state:
    st.session_state.usage = 0

# ===== تسجيل دخول بسيط =====
if not st.session_state.user:
    st.title("🔐 تسجيل الدخول")

    username = st.text_input("اسم المستخدم")
    password = st.text_input("كلمة المرور", type="password")

    if st.button("دخول"):
        if username and password:
            st.session_state.user = username
            st.rerun()
    st.stop()

# ===== خطة المستخدم =====
plan = st.radio("نوع الحساب", ["مجاني", "مدفوع"])

# ===== حد الاستخدام =====
if plan == "مجاني":
    if st.session_state.usage >= 5:
        st.error("وصلت الحد المجاني اليومي")
        st.stop()

# ===== أنواع التقارير =====
REPORTS = {
    "📊 مشروع": ["الإنجازات", "التحديات", "النتائج", "التوصيات"],
    "🎓 تدريب": ["الهدف", "المشاركين", "التفاعل", "النتائج"],
    "💰 مالي": ["المصاريف", "الميزانية", "الفرق", "التوصيات"]
}

st.title("🚀 المنصور AI")

rtype = st.selectbox("نوع التقرير", list(REPORTS.keys()))
project = st.text_input("اسم المشروع")

answers = []
for i, q in enumerate(REPORTS[rtype]):
    answers.append(st.text_area(q, key=i))

# ===== AI آمن =====
def generate(prompt):
    try:
        return model.generate_content(prompt).text
    except:
        return "⚠️ خطأ في التوليد"

# ===== توليد =====
if st.button("🚀 توليد التقرير"):
    if not project or not all(answers):
        st.warning("أكمل البيانات")
    else:
        with st.spinner("جاري التوليد..."):
            prompt = f"تقرير احترافي {project}\n"
            for q, a in zip(REPORTS[rtype], answers):
                prompt += f"{q}: {a}\n"

            report = generate(prompt)

        st.session_state.usage += 1

        st.success("تم إنشاء التقرير")
        st.text_area("📄 التقرير", report, height=300)

        # ===== حفظ =====
        if "history" not in st.session_state:
            st.session_state.history = []

        st.session_state.history.append({
            "date": datetime.now().strftime("%Y-%m-%d"),
            "report": report
        })

        # ===== Word =====
        doc = Document()
        doc.add_paragraph(report)
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.download_button("📄 تحميل Word", buffer, "report.docx")

        # ===== PDF (مدفوع فقط) =====
        if plan == "مدفوع":
            pdf_buffer = BytesIO()
            doc_pdf = SimpleDocTemplate(pdf_buffer)
            styles = getSampleStyleSheet()
            content = [Paragraph(report, styles["Normal"])]
            doc_pdf.build(content)
            pdf_buffer.seek(0)

            st.download_button("📕 تحميل PDF", pdf_buffer, "report.pdf")
        else:
            st.info("PDF متاح فقط للحساب المدفوع")

# ===== سجل =====
st.write("📊 سجل التقارير")

if "history" in st.session_state:
    for item in st.session_state.history:
        st.write(item["date"])
