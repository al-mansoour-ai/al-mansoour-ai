import streamlit as st
import google.generativeai as genai
import os
from io import BytesIO
from docx import Document

# 1. إعدادات الصفحة والخط (Cairo) والتصميم الرسمي
st.set_page_config(page_title="منصة المنصور الاستراتيجية", layout="wide")

st.markdown("""
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;700&display=swap" rel="stylesheet">
    <style>
    * {font-family: 'Cairo', sans-serif; direction: rtl; text-align: right;}
    .stApp {background-color: #f4f7f9;}
    
    /* الهيدر العلوي */
    .main-header {
        background: linear-gradient(90deg, #1e3a8a 0%, #152e6d 100%);
        color: #d4af37;
        padding: 2rem;
        border-radius: 15px;
        text-align: center;
        border-bottom: 5px solid #d4af37;
        margin-bottom: 2rem;
    }
    
    /* الأزرار */
    .stButton>button {
        width: 100%;
        background: #d4af37 !important;
        color: #1e3a8a !important;
        font-weight: bold !important;
        border: none !important;
        transition: 0.3s;
    }
    .stButton>button:hover {
        background: #b8962e !important;
        transform: translateY(-2px);
    }
    
    /* صناديق المدخلات */
    .stTextArea textarea {border: 1px solid #1e3a8a !important;}
    </style>
    
    <div class="main-header">
        <h1>🚀 منصة المنصور الاستراتيجية للذكاء الاصطناعي</h1>
        <p style="color: white;">الجيل القادم في صياغة التقارير وخطط العمل الاحترافية</p>
    </div>
""", unsafe_allow_html=True)

# 2. ربط الذكاء الاصطناعي
api_key = st.secrets["GEMINI_API_KEY"]
genai.configure(api_key=api_key)
model = genai.GenerativeModel("gemini-1.5-flash")

# 3. محرك التقارير الاستشارية
tabs = st.tabs(["📝 صياغة تقرير احترافي", "🔍 تحسين وتدقيق نص", "💡 أمثلة ونماذج"])

with tabs[0]:
    col1, col2 = st.columns([1, 2])
    with col1:
        st.subheader("إعدادات التقرير")
        report_type = st.selectbox("نوع التقرير", [
            "تقرير إنجاز دوري", 
            "خطة عمل استراتيجية", 
            "تقرير تقييم أداء", 
            "مقترح مشروع (Proposal)",
            "دراسة حالة استشارية"
        ])
        tone = st.select_slider("لهجة الكتابة", ["رسمي بسيط", "احترافي استشاري", "قيادي رفيع"])
        
    with col2:
        st.subheader("بيانات التقرير")
        context = st.text_area("أدخل النقاط الأساسية أو المسودة (حتى لو كانت غير مرتبة)", height=200, 
                              placeholder="مثال: تم تدريب 50 شخص، واجهنا مشكلة في القاعة، النتيجة كانت ممتازة...")
        
        if st.button("توليد التقرير النهائي"):
            if context:
                with st.spinner("جاري الصياغة بأسلوب استشاري رفيع..."):
                    prompt = f"بصفتك خبير تطوير مؤسسي، صغ لي {report_type} بلهجة {tone}. استخدم لغة عربية فصحى قوية، رتب الأفكار في نقاط، وأضف توصيات استراتيجية بناءً على هذا السياق: {context}. اجعل الخطاب موجه للإدارة العليا."
                    response = model.generate_content(prompt)
                    st.session_state['last_report'] = response.text
                    st.markdown("### النتيجة النهائية:")
                    st.info(response.text)
            else:
                st.error("الرجاء إدخال بيانات التقرير أولاً")

with tabs[1]:
    st.subheader("محرك تحسين النصوص")
    raw_text = st.text_area("ألصق النص الذي تريد تحسينه هنا...", height=150)
    if st.button("تحسين النص لغوياً واستراتيجياً"):
        if raw_text:
            with st.spinner("جاري معالجة النص..."):
                prompt = f"قم بإعادة صياغة هذا النص ليكون أكثر احترافية وقوة، صحح الأخطاء الإملائية، واستخدم مصطلحات إدارية حديثة: {raw_text}"
                response = model.generate_content(prompt)
                st.success("النص بعد التحسين:")
                st.write(response.text)

with tabs[2]:
    st.info("💡 **أمثلة لما يمكنك القيام به:**")
    st.write("""
    * **تقرير تدريب:** اذكر فقط (اسم الدورة، عدد المتدربين، انطباعك) وسيقوم النظام ببناء مقدمة، محاور، نتائج، وتوصيات.
    * **خطة عمل:** اذكر (الهدف، المدة، الميزانية) وسيقوم النظام بتوزيع المهام وتحليل المخاطر.
    * **تحسين إيميل:** ألصق إيميل عادي وسيحوله إلى خطاب رسمي موجه لمدير شركة أو منظمة.
    """)

# 4. وظيفة تحميل ملف Word
def to_word(text):
    doc = Document()
    doc.add_heading("تقرير منصة المنصور الاستراتيجية", 0)
    p = doc.add_paragraph(text)
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

if 'last_report' in st.session_state:
    st.download_button("⬇️ تحميل التقرير كملف Word", data=to_word(st.session_state['last_report']), file_name="Almansour_Report.docx")
