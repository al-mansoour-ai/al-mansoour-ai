import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# ==========================================
# 1. إعدادات الصفحة والهوية البصرية (Branding)
# ==========================================
st.set_page_config(page_title="المنصور الاستراتيجية | التقارير السيادية", layout="wide", initial_sidebar_state="expanded")

# تخصيص CSS للون الأسود الفخم والذهبي والتوافق مع الجوال
custom_css = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700&display=swap');
    
    html, body, [class*="st-"] {
        font-family: 'Cairo', sans-serif !important;
        background-color: #0c0c0c !important;
        color: #ffffff !important;
    }
    
    .stMarkdown, .stTextInput, .stTextArea, .stRadio, .stButton, div[data-testid="stSidebar"] {
        direction: rtl;
        text-align: right;
    }
    
    h1, h2, h3 {
        color: #D4AF37 !important;
        font-weight: bold;
        direction: rtl;
        text-align: right;
    }
    
    .stTextInput>div>div>input, .stTextArea>div>div>textarea {
        background-color: #1a1a1a !important;
        color: #ffffff !important;
        border: 1px solid #D4AF37 !important;
        border-radius: 4px;
        direction: rtl;
        text-align: right;
    }
    
    .stButton>button {
        background-color: #D4AF37 !important;
        color: #0c0c0c !important;
        font-weight: 700 !important;
        font-family: 'Cairo', sans-serif !important;
        border: none !important;
        padding: 10px 24px !important;
        width: 100%;
        transition: 0.3s;
    }
    .stButton>button:hover {
        background-color: #ffffff !important;
        color: #000000 !important;
    }
    
    [data-testid="stSidebar"] {
        background-color: #111111 !important;
        border-left: 2px solid #D4AF37;
    }

    @media (max-width: 768px) {
        [data-testid="column"] {
            width: 100% !important;
            flex: 1 1 100% !important;
            min-width: 100% !important;
        }
    }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ==========================================
# 2. دوال المخرجات (Output Functions)
# ==========================================
def create_word_doc(text_content, title_text):
    """دالة تحويل النص المولد إلى ملف Word منسق"""
    doc = Document()
    title = doc.add_heading(f'{title_text} - سري وتنفيذي', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    for paragraph in text_content.split('\n'):
        if paragraph.strip():
            p = doc.add_paragraph(paragraph.strip())
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# ==========================================
# 3. قاعدة البيانات الاستقصائية (The Brain)
# ==========================================
reports_database = {
    "مسار الرقابة: تقرير النزول الميداني": [
        "النطاق الجغرافي والمقاول المنفذ:",
        "نسبة الإنجاز الفعلي مقارنة بالمستهدف (%):",
        "مسببات الانحراف أو التأخير (إن وجدت):",
        "حالات عدم المطابقة الفنية مع كراسة الشروط:",
        "مظاهر الهدر المالي أو تكدس الموارد:",
        "المخاطر الكامنة وبروتوكولات السلامة:",
        "التوجيهات التصحيحية العاجلة المطلوبة:"
    ],
    "مسار الأثر: تقرير ختام مشروع": [
        "الغاية الاستراتيجية والمشكلة الجذرية للمشروع:",
        "عدد المستفيدين (المباشرين وغير المباشرين):",
        "العائد المجتمعي والأثر الملموس:",
        "مؤشرات النجاح والأرقام المحققة (KPIs):",
        "القصة البارزة (حالة واقعية للنجاح):",
        "الدروس المستفادة والتوصية المستقبلية:"
    ],
    "مسار العمليات: تقرير الإنجاز الدوري": [
        "الأهداف التشغيلية المخطط إنجازها:",
        "الإنجاز الفعلي بلغة الأرقام:",
        "الفجوة التشغيلية وأسبابها الجذرية:",
        "التحديات اللوجستية التي واجهت العمل:",
        "التدخلات السريعة والقرارات الإدارية المتخذة:"
    ]
}

# ==========================================
# 4. واجهة المستخدم الديناميكية (Dynamic UI)
# ==========================================
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/c/c4/Gold_Star_Solid.svg/1024px-Gold_Star_Solid.svg.png", width=50)
    st.header("إعدادات النظام السيادي")
    api_key_input = st.text_input("مفتاح المنصة (API Key):", type="password")
    st.markdown("---")
    selected_report = st.radio("حدد المسار الاستراتيجي:", list(reports_database.keys()))

st.title("المنصور الاستراتيجية")
st.markdown(f"**غرفة الاستنطاق | {selected_report}**")
st.markdown("أدخل المعطيات بدقة، وسيتولى المحرك صياغة الوثيقة التنفيذية.")
st.markdown("---")

# توليد الأسئلة برمجياً بناءً على اختيار العميل
answers_dict = {}
for i, question in enumerate(reports_database[selected_report]):
    if "مسببات" in question or "التوجيهات" in question or "القصة" in question or "الفجوة" in question:
        answers_dict[question] = st.text_area(f"{i+1}. {question}")
    else:
        answers_dict[question] = st.text_input(f"{i+1}. {question}")

st.markdown("---")

# ==========================================
# 5. منطق التوليد الموجه (Contextual Generation)
# ==========================================
if st.button("توليد التقرير السيادي", use_container_width=True):
    if not api_key_input:
        st.error("خطأ تنفيذي: المنصة تتطلب إدراج المفتاح السري (API Key).")
    elif not any(answers_dict.values()):
        st.warning("تنبيه: لا يمكن توليد تقرير من فراغ، أجب عن معطى واحد على الأقل.")
    else:
        with st.spinner("جاري معالجة البيانات وصياغة الوثيقة..."):
            try:
                formatted_answers = "\n".join([f"- {k} {v}" for k, v in answers_dict.items() if v])
                
                # توجيه المحرك
                genai.configure(api_key=api_key_input)
                model = genai.GenerativeModel('gemini-1.5-pro')
                
                system_prompt = f"""
                أنت مستشار تنفيذي بمنصة 'المنصور الاستراتيجية'. 
                المطلوب صياغة '{selected_report}' فخم، صارم، ومباشر.
                استخدم لغة الأرقام، وتجنب الحشو. نظم التقرير في فقرات واضحة ومهنية تعكس الجدية والسيادة.
                المعطيات:
                {formatted_answers}
                """
                
                response = model.generate_content(system_prompt)
                final_report = response.text
                
                st.success("تم الاعتماد والتوليد.")
                st.info(final_report)
                
                # تجهيز ملف التحميل
                docx_file = create_word_doc(final_report, selected_report)
                st.download_button(
                    label="تحميل الوثيقة الرسمية (Word)",
                    data=docx_file,
                    file_name=f"{selected_report.replace(' ', '_').replace(':', '')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"انقطاع في الاتصال بالمحرك: {e}")
