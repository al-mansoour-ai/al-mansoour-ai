import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# ==========================================
# 1. إعدادات الصفحة والهوية البصرية (Branding)
# ==========================================
st.set_page_config(page_title="منصة التقارير السيادية", layout="wide", initial_sidebar_state="expanded")

# تخصيص CSS للون الأسود الفخم والذهبي وخط Cairo
custom_css = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700&display=swap');
    
    /* الإعدادات الأساسية والاتجاه */
    html, body, [class*="st-"] {
        font-family: 'Cairo', sans-serif !important;
        direction: rtl;
        text-align: right;
        background-color: #0c0c0c !important;
        color: #ffffff !important;
    }
    
    /* العناوين باللون الذهبي */
    h1, h2, h3 {
        color: #D4AF37 !important;
        font-weight: bold;
    }
    
    /* تصميم حقول الإدخال */
    .stTextInput>div>div>input, .stTextArea>div>div>textarea {
        background-color: #1a1a1a !important;
        color: #ffffff !important;
        border: 1px solid #D4AF37 !important;
        border-radius: 4px;
    }
    
    /* تصميم الأزرار التنفيذية */
    .stButton>button {
        background-color: #D4AF37 !important;
        color: #0c0c0c !important;
        font-weight: 700 !important;
        font-family: 'Cairo', sans-serif !important;
        border: none !important;
        padding: 10px 24px !important;
        transition: 0.3s;
    }
    .stButton>button:hover {
        background-color: #ffffff !important;
        color: #000000 !important;
    }
    
    /* اللوحة الجانبية */
    [data-testid="stSidebar"] {
        background-color: #111111 !important;
        border-left: 2px solid #D4AF37;
    }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ==========================================
# 2. دوال المعالجة والمخرجات (Processing & Output)
# ==========================================

def generate_report(api_key, answers):
    """دالة إرسال البيانات إلى Gemini واستلام التقرير المصاغ"""
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-1.5-pro') # أو gemini-pro
    
    # هندسة الأوامر (Prompt Engineering) المخفية
    system_prompt = f"""
    أنت مستشار تنفيذي فائق الاحترافية. مهمتك كتابة 'تقرير نزول ميداني' رصين، صارم، ومباشر بناءً على المعطيات التالية فقط.
    استخدم لغة الأرقام، وتجنب الحشو والعبارات الإنشائية. صغ التقرير في 4 أقسام رئيسية:
    1. الملخص التنفيذي والجدول الزمني.
    2. المطابقة الفنية وكفاءة التشغيل.
    3. الرقابة المالية وإدارة المخاطر.
    4. التوصيات والتدخلات العاجلة.
    
    المعطيات من المفتش الميداني:
    {answers}
    """
    
    response = model.generate_content(system_prompt)
    return response.text

def create_word_doc(text_content):
    """دالة تحويل النص المولد إلى ملف Word منسق"""
    doc = Document()
    
    # عنوان التقرير
    title = doc.add_heading('تقرير النزول الميداني - سري وتنفيذي', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # إضافة النص مع ضبط الاتجاه لليمين
    for paragraph in text_content.split('\n'):
        if paragraph.strip():
            p = doc.add_paragraph(paragraph.strip())
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            # ملاحظة: التنسيق المتقدم للخطوط العربية في Word يتطلب إعدادات XML، 
            # لكن هذا التنسيق يفي بالغرض للنسخة الأولية.
            
    # حفظ الملف في الذاكرة لتنزيله
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# ==========================================
# 3. واجهة المستخدم (UI Build)
# ==========================================

# اللوحة الجانبية (Sidebar)
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/c/c4/Gold_Star_Solid.svg/1024px-Gold_Star_Solid.svg.png", width=50) # شعار مؤقت
    st.header("إعدادات النظام")
    api_key_input = st.text_input("أدخل مفتاح Gemini API:", type="password")
    st.markdown("---")
    st.markdown("**مسار العمليات الحالي:**")
    st.radio("اختر نوع التقرير:", ["تقرير النزول الميداني", "تقرير الامتثال (قريباً)", "تقرير الأثر (قريباً)"])

# الواجهة الرئيسية
st.title("المنصور الاستراتيجية | منصة التقارير السيادية")
st.markdown("أهلاً بك في غرفة الاستنطاق. أجب عن المعطيات التالية بدقة لبناء وثيقتك التنفيذية.")
st.markdown("---")

# هيكل المدخلات (The Inputs)
col1, col2 = st.columns(2)

with col1:
    st.subheader("أولاً: الإطار والجدول الزمني")
    q1 = st.text_input("1. النطاق الجغرافي والمقاول المنفذ:")
    q2 = st.text_input("2. نسبة الإنجاز الفعلي مقارنة بالمستهدف (%):")
    q3 = st.text_area("3. مسببات الانحراف أو التأخير (إن وجدت):")
    
    st.subheader("ثالثاً: الرقابة المالية")
    q6 = st.text_area("6. مظاهر الهدر المالي أو تكدس العمالة والموارد:")
    q7 = st.text_input("7. النفقات غير المجدولة أو المطالبات الإضافية:")

with col2:
    st.subheader("ثانياً: المطابقة الفنية")
    q4 = st.text_area("4. حالات عدم المطابقة الفنية مع كراسة الشروط:")
    q5 = st.text_input("5. تقييم كفاءة المواد والمعدات في الموقع:")
    
    st.subheader("رابعاً: المخاطر والقرارات")
    q8 = st.text_input("8. مستوى الالتزام ببروتوكولات السلامة (HSE):")
    q9 = st.text_input("9. المخاطر الكامنة (أمنية، بيئية، تشغيلية):")
    q10 = st.text_area("10. التوجيهات التصحيحية العاجلة المطلوبة:")

# تجميع الإجابات
answers_dict = {
    "النطاق": q1, "نسبة الإنجاز": q2, "مسببات التأخير": q3,
    "عدم المطابقة": q4, "الكفاءة": q5, "الهدر المالي": q6,
    "النفقات الإضافية": q7, "السلامة": q8, "المخاطر": q9,
    "التوجيهات العاجلة": q10
}

st.markdown("---")

# ==========================================
# 4. منطق التوليد (Generation Logic)
# ==========================================
if st.button("توليد التقرير السيادي"):
    if not api_key_input:
        st.error("خطأ تنفيذي: يرجى إدخال مفتاح API الخاص بك في اللوحة الجانبية أولاً.")
    elif not any(answers_dict.values()):
        st.warning("تنبيه: يرجى تعبئة حقل واحد على الأقل قبل التوليد.")
    else:
        with st.spinner("جاري صياغة الوثيقة التنفيذية..."):
            try:
                # تجميع البيانات في نص واحد للمحرك
                formatted_answers = "\n".join([f"- {k}: {v}" for k, v in answers_dict.items() if v])
                
                # استدعاء دالة التوليد
                final_report = generate_report(api_key_input, formatted_answers)
                
                # عرض النتيجة
                st.success("تم توليد التقرير بنجاح.")
                st.markdown("### معاينة التقرير:")
                st.info(final_report)
                
                # تجهيز ملف Word للتحميل
                docx_file = create_word_doc(final_report)
                st.download_button(
                    label="تحميل التقرير (Word)",
                    data=docx_file,
                    file_name="تقرير_نزول_ميداني_تنفيذي.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
            except Exception as e:
                st.error(f"حدث خطأ أثناء الاتصال بالمحرك: {e}")
