import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# ==========================================
# 1. الإعدادات السيادية والهوية البصرية
# ==========================================
st.set_page_config(
    page_title="المنصور الاستراتيجية | إدارة التقارير السيادية",
    layout="wide"
)

# كود CSS لتنظيف الواجهة وتنسيقها للجوال
clean_ui_css = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700&display=swap');
    
    /* إخفاء القوائم الافتراضية لـ Streamlit */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* الهوية البصرية (أسود وذهبي) */
    .stApp { background-color: #0c0c0c !important; }
    h1, h2, h3, p, span, div, label, li { font-family: 'Cairo', sans-serif !important; }
    
    h1, h2, h3 { color: #D4AF37 !important; text-align: right !important; direction: rtl !important; }
    .stMarkdown, label, .stRadio, p, .stSelectbox { text-align: right !important; direction: rtl !important; color: #ffffff !important; }
    
    /* تنسيق الحقول والقوائم المنسدلة */
    div[data-baseweb="input"] > div, div[data-baseweb="textarea"] > div, div[data-baseweb="select"] > div {
        background-color: #1a1a1a !important;
        border: 1px solid #D4AF37 !important;
        color: #ffffff !important;
    }
    input, textarea, div[role="listbox"] { color: #ffffff !important; text-align: right !important; direction: rtl !important; }
    
    /* الزر التنفيذي */
    .stButton > button {
        background-color: #D4AF37 !important;
        color: #0c0c0c !important;
        font-weight: 700 !important;
        width: 100% !important;
        border: none !important;
        padding: 15px !important;
        font-size: 18px !important;
        margin-top: 20px !important;
    }
</style>
"""
st.markdown(clean_ui_css, unsafe_allow_html=True)

# ==========================================
# 2. قاعدة البيانات الشجرية (The Sovereign Tree)
# ==========================================
reports_tree = {
    "مسار الرقابة والامتثال": {
        "تقرير النزول الميداني": [
            "النطاق الجغرافي والمقاول المنفذ:", "نسبة الإنجاز الفعلي مقارنة بالمستهدف (%):",
            "مسببات الانحراف أو التأخير:", "حالات عدم المطابقة مع كراسة الشروط:",
            "مظاهر الهدر المالي أو الموارد:", "المخاطر الكامنة وبروتوكولات السلامة:",
            "التوجيهات التصحيحية العاجلة المطلوبة:"
        ],
        "تقرير تفتيش الامتثال": [
            "المعيار القانوني محل التفتيش:", "درجة الالتزام (عالية/متوسطة/منخفضة):",
            "المخالفات المرصودة بالأدلة:", "الأثر المترجم للمخالفة:", "الإجراء العقابي أو التصحيحي المقترح:"
        ]
    },
    "مسار الأثر": {
        "تقرير ختام وتقييم مشروع": [
            "الغاية الاستراتيجية للمشروع:", "إجمالي عدد المستفيدين (أرقام):",
            "العائد المجتمعي الملموس:", "مؤشرات النجاح المحققة (KPIs):",
            "الدروس المستفادة للاستدامة:", "التوصية بنسخ التجربة (نعم/لا مع السبب):"
        ]
    },
    "مسار الاستراتيجية": {
        "دراسة جدوى ومخاطر": [
            "الفرصة السوقية المستهدفة:", "حجم الاستثمار الرأسمالي (CAPEX):",
            "الميزة التنافسية السيادية:", "أخطر 3 تهديدات للفشل وكيفية علاجها:",
            "فترة استرداد رأس المال المتوقعة:"
        ]
    },
    "مسار العمليات": {
        "تقرير الإنجاز الدوري": [
            "المستهدفات التشغيلية للفترة:", "نسبة تحقق الأداء الفعلي:",
            "الفجوات التنفيذية وأسبابها:", "كفاءة استهلاك الموازنة التشغيلية:",
            "خطة العمل للمرحلة القادمة:"
        ]
    },
    "مسار العلاقات": {
        "تقرير التغطية الإعلامية": [
            "الرسالة الذهنية المستهدفة:", "حجم الوصول الرقمي والميداني:",
            "قائمة الشركاء والمؤثرين المشاركين:", "تحليل انطباعات الجمهور:",
            "الأصول الرقمية الموثقة (صور/فيديو):"
        ]
    }
}

# ==========================================
# 3. المنطق الخلفي والمخرجات
# ==========================================
def create_docx(text, title):
    doc = Document()
    h = doc.add_heading(title, 0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for line in text.split('\n'):
        if line.strip():
            p = doc.add_paragraph(line.strip())
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# ==========================================
# 4. بناء الواجهة الرئيسية (التوجيه المباشر)
# ==========================================
st.title("المنصور الاستراتيجية")
st.markdown("### 🏛️ توجيه المسار التنفيذي")

# نقلنا القوائم من الجانب إلى الواجهة الرئيسية
pillar = st.selectbox("1. حدد المسار الاستراتيجي أولاً:", list(reports_tree.keys()))
report_type = st.selectbox("2. حدد نوع الوثيقة المطلوبة:", list(reports_tree[pillar].keys()))

st.markdown("---")
st.subheader(f"غرفة الاستنطاق: {report_type}")
st.markdown("---")

# توليد حقول الأسئلة بناءً على الاختيار الشجري
answers = {}
questions = reports_tree[pillar][report_type]

for i, q in enumerate(questions):
    if "مسببات" in q or "التوجيهات" in q or "الأثر" in q or "الدروس" in q or "الفجوات" in q:
        answers[q] = st.text_area(f"{i+1}. {q}", height=100)
    else:
        answers[q] = st.text_input(f"{i+1}. {q}")

st.markdown("---")

# ==========================================
# 5. محرك التوليد (The Secret Engine)
# ==========================================
if st.button("توليد واعتماد الوثيقة السيادية"):
    try:
        # استدعاء المفتاح من خزنة الأسرار
        api_key = st.secrets["GEMINI_API_KEY"]
        
        if not any(answers.values()):
            st.warning("تنبيه: يرجى تقديم بيانات استقصائية ليتمكن المحرك من التحليل.")
        else:
            with st.spinner("جاري صهر البيانات في القالب السيادي..."):
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel('gemini-1.5-pro')
                
                context = "\n".join([f"- {k}: {v}" for k, v in answers.items() if v])
                prompt = f"""
                بصفتك مستشاراً تنفيذياً خبيراً، صغ {report_type} بأسلوب سيادي، فخم، ومقتضب.
                استخدم لغة الأرقام والنتائج فقط. تجنب العبارات الإنشائية. 
                نظم الوثيقة في أقسام احترافية واضحة تعكس الجدية.
                البيانات المستخلصة:
                {context}
                """
                
                response = model.generate_content(prompt)
                report_text = response.text
                
                st.success("تم التوليد بنجاح.")
                st.info(report_text)
                
                # إعداد التنزيل
                file_data = create_docx(report_text, report_type)
                st.download_button(
                    label="تحميل الوثيقة الرسمية (Word)",
                    data=file_data,
                    file_name=f"{report_type.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
    except KeyError:
        st.error("خطأ في النظام: لم يتم العثور على مفتاح API في خزنة الأسرار (Secrets). يرجى إضافته من إعدادات المنصة السحابية.")
    except Exception as e:
        st.error(f"فشل في الاتصال بالمحرك الذكي: {e}")
