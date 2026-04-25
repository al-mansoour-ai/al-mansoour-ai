import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
import PyPDF2

# --- 1. الإعدادات البصرية المتقدمة (هوية المنصور الفاخرة) ---
st.set_page_config(page_title="منصة المنصور الاستراتيجية", layout="centered")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
    
    /* ضبط الخلفية والخط الأساسي */
    * { font-family: 'Cairo', sans-serif !important; direction: rtl; }
    .stApp { 
        background: radial-gradient(circle at top, #1e293b 0%, #0f172a 100%); 
        color: #f8fafc;
    }

    /* إخفاء العناصر الافتراضية لزيادة الاحترافية */
    #MainMenu, footer, header { visibility: hidden; }

    /* حاوية العنوان الرئيسي */
    .brand-container {
        text-align: center;
        padding: 40px 20px;
        background: rgba(255, 255, 255, 0.03);
        border-radius: 20px;
        border: 1px solid rgba(197, 160, 89, 0.2);
        margin-bottom: 40px;
        backdrop-filter: blur(10px);
    }
    .brand-title { 
        color: #C5A059; 
        font-weight: 900; 
        font-size: 2.8rem; 
        margin: 0;
        text-shadow: 0 4px 15px rgba(0,0,0,0.5);
    }
    .brand-subtitle { 
        color: #94a3b8; 
        font-weight: 400; 
        font-size: 1.1rem;
        margin-top: 10px;
    }

    /* تصميم البطاقات (Module Cards) */
    .module-card { 
        background: rgba(30, 41, 59, 0.7); 
        border-right: 6px solid #C5A059; 
        padding: 30px; 
        border-radius: 15px; 
        box-shadow: 0 10px 30px rgba(0,0,0,0.3); 
        margin-bottom: 30px;
    }
    
    .module-label { 
        background: #C5A059; 
        color: #0f172a; 
        padding: 8px 20px; 
        border-radius: 5px; 
        font-weight: 900; 
        font-size: 1rem; 
        margin-bottom: 20px; 
        display: inline-block;
    }

    /* نصوص الأسئلة والتلميحات */
    .q-text { color: #e2e8f0; font-weight: 700; font-size: 1.15rem; margin-bottom: 10px; display: block; }
    .hint-style { 
        color: #C5A059; font-size: 0.85rem; background: rgba(197, 160, 89, 0.05); 
        padding: 12px; border-radius: 10px; border: 1px solid rgba(197, 160, 89, 0.2); 
        margin-bottom: 15px; line-height: 1.6; 
    }

    /* حل مشكلة تداخل نصوص الرفع */
    [data-testid="stFileUploadDropzone"] button { opacity: 0 !important; position: relative; z-index: 2; }
    [data-testid="stFileUploadDropzone"] section::after {
        content: "📁 اسحب وأفلت الوثائق المرجعية هنا";
        position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%);
        color: #C5A059; font-weight: 700; z-index: 1; pointer-events: none; width: 100%; text-align: center;
    }

    /* الأزرار الذهبية */
    .stButton>button { 
        background: linear-gradient(90deg, #C5A059 0%, #9a7b41 100%) !important; 
        color: #0f172a !important; font-weight: 900 !important; height: 55px !important; 
        border-radius: 10px !important; border: none !important; width: 100%;
        box-shadow: 0 5px 15px rgba(197, 160, 89, 0.3) !important;
    }
    
    /* تنسيق حقول الإدخال */
    .stTextInput>div>div>input, .stTextArea>div>div>textarea, .stSelectbox>div>div>div {
        background-color: #1e293b !important;
        color: #ffffff !important;
        border: 1px solid rgba(197, 160, 89, 0.2) !important;
        border-radius: 8px !important;
    }
</style>
""", unsafe_allow_html=True)

# --- 2. المنطق البرمجي والذكاء الاصطناعي ---
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")
except:
    st.error("⚠️ يرجى ضبط مفتاح GEMINI_API_KEY في الإعدادات")

def extract_file_content(file):
    if file.type == "application/pdf":
        reader = PyPDF2.PdfReader(file)
        return " ".join([page.extract_text() for page in reader.pages])
    return file.getvalue().decode("utf-8")

# بنك التخصصات الاستراتيجية الـ 12
REPORTS_BANK = {
    "📑 تقرير إنجاز دوري (إداري)": [
        ("الملخص التنفيذي للأداء العام", "مثال: تم تحقيق 95% من الأهداف التشغيلية للربع الحالي..."),
        ("تحليل الأنشطة والمنجزات المحققة", "مثال: إنجاز نظام الأرشفة الإلكتروني، تدريب 20 موظفاً..."),
        ("إدارة الانحرافات والتحديات", "مثال: تأخر توريد الأجهزة وتم حله عبر المورد البديل..."),
        ("خطة العمل للفترة القادمة", "مثال: إطلاق المرحلة الثانية من التحول الرقمي...")
    ],
    "🎓 تقرير برنامج تدريبي ختامي": [
        ("بيانات المنهجية والمدرب", "مثال: استخدام منهجية التدريب التشاركي (Master Coach)..."),
        ("إحصائيات الحضور والنوع الاجتماعي", "مثال: إجمالي 40 مشاركاً (25 ذكور / 15 إناث)..."),
        ("قياس الأثر والتقييم القبلي/البعدي", "مثال: ارتفاع مهارات المشاركين بنسبة زيادة 60%..."),
        ("توصيات الاستدامة المهنية", "مثال: تفعيل مجموعات دعم الأقران لمتابعة التطبيق...")
    ],
    "🔍 متابعة وتقييم (M&E)": [
        ("مؤشرات التحقق KPIs", "مثال: الوصول لـ 500 مستفيد فعلي مقابل 450 مستهدف..."),
        ("الدروس المستفادة والجودة", "مثال: كفاءة التوزيع كانت عالية ولكن التوقيت يحتاج تحسين...")
    ],
    "💰 دراسة جدوى استثمارية": [
        ("تحليل الفجوة السوقية", "مثال: طلب متزايد بنسبة 20% سنوياً على خدمات الطاقة..."),
        ("التحليل المالي وتقدير المخاطر", "مثال: فترة استرداد رأس المال 18 شهراً مع هامش ربح 25%...")
    ],
    "🏗️ تقرير هندسي وفني": [("سير الأعمال والمطابقة", "مثال: مطابقة الخرسانة للكود العالمي..."), ("تقرير السلامة والمخاطر", "إجراءات الوقاية المتبعة في الموقع...")],
    "🏛️ حوكمة وامتثال": [("الالتزام المؤسسي", "مطابقة السياسات المالية للقانون المحلي..."), ("نتائج التدقيق الداخلي", "رصد تحسن ملحوظ في شفافية المشتريات...")],
    "🚑 تقييم احتياجات إنسانية": [("تحليل الوضع الميداني", "فجوة كبيرة في المياه والاصحاح البيئي..."), ("الاستجابة المقترحة", "توزيع سلال غذائية عاجلة لـ 200 أسرة...")],
    "💰 تقرير أداء مالي": [("السيولة والتدفقات", "تحليل المصروفات التشغيلية مقابل الميزانية..."), ("انحراف الموازنة", "تجاوز في بند الصيانة بنسبة 5%...")],
    "🌍 أثر بيئي واجتماعي": [("تحليل التأثيرات الجانبية", "تأثير المشروع على السكان المحليين..."), ("خطة التخفيف", "إجراءات الحد من التلوث الصوتي...")],
    "📝 تحليل مناقصات وترسية": [("التقييم الفني والمالي", "مقارنة العروض واختيار الأفضل معيارياً..."), ("توصية اللجنة", "ترسية المشروع على شركة (س) لخبرتها...")],
    "⚠️ إدارة مخاطر وطوارئ": [("سجل المخاطر المحدث", "ظهور مخاطر أمنية في طريق الإمداد..."), ("خطة الاستجابة", "تفعيل المسار البديل لتأمين الوصول...")],
    "🌟 تقرير استراتيجي سنوي": [("حصاد الرؤية والأهداف", "تحقيق 80% من الرؤية الخمسية للمؤسسة..."), ("التوجهات المستقبلية", "التحول الكامل نحو الذكاء الاصطناعي...")],
}

# --- 3. بناء الواجهة التنفيذية ---
st.markdown("""
<div class="brand-container">
    <h1 class="brand-title">منصة المنصور الاستراتيجية AI</h1>
    <p class="brand-subtitle">النظام الاستشاري المتكامل لصياغة التقارير والتحليل المؤسسي</p>
</div>
""", unsafe_allow_html=True)

# بطاقة 1: الإعداد
st.markdown('<div class="module-card">', unsafe_allow_html=True)
st.markdown('<div class="module-label">📁 الخطوة 1: المدخلات والمراجع</div>', unsafe_allow_html=True)
rtype = st.selectbox("اختر تخصص التقرير المطلوب:", list(REPORTS_BANK.keys()))
up_file = st.file_uploader("", type=['pdf', 'docx', 'txt'])
st.markdown('</div>', unsafe_allow_html=True)

# بطاقة 2: البيانات الرسمية
st.markdown('<div class="module-card">', unsafe_allow_html=True)
st.markdown('<div class="module-label">🛡️ الخطوة 2: الغلاف الرسمي</div>', unsafe_allow_html=True)
col1, col2 = st.columns(2)
with col1:
    p_name = st.text_input("عنوان التقرير / المشروع *")
    p_agency = st.text_input("الجهة المُعِدّة")
with col2:
    p_donor = st.text_input("الجهة الموجه إليها")
    p_loc = st.text_input("الموقع الجغرافي")
p_date = st.text_input("التاريخ المعتمد", value=datetime.now().strftime('%Y-%m-%d'))
st.markdown('</div>', unsafe_allow_html=True)

# بطاقة 3: المحاور التفصيلية
st.markdown(f"### 🔍 محاور تقرير {rtype}")
user_ans = {}
for i, (pillar, hint) in enumerate(REPORTS_BANK[rtype]):
    st.markdown('<div class="module-card">', unsafe_allow_html=True)
    st.markdown(f'<span class="q-text">{pillar}</span>', unsafe_allow_html=True)
    st.markdown(f'<div class="hint-style">💡 {hint}</div>', unsafe_allow_html=True)
    txt = st.text_area("", key=f"v24_{i}", height=150, label_visibility="collapsed", placeholder="أدخل البيانات الأساسية هنا ليقوم الذكاء الاصطناعي بتطويرها...")
    
    if st.button(f"✨ تحسين صياغة {pillar}", key=f"btn_v24_{i}"):
        if txt:
            with st.spinner("جاري الصياغة الاحترافية..."):
                res = model.generate_content(f"بصفتك مستشار إداري دولي، صغ هذا النص بأسلوب رصين وفخم: {txt}")
                st.success(res.text)
        else: st.warning("يرجى كتابة نص أولاً لتحسينه.")
    user_ans[pillar] = txt
    st.markdown('</div>', unsafe_allow_html=True)

# بطاقة 4: التوقيعات والتوليد
st.markdown('<div class="module-card">', unsafe_allow_html=True)
st.markdown('<div class="module-label">🖊️ الخطوة 4: الاعتماد النهائي</div>', unsafe_allow_html=True)
c1, c2, c3 = st.columns(3)
with c1: p_pre = st.text_input("إعداد:")
with c2: p_rev = st.text_input("مراجعة:")
with c3: p_app = st.text_input("اعتماد:")

if st.button("🚀 توليد ومعالجة التقرير الاستراتيجي الشامل"):
    if p_name and any(user_ans.values()):
        with st.spinner("جاري تحليل البيانات ودمج المراجع..."):
            file_context = extract_file_content(up_file) if up_file else "لا توجد مراجع إضافية."
            summary = "\n".join([f"محور {k}: {v}" for k, v in user_ans.items() if v])
            
            mega_prompt = f"""
            أنت خبير استشارات إدارية دولي متخصص في كتابة التقارير.
            المطلوب: صياغة تقرير استراتيجي كامل بعنوان ({p_name}).
            الجهة المعدة: {p_agency} | الموجه إليه: {p_donor} | الموقع: {p_loc}
            
            البيانات المدخلة:
            {summary}
            
            سياق المراجع المرفقة:
            {file_context[:1500]}
            
            التعليمات:
            1. استخدم لغة عربية فصحى فخمة.
            2. اجعل التقرير هيكلياً (عناوين، نقاط، توصيات).
            3. أضف فقرة "الخاتمة والتوصيات الاستراتيجية" في النهاية.
            4. ضع أسماء المعتمدين في نهاية التقرير: {p_pre}, {p_rev}, {p_app}.
            """
            
            res = model.generate_content(mega_prompt)
            st.session_state['final_report'] = res.text
            st.markdown("---")
            st.markdown(f'<div style="background: white; color: #1e293b; padding: 40px; border-radius: 10px; border: 2px solid #C5A059; line-height: 2;">{res.text}</div>', unsafe_allow_html=True)
    else:
        st.error("يرجى التأكد من كتابة اسم التقرير وملء محور واحد على الأقل.")
st.markdown('</div>', unsafe_allow_html=True)

# تصدير الملف
if 'final_report' in st.session_state:
    doc = Document()
    p = doc.add_heading(p_name, 0)
    doc.add_paragraph(st.session_state['final_report'])
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    st.download_button("💾 تحميل ملف Word المعتمد", bio, f"{p_name}.docx")

# التذييل (WhatsApp Support)
st.markdown(f'<br><center><a href="https://wa.me/967774575749" style="text-decoration:none; color:#C5A059; font-weight:900; font-size:1.1rem;">💬 الدعم الفني المباشر (المنصور): 774575749</a></center>', unsafe_allow_html=True)
