import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document
from datetime import datetime

# 1. إعدادات الصفحة والهوية البصرية (Cairo + Navy & Gold)
st.set_page_config(page_title="منصة المنصور AI", layout="centered")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
    
    /* توحيد الخط والاتجاه */
    * { font-family: 'Cairo', sans-serif !important; direction: rtl; text-align: right; }
    .stApp { background-color: #f0f4f8; }
    #MainMenu, footer, header { visibility: hidden; }

    /* الحاوية الرئيسية الواضحة */
    .report-container { 
        background: white; border-top: 10px solid #1e3a8a; 
        padding: 30px; border-radius: 15px; 
        box-shadow: 0 10px 25px rgba(0,0,0,0.1); margin-top: -50px; 
    }
    
    .main-title { color: #1e3a8a; font-weight: 900; font-size: 2rem; text-align: center; margin-bottom: 5px; }
    .main-subtitle { color: #d4af37; text-align: center; font-weight: 700; font-size: 0.9rem; margin-bottom: 30px; }

    /* تنسيق العناوين لمنع التداخل */
    .section-label { 
        background: #1e3a8a; color: white; padding: 12px; 
        border-radius: 8px; font-weight: 700; margin: 25px 0 15px 0; display: block;
    }
    
    .question-title { color: #1e293b; font-weight: 800; border-right: 5px solid #d4af37; padding-right: 12px; margin-top: 25px; display: block; }
    .example-hint { color: #64748b; font-size: 0.8rem; background: #fffbeb; padding: 10px; border-radius: 8px; border: 1px solid #fef3c7; margin-bottom: 10px; line-height: 1.6; }

    /* الأزرار الاحترافية */
    .stButton>button { 
        background: linear-gradient(135deg, #1e3a8a 0%, #152e6d 100%) !important; color: white !important; 
        font-weight: 700 !important; height: 55px !important; border-radius: 10px !important; width: 100%; border: none;
    }
    
    .magic-btn button { 
        background: #fdfaf3 !important; color: #856404 !important; border: 1px dashed #d4af37 !important; 
        height: 38px !important; font-size: 0.8rem !important; width: auto !important;
    }
</style>
""", unsafe_allow_html=True)

# تفعيل Gemini
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")
except:
    st.error("⚠️ خطأ في الإعدادات: يرجى التأكد من إضافة GEMINI_API_KEY")

# --- بنك الـ 12 تخصصاً كاملاً (بدون نقص حرف واحد) ---
REPORTS_DB = {
    "📑 تقرير إنجاز دوري": [
        ("الملخص التنفيذي للأداء", "مثال: تم إنجاز 90% من المهام المخططة..."),
        ("تحليل الأنشطة المنفذة", "مثال: عقد 3 ورش تدريبية، وتوريد 50 وحدة..."),
        ("إدارة الانحرافات", "مثال: تأخر التوريد بسبب اللوجستيات وتم المعالجة..."),
        ("التحديات والحلول", "مثال: ضعف الإنترنت في المواقع والحل كان Starlink..."),
        ("الخطة المستقبلية", "مثال: البدء في المرحلة الثانية من التدريب...")
    ],
    "💰 دراسة جدوى استثمارية": [
        ("تحليل الفجوة والاحتياج", "مثال: يوجد عجز بنسبة 20% في خدمات المنطقة..."),
        ("المتطلبات الفنية", "مثال: نحتاج خوادم بسعة 10 تيرابايت..."),
        ("النمذجة المالية", "مثال: العائد المتوقع على الاستثمار هو 15%..."),
        ("تحليل SWOT والمنافسة", "مثال: القوة في التكنولوجيا والتهديد في الصرف..."),
        ("قرار الاستثمار النهائي", "بناءً على الأرقام، المشروع مجدٍ اقتصادياً...")
    ],
    "🎓 تقرير ختامي لتدريب": [
        ("الأهداف والمنهجية", "مثال: تزويد 50 متدرباً بمهارات القيادة..."),
        ("نتائج التقييم القبلي والبعدي", "ارتفع مستوى المهارة من 30% إلى 85%..."),
        ("تقييم المدرب واللوجستيات", "حصل المدرب على تقييم 4.9/5..."),
        ("توصيات استدامة الأثر", "إنشاء مجموعة متابعة لتبادل الخبرات...")
    ],
    "🔍 متابعة وتقييم (M&E)": [
        ("قياس مؤشرات الأداء KPIs", "تم تحقيق 95% من مؤشر الحضور النوعي..."),
        ("جودة المخرجات والامتثال", "تطابق المخرجات مع معايير الجودة الدولية..."),
        ("رضا المستفيدين والدروس", "أظهرت الاستبيانات رضا بنسبة 92%...")
    ],
    "🚑 تقييم احتياجات": [("وصف الاحتياج", "نقص حاد في مياه الشرب..."), ("ديموغرافيا المتضررين", "النساء والأطفال هم الأكثر تأثراً..."), ("خارطة التدخل", "البدء بإنشاء عيادة متنقلة...")],
    "🏛️ حوكمة وامتثال": [("الالتزام باللوائح", "مراجعة العقود وتطابقها مع القانون..."), ("الثغرات وإجراءات التصحيح", "اعتماد نظام ERP جديد...")],
    "💰 أداء مالي": [("بيان المصروفات", "تحليل شامل لبنود الصرف..."), ("انحرافات الميزانية", "أسباب تجاوز الميزانية...")],
    "🏗️ فني وهندسي": [("المواصفات الفنية", "مطابقة المواد للكود الهندسي..."), ("السلامة المهنية", "الالتزام بأدوات السلامة...")],
    "🌍 أثر بيئي": [("الأثر الحيوي", "تأثير المشروع على البيئة..."), ("المسؤولية المجتمعية", "مستوى تقبل المجتمع...")],
    "📝 تحليل مناقصات": [("التقييم الفني", "مقارنة العروض الفنية..."), ("توصية الترسية", "اختيار المورد الأنسب...")],
    "⚠️ إدارة مخاطر": [("سجل المخاطر", "تحديد المخاطر الأمنية..."), ("خطط الاستجابة", "خطة الطوارئ المعتمدة...")],
    "🌟 استراتيجي سنوي": [("المنجز العام", "حصاد عام كامل من النجاحات..."), ("أهداف العام القادم", "الرؤية المستقبلية للتوسع...")]
}

# --- جسم التطبيق المنظم ---
st.markdown('<div class="report-container">', unsafe_allow_html=True)
st.markdown('<h1 class="main-title">منصة المنصور الاستراتيجية AI</h1>', unsafe_allow_html=True)
st.markdown('<p class="main-subtitle">النظام الاستشاري الشامل - إصدار الاستقرار 2026</p>', unsafe_allow_html=True)

# 1. اختيار التخصص ورفع الملف
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3135/3135715.png", width=70)
    rtype = st.selectbox("🎯 نوع التقرير المطلوب:", list(REPORTS_DB.keys()))
    st.write("---")
    up_file = st.file_uploader("📂 ارفع مرجعاً (اختياري)", type=['pdf', 'docx', 'txt'])

# 2. بيانات الغلاف (بدون أعمدة لمنع التداخل)
st.markdown('<span class="section-label">🛡️ أولاً: بيانات الغلاف الرسمي</span>', unsafe_allow_html=True)
p_name = st.text_input("عنوان التقرير (اسم المشروع) *")
p_agency = st.text_input("الجهة المُعِدّة (المؤسسة)")
p_donor = st.text_input("الجهة الموجه إليها (العميل)")
p_loc = st.text_input("المكان والنطاق الجغرافي")
p_date = st.text_input("التاريخ", value=datetime.now().strftime('%Y-%m-%d'))

# 3. المحاور (استعادة كل الأسئلة وزر التحسين)
st.markdown(f'<span class="section-label">🔍 ثانياً: محاور {rtype}</span>', unsafe_allow_html=True)
user_ans = {}
for i, (pillar, hint) in enumerate(REPORTS_DB[rtype]):
    st.markdown(f'<span class="question-title">{pillar}</span>', unsafe_allow_html=True)
    st.markdown(f'<div class="example-hint">💡 {hint}</div>', unsafe_allow_html=True)
    txt = st.text_area("", key=f"v16_{i}", height=120, label_visibility="collapsed")
    
    if st.button(f"✨ تحسين صياغة {pillar}", key=f"btn_v16_{i}"):
        if txt:
            with st.spinner("جاري التحسين..."):
                res = model.generate_content(f"صغ هذا المحور بأسلوب استشاري فخم وصوت نشط: {txt}")
                st.info(res.text)
        else: st.warning("أدخل نصاً أولاً")
    user_ans[pillar] = txt

# 4. الإضافة المخصصة (تصحيح الخطأ البرمجي)
st.markdown('<span class="section-label">➕ ثالثاً: إضافة أقسام مخصصة</span>', unsafe_allow_html=True)
st.info("إذا كان لديك بيانات أخرى، اضغط على الزر وخصص قسماً جديداً:")
if 'extra_v16' not in st.session_state: st.session_state.extra_v16 = []

new_sec = st.text_input("اسم القسم الإضافي الجديد:")
if st.button("أنشئ القسم المخصص"):
    if new_sec and new_sec not in st.session_state.extra_v16:
        st.session_state.extra_v16.append(new_sec)
        st.rerun()

for ex in st.session_state.extra_v16:
    st.markdown(f"**⭐ القسم المخصص: {ex}**")
    user_ans[ex] = st.text_area(f"بيانات {ex}...", key=f"ex16_{ex}")
    if st.button(f"حذف {ex}"): st.session_state.extra_v16.remove(ex); st.rerun()

# 5. التوقيعات
st.markdown('<span class="section-label">🖊️ رابعاً: هيكل الاعتماد والتوقيع</span>', unsafe_allow_html=True)
p_pre = st.text_input("إعداد:")
p_rev = st.text_input("مراجعة:")
p_app = st.text_input("اعتماد:")

st.write("---")
if st.button("🚀 توليد ومعالجة التقرير الاستراتيجي الشامل"):
    if p_name and any(user_ans.values()):
        with st.spinner("جاري صهر البيانات استراتيجياً..."):
            summary = "\n".join([f"- {k}: {v}" for k, v in user_ans.items() if v])
            prompt = f"صغ تقريراً استراتيجياً لـ {p_name}. الجهة: {p_agency}. المحاور: {summary}. التوقيعات: {p_pre}, {p_rev}, {p_app}. التزم بترقيم ISO 2145."
            res = model.generate_content(prompt)
            st.markdown(res.text)
            st.session_state['v16_out'] = res.text
    else: st.warning("يرجى ملء البيانات الأساسية.")

if 'v16_out' in st.session_state:
    doc = Document()
    doc.add_heading(p_name, 0)
    doc.add_paragraph(st.session_state['v16_out'])
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    st.download_button("💾 تحميل ملف Word المعتمد", bio, f"{p_name}.docx")

st.markdown('<center><a href="https://wa.me/967774575749" style="text-decoration:none; color:#25d366; font-weight:bold;">💬 تواصل معنا للدعم: 774575749</a></center>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)
