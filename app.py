import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document
from datetime import datetime

# 1. الهندسة البصرية وتوحيد الهوية (Cairo + Navy & Gold)
st.set_page_config(page_title="منصة المنصور الاستراتيجية AI", layout="centered")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
    
    /* توحيد الخط والاتجاه */
    * { font-family: 'Cairo', sans-serif !important; direction: rtl; text-align: right; }
    .stApp { background-color: #f1f5f9; }
    #MainMenu, footer, header { visibility: hidden; }

    /* الحاوية الرئيسية (لمنع التداخل) */
    .main-card { 
        background: white; border-top: 10px solid #0f172a; 
        padding: 30px; border-radius: 15px; 
        box-shadow: 0 10px 30px rgba(0,0,0,0.05); margin-top: -50px; 
    }
    
    .brand-title { color: #0f172a; font-weight: 900; font-size: 2rem; text-align: center; margin-bottom: 5px; }
    .brand-subtitle { color: #c5a059; text-align: center; font-weight: 700; font-size: 0.9rem; margin-bottom: 30px; }

    /* العناوين الفرعية الرشيقة */
    .section-header { 
        background: #0f172a; color: white; padding: 12px 20px; 
        border-radius: 8px; font-weight: 700; font-size: 1.1rem; 
        margin: 25px 0 15px 0; display: block;
    }
    
    .q-label { color: #1e293b; font-weight: 800; border-right: 5px solid #c5a059; padding-right: 12px; margin-top: 20px; display: block; }
    .hint-box { color: #64748b; font-size: 0.82rem; background: #fdfaf3; padding: 12px; border-radius: 8px; border: 1px solid #f1e6d0; margin-bottom: 12px; line-height: 1.6; }

    /* الأزرار الملكية */
    .stButton>button { 
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%) !important; color: white !important; 
        font-weight: 700 !important; height: 55px !important; border-radius: 10px !important; 
        border: none !important; width: 100%; transition: 0.3s;
    }
    
    .magic-btn button { 
        background: #fff9db !important; color: #856404 !important; border: 1px dashed #fab005 !important; 
        height: 38px !important; font-size: 0.8rem !important; margin-top: 5px; width: auto !important;
    }

    .whatsapp-btn { background: #25d366; color: white !important; padding: 12px 25px; border-radius: 50px; text-decoration: none; font-weight: 700; display: inline-flex; align-items: center; gap: 10px; margin-top: 20px; }
</style>
""", unsafe_allow_html=True)

# تفعيل Gemini
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")
except:
    st.error("⚠️ يرجى ضبط المفتاح في الإعدادات")

# --- بنك الـ 12 تخصصاً كاملاً (نفس المحاور السابقة دون حذف) ---
STRATEGY_BANK = {
    "📑 تقرير إنجاز دوري": [
        ("الملخص التنفيذي للأداء", "مثال: تم إنجاز 90% من المهام المخططة للفترة الحالية بنجاح..."),
        ("تحليل الأنشطة والمهام المنفذة", "مثال: عقد 3 ورش تدريبية، وتوريد 50 وحدة تقنية..."),
        ("إدارة الانحرافات والجدول الزمني", "مثال: تأخر توريد المواد بسبب اللوجستيات وتم تعويضه بالعمل الإضافي..."),
        ("التحديات والحلول التصحيحية", "مثال: ضعف الإنترنت في المواقع، والحل كان توفير أجهزة Starlink..."),
        ("الخطة المستقبلية", "مثال: البدء في المرحلة الثانية من التدريب التقني الميداني...")
    ],
    "💰 دراسة جدوى استثمارية": [
        ("تحليل الفجوة والاحتياج السوقي", "مثال: يوجد عجز بنسبة 20% في خدمات المنطقة..."),
        ("المتطلبات الفنية", "مثال: نحتاج إلى خوادم بسعة 10 تيرابايت ونظام حماية متطور..."),
        ("النمذجة المالية وتوقعات الدخل", "مثال: العائد المتوقع على الاستثمار (ROI) هو 15% سنوياً..."),
        ("تحليل SWOT والمنافسة", "مثال: القوة في التكنولوجيا، والتهديد في تقلب الصرف..."),
        ("قرار الاستثمار النهائي", "مثال: بناءً على الأرقام، المشروع مجدٍ اقتصادياً ويُنصح بالبدء...")
    ],
    "🎓 تقرير ختامي لتدريب": [
        ("الأهداف والمنهجية التدريبية", "تزويد المتدربين بمهارات القيادة الاستراتيجية..."),
        ("نتائج التقييم القبلي والبعدي", "ارتفع مستوى المهارة من 30% إلى 85%..."),
        ("تقييم المدرب واللوجستيات", "حصل المدرب على تقييم 4.9/5 من قبل المشاركين..."),
        ("توصيات استدامة الأثر", "إنشاء مجموعة متابعة (WhatsApp) لتبادل الخبرات المستمرة...")
    ],
    "🔍 متابعة وتقييم (M&E)": [
        ("قياس KPIs ومؤشرات الأداء", "تم تحقيق 95% من مؤشر الحضور النوعي للمستفيدين..."),
        ("جودة المخرجات والامتثال", "تطابق جميع المخرجات مع معايير الجودة (ISO)..."),
        ("رضا المستفيدين والدروس", "أظهرت الاستبيانات رضا بنسبة 92% عن سرعة الاستجابة...")
    ],
    "🚑 تقييم احتياجات": [
        ("وصف الاحتياج الراهن", "تفتقر المنطقة لمركز صحي متكامل يخدم 5000 نسمة..."),
        ("الفئات المتضررة ديموغرافياً", "الأطفال والنساء في القرى النائية هم الأكثر تضرراً..."),
        ("خارطة التدخل المقترحة", "المقترح البدء بإنشاء عيادة متنقلة كمرحلة أولى...")
    ],
    "🏛️ حوكمة وامتثال": [
        ("الالتزام باللوائح والسياسات", "تمت مراجعة العقود وتطابقها مع قانون العمل..."),
        ("الثغرات المرصودة وإجراءات التصحيح", "اعتماد نظام (ERP) جديد لضبط العمليات الإدارية...")
    ],
    "💰 أداء مالي": [("بيان المصروفات", "تحليل شامل لبنود الصرف..."), ("انحرافات الميزانية", "أسباب تجاوز الميزانية في بند النقل...")],
    "🏗️ فني وهندسي": [("المواصفات الفنية", "مطابقة الخرسانة للمواصفات..."), ("السلامة المهنية", "تقرير الالتزام بأدوات السلامة...")],
    "🌍 أثر بيئي": [("الأثر الحيوي", "دراسة تأثير المشروع على المياه الجوفية..."), ("المسؤولية المجتمعية", "مستوى تقبل المجتمع للمشروع...")],
    "📝 تحليل مناقصات": [("التقييم الفني", "مقارنة العروض الفنية للموردين..."), ("توصية الترسية", "اختيار المورد الأنسب بناءً على الجودة...")],
    "⚠️ إدارة مخاطر": [("سجل المخاطر", "تحديد المخاطر الأمنية والمالية..."), ("خطط الاستجابة", "خطة الطوارئ المعتمدة...")],
    "🌟 استراتيجي سنوي": [("المنجز العام", "حصاد عام كامل من النجاحات..."), ("أهداف العام القادم", "الرؤية المستقبلية للتوسع...")]
}

# --- الواجهة الرئيسية ---
st.markdown('<div class="main-card">', unsafe_allow_html=True)
st.markdown('<h1 class="brand-title">منصة المنصور الاستراتيجية AI</h1>', unsafe_allow_html=True)
st.markdown('<p class="brand-subtitle">نظام الصياغة والتحليل المؤسسي الشامل - إصدار 2026</p>', unsafe_allow_html=True)

# 1. الاختيار ورفع الوثائق
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3135/3135715.png", width=70)
    rtype = st.selectbox("🎯 نوع التقرير الاستراتيجي:", list(STRATEGY_BANK.keys()))
    st.write("---")
    uploaded_file = st.file_uploader("📂 ارفع مرجعاً (اختياري)", type=['pdf', 'docx', 'txt'])

# 2. بيانات الغلاف (تم إلغاء الأعمدة لمنع تداخل النصوص)
st.markdown('<div class="section-header">🛡️ أولاً: بيانات الغلاف والبيانات الرسمية</div>', unsafe_allow_html=True)
p_name = st.text_input("عنوان التقرير (اسم المشروع) *")
p_ref = st.text_input("الرقم المرجعي (Ref No.)")
p_agency = st.text_input("الجهة المُعِدّة (المؤسسة)")
p_donor = st.text_input("الجهة الموجه إليها (العميل)")
p_loc = st.text_input("المكان والنطاق الجغرافي")
p_date = st.text_input("تاريخ الإصدار", value=datetime.now().strftime('%Y-%m-%d'))

# 3. المقدمة والشكر
with st.expander("🤝 ثانياً: خطاب الإرسال والشكر والمقدمة"):
    p_thanks = st.text_area("أدخل نص الشكر أو مقدمة التقرير هنا...")

# 4. محاور التقرير (مع الأمثلة وزر التحسين)
st.markdown(f'<div class="section-header">🔍 ثالثاً: محاور {rtype}</div>', unsafe_allow_html=True)
responses = {}
current_pillars = STRATEGY_BANK.get(rtype, [])

for i, (pillar, hint) in enumerate(current_pillars):
    st.markdown(f'<span class="q-label">{pillar}</span>', unsafe_allow_html=True)
    st.markdown(f'<div class="hint-box">💡 {hint}</div>', unsafe_allow_html=True)
    txt = st.text_area("", key=f"f_v15_{i}", height=120, label_visibility="collapsed")
    
    # زر التحسين المستعاد
    if st.button(f"✨ تحسين صياغة {pillar}", key=f"btn_v15_{i}"):
        if txt:
            with st.spinner("جاري التحسين..."):
                res = model.generate_content(f"صغ هذا المحور بأسلوب استشاري فخم وصوت نشط: {txt}")
                st.info(res.text)
        else: st.warning("أدخل نصاً أولاً")
    responses[pillar] = txt

# 5. الإضافات المخصصة (المستعادة كما طلبت)
st.markdown('<div class="section-header">➕ رابعاً: إضافة أقسام مخصصة</div>', unsafe_allow_html=True)
st.info("إذا كان لديك بيانات أخرى وتريد إضافتها، اضغط على الزر وخصص قسماً جديداً:")
if 'extra_v15' not in st.session_state: st.session_state.extra_v15 = []

new_sec = st.text_input("اسم القسم الإضافي الذي تود إنشاءه:")
if st.button("أنشئ القسم المخصص الآن"):
    if new_sec and new_sec not in st.session_state.extra_v15:
        st.session_state.extra_v15.append(new_sec)
        st.rerun()

for ex in st.session_state.extra_v15:
    st.markdown(f"**⭐ القسم المخصص: {ex}**")
    responses[ex] = st.text_area(f"بيانات {ex}...", key=f"ex15_{ex}")
    if st.button(f"حذف {ex}"): st.session_state.extra_v15.remove(ex); st.rerun()

# 6. الخاتمة والتوقيعات
with st.expander("📌 خامساً: الخاتمة والاعتماد والتوقيع"):
    p_concl = st.text_area("الخاتمة والتوصيات الاستراتيجية:")
    p_pre = st.text_input("إعداد:")
    p_rev = st.text_input("مراجعة:")
    p_app = st.text_input("اعتماد:")

# 7. التوليد النهائي
st.write("---")
if st.button("🚀 توليد ومعالجة التقرير النهائي الشامل"):
    if p_name and any(responses.values()):
        with st.spinner("جاري الصهر والتحليل الاستراتيجي..."):
            all_txt = "\n".join([f"{k}: {v}" for k, v in responses.items() if v])
            prompt = f"صغ تقريراً استراتيجياً لـ {p_name}. الجهة: {p_agency}. المحاور: {all_txt}. الخاتمة: {p_concl}. التوقيعات: {p_pre}, {p_rev}, {p_app}."
            res = model.generate_content(prompt)
            st.markdown("### المعاينة:")
            st.markdown(res.text)
            st.session_state['v15_out'] = res.text
    else: st.warning("يرجى ملء البيانات الأساسية.")

if 'v15_out' in st.session_state:
    doc = Document()
    doc.add_heading(p_name, 0)
    doc.add_paragraph(st.session_state['v15_out'])
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    st.download_button("💾 تحميل ملف Word المعتمد", bio, f"{p_name}.docx")

st.markdown('<center><a href="https://wa.me/967774575749" class="whatsapp-btn">💬 تواصل معنا للدعم الفني: 774575749</a></center>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)
