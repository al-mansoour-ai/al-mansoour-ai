import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document
from datetime import datetime

# 1. إعدادات الهوية البصرية الصارمة (Cairo + Premium Navy)
st.set_page_config(page_title="منصة المنصور AI - المستشار الذكي", layout="centered")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
    
    /* توحيد الخط والاتجاه */
    * { font-family: 'Cairo', sans-serif !important; direction: rtl; text-align: right; }
    .stApp { background-color: #f1f5f9; }
    #MainMenu, footer, header { visibility: hidden; }

    /* الحاوية الرئيسية - تصميم متباعد لمنع التداخل */
    .main-box { 
        background: white; border-top: 10px solid #0f172a; 
        padding: 30px; border-radius: 15px; 
        box-shadow: 0 10px 25px rgba(0,0,0,0.05); margin-top: -50px; 
    }
    
    .brand-title { color: #0f172a; font-weight: 900; font-size: 1.8rem; text-align: center; margin-bottom: 5px; }
    .brand-subtitle { color: #c5a059; text-align: center; font-weight: 700; font-size: 0.8rem; margin-bottom: 30px; }

    /* تنسيق العناوين الاستشارية */
    .section-header { 
        background: #0f172a; color: white; padding: 12px 18px; 
        border-radius: 8px; font-weight: 700; margin: 25px 0 15px 0; display: block;
        font-size: 1rem;
    }
    
    .q-label { color: #1e293b; font-weight: 800; border-right: 5px solid #c5a059; padding-right: 12px; margin-top: 20px; display: block; font-size: 1rem; }
    .hint-box { color: #64748b; font-size: 0.85rem; background: #fffbeb; padding: 12px; border-radius: 8px; border: 1px solid #fef3c7; margin-bottom: 15px; line-height: 1.6; }

    /* الأزرار الرسمية */
    .stButton>button { 
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%) !important; color: white !important; 
        font-weight: 700 !important; height: 50px !important; border-radius: 10px !important; width: 100%; border: none;
    }
    .magic-btn button { 
        background: #fdfaf3 !important; color: #856404 !important; border: 1px dashed #c5a059 !important; 
        height: 38px !important; font-size: 0.8rem !important; width: auto !important; margin-top: 5px;
    }
</style>
""", unsafe_allow_html=True)

# تفعيل Gemini
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")
except:
    st.error("⚠️ يرجى ضبط المفتاح في الإعدادات")

# --- بنك المعرفة الكامل (12 تخصصاً مع الأمثلة) ---
REPORTS_DB = {
    "📑 تقرير إنجاز دوري": [
        ("الملخص التنفيذي للأداء", "مثال: تم إنجاز 90% من المهام المخططة للفترة الحالية..."),
        ("تحليل الأنشطة المنفذة", "مثال: عقد 3 ورش تدريبية، وتوريد 50 وحدة تقنية..."),
        ("إدارة التحديات والحلول", "مثال: تأخر توريد المواد بسبب اللوجستيات وتم المعالجة..."),
        ("الخطة المستقبلية", "مثال: البدء في المرحلة الثانية من التدريب الميداني...")
    ],
    "💰 دراسة جدوى استثمارية": [
        ("تحليل الفجوة والاحتياج", "مثال: يوجد عجز بنسبة 20% في خدمات الطاقة بالمنطقة..."),
        ("النمذجة والتقديرات المالية", "مثال: العائد المتوقع يبدأ من السنة الثانية بنسبة 15%..."),
        ("تحليل SWOT والمنافسة", "مثال: القوة في التكنولوجيا المتاحة والتهديد في تقلب العملة...")
    ],
    "🎓 تقرير ختامي لتدريب": [
        ("الأهداف والمنهجية التدريبية", "مثال: تزويد المتدربين بمهارات الإدارة الاستراتيجية..."),
        ("نتائج التقييم القبلي والبعدي", "ارتفع مستوى المعرفة من 40% إلى 95% بعد الدورة...")
    ],
    "🔍 متابعة وتقييم (M&E)": [
        ("قياس KPIs ومؤشرات الأداء", "تم تحقيق 95% من مؤشر الحضور النوعي للمستفيدين..."),
        ("جودة المخرجات والامتثال", "تطابق جميع المخرجات مع معايير الجودة الدولية...")
    ],
    "🚑 تقييم احتياجات": [("وصف الاحتياج الراهن", "نقص حاد في مياه الشرب بالمديرية..."), ("ديموغرافيا المتضررين", "النساء والأطفال هم الأكثر تأثراً..."), ("خارطة التدخل", "البدء بإنشاء عيادة متنقلة...")],
    "🏛️ حوكمة وامتثال": [("الالتزام باللوائح", "مراجعة العقود وتطابقها مع قانون العمل..."), ("الثغرات وإجراءات التصحيح", "اعتماد نظام ERP جديد لضبط العمليات...")],
    "🏗️ فني وهندسي": [("المواصفات الفنية", "مطابقة المواد للكود الهندسي..."), ("السلامة المهنية", "تقرير الالتزام بأدوات السلامة الميدانية...")],
    "📝 تحليل مناقصات": [("التقييم الفني", "مقارنة العروض الفنية للموردين..."), ("توصية الترسية", "اختيار المورد الأنسب بناءً على الجودة والسرعة...")],
    "⚠️ إدارة مخاطر": [("سجل المخاطر", "تحديد المخاطر المالية والأمنية المحتملة..."), ("خطط الاستجابة", "خطة الطوارئ المعتمدة في حال الأزمات...")],
    "🌟 استراتيجي سنوي": [("المنجز العام", "حصاد عام كامل من المبادرات والنجاحات..."), ("أهداف العام القادم", "الرؤية المستقبلية للتوسع والنمو...")],
    "💰 أداء مالي": [("بيان المصروفات", "تحليل شامل لبنود الصرف..."), ("انحرافات الميزانية", "أسباب تجاوز الميزانية في بنود التشغيل...")],
    "🌍 أثر بيئي": [("الأثر الحيوي", "دراسة تأثير المشروع على البيئة المحلية..."), ("المسؤولية المجتمعية", "مستوى تقبل المجتمع المحلي للمشروع...")]
}

# --- جسم التطبيق السيادي ---
st.markdown('<div class="main-box">', unsafe_allow_html=True)
st.markdown('<h1 class="brand-title">منصة المنصور الاستراتيجية AI</h1>', unsafe_allow_html=True)
st.markdown('<p class="brand-subtitle">نظام الصياغة والتحليل المؤسسي الشامل - 2026</p>', unsafe_allow_html=True)

# 1. القائمة الجانبية (لمنع الزحمة في الواجهة)
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3135/3135715.png", width=65)
    rtype = st.selectbox("🎯 حدد تخصص التقرير:", list(REPORTS_DB.keys()))
    st.write("---")
    up_file = st.file_uploader("📂 ارفع مرجعاً أو مسودة (اختياري)", type=['pdf', 'docx', 'txt'])

# 2. بيانات الغلاف (بدون أعمدة نهائياً لضمان عدم التداخل)
st.markdown('<span class="section-header">🛡️ أولاً: صفحة الغلاف والبيانات الرسمية</span>', unsafe_allow_html=True)
p_name = st.text_input("عنوان التقرير (اسم المشروع) *")
p_agency = st.text_input("الجهة المُعِدّة (المؤسسة)")
p_donor = st.text_input("الجهة الموجه إليها (العميل)")
p_loc = st.text_input("المكان والنطاق الجغرافي")
p_ref = st.text_input("الرقم المرجعي (Ref No.)")
p_date = st.text_input("تاريخ الإصدار", value=datetime.now().strftime('%Y-%m-%d'))

# 3. المحاور الاستراتيجية (مع استعادة كل سؤال وزر التحسين)
st.markdown(f'<span class="section-header">🔍 ثانياً: محاور تقرير {rtype}</span>', unsafe_allow_html=True)
user_ans = {}
for i, (pillar, hint) in enumerate(REPORTS_DB[rtype]):
    st.markdown(f'<span class="q-label">{pillar}</span>', unsafe_allow_html=True)
    st.markdown(f'<div class="hint-box">💡 {hint}</div>', unsafe_allow_html=True)
    txt = st.text_area("", key=f"v17_{i}", height=120, label_visibility="collapsed")
    
    if st.button(f"✨ تحسين صياغة {pillar}", key=f"btn_v17_{i}"):
        if txt:
            with st.spinner("جاري المعالجة..."):
                res = model.generate_content(f"صغ هذا المحور بأسلوب استشاري فخم وصوت نشط: {txt}")
                st.info(res.text)
        else: st.warning("أدخل نصاً أولاً")
    user_ans[pillar] = txt

# 4. الإضافة المخصصة (المستقرة)
st.markdown('<span class="section-header">➕ ثالثاً: إضافة أقسام مخصصة</span>', unsafe_allow_html=True)
st.info("إذا كان لديك بيانات أخرى، اضغط على الزر وخصص قسماً جديداً:")
if 'extra_v17' not in st.session_state: st.session_state.extra_v17 = []

new_sec = st.text_input("اكتب اسم القسم الجديد هنا:")
if st.button("أنشئ القسم المخصص الآن"):
    if new_sec and new_sec not in st.session_state.extra_v17:
        st.session_state.extra_v17.append(new_sec)
        st.rerun()

for ex in st.session_state.extra_v17:
    st.markdown(f"**⭐ القسم المخصص: {ex}**")
    user_ans[ex] = st.text_area(f"أدخل بيانات {ex}...", key=f"ex17_{ex}")
    if st.button(f"حذف {ex}"): st.session_state.extra_v17.remove(ex); st.rerun()

# 5. الخاتمة والاعتماد
st.markdown('<span class="section-header">🖊️ رابعاً: هيكل الاعتماد والتوقيع</span>', unsafe_allow_html=True)
p_concl = st.text_area("الخاتمة والتوصيات الاستراتيجية:")
p_pre = st.text_input("إعداد:")
p_rev = st.text_input("مراجعة:")
p_app = st.text_input("اعتماد:")

st.write("---")
if st.button("🚀 توليد ومعالجة التقرير الاستراتيجي الشامل"):
    if p_name and any(user_ans.values()):
        with st.spinner("جاري صهر البيانات وفق معايير الجودة..."):
            summary = "\n".join([f"- {k}: {v}" for k, v in user_ans.items() if v])
            prompt = f"صغ تقريراً استراتيجياً لـ {p_name}. الجهة: {p_agency}. المحاور: {summary}. الخاتمة: {p_concl}. التوقيعات: {p_pre}, {p_rev}, {p_app}."
            res = model.generate_content(prompt)
            st.markdown("### المعاينة:")
            st.markdown(res.text)
            st.session_state['v17_out'] = res.text
    else: st.warning("يرجى ملء البيانات الأساسية.")

if 'v17_out' in st.session_state:
    doc = Document()
    doc.add_heading(p_name, 0)
    doc.add_paragraph(st.session_state['v17_out'])
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    st.download_button("💾 تحميل ملف Word المعتمد", bio, f"{p_name}.docx")

st.markdown('<center><a href="https://wa.me/967774575749" style="text-decoration:none; color:#25d366; font-weight:bold;">💬 الدعم الفني: 774575749</a></center>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)
