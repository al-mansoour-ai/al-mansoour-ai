import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document
from datetime import datetime

# 1. الهندسة البصرية وتنسيق الخطوط (Cairo)
st.set_page_config(page_title="منصة المنصور الاستراتيجية AI", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
    * { font-family: 'Cairo', sans-serif !important; direction: rtl; text-align: right; }
    .stApp { background-color: #f0f4f8; }
    
    /* تصميم البطاقات المؤسسية */
    .card { 
        background: white; border-radius: 12px; padding: 25px; 
        border-right: 8px solid #1e3a8a; box-shadow: 0 4px 15px rgba(0,0,0,0.05);
        margin-bottom: 25px; overflow: hidden;
    }
    
    .header-box { background: #1e3a8a; color: white; padding: 25px; border-radius: 15px; text-align: center; margin-bottom: 40px; }
    .header-title { font-weight: 900; font-size: 2.2rem; margin: 0; }
    
    .q-label { color: #1e3a8a; font-weight: 800; font-size: 1.15rem; margin-bottom: 8px; display: block; }
    .example-hint { color: #64748b; font-size: 0.85rem; background: #fdfaf3; padding: 12px; border-radius: 8px; border: 1px solid #f1e6d0; margin-bottom: 12px; line-height: 1.6; }

    /* تحسين الأزرار */
    .stButton>button { 
        background: linear-gradient(90deg, #1e3a8a, #152e6d) !important; color: white !important; 
        font-weight: 700 !important; border-radius: 10px !important; width: 100%; height: 50px; border: none;
    }
    .magic-btn button { 
        background: #fff9db !important; color: #856404 !important; border: 1px dashed #fab005 !important; 
        height: 38px !important; font-size: 0.85rem !important; width: auto !important; margin-top: 5px;
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

# بنك الـ 12 تخصصاً كاملاً مع الأمثلة (بدون حذف)
STRATEGY_DATA = {
    "📑 تقرير إنجاز دوري": [
        ("الملخص التنفيذي للأداء", "مثال: تم إنجاز 90% من المهام المخططة للفترة الحالية بنجاح..."),
        ("تحليل الأنشطة المنفذة", "مثال: عقد 3 ورش تدريبية، وتوريد 50 وحدة تقنية..."),
        ("إدارة الانحرافات", "مثال: تأخر توريد المواد بسبب اللوجستيات وتم تعويضه بالعمل الإضافي..."),
        ("التحديات والحلول", "مثال: ضعف الإنترنت في المواقع، والحل كان توفير أجهزة Starlink..."),
        ("الخطة المستقبلية", "مثال: البدء في المرحلة الثانية من التدريب التقني الميداني...")
    ],
    "💰 دراسة جدوى استثمارية": [
        ("تحليل الفجوة والاحتياج", "مثال: يوجد عجز بنسبة 20% في تغطية الخدمات الرقمية في المنطقة..."),
        ("المتطلبات الفنية", "مثال: نحتاج إلى خوادم بسعة 10 تيرابايت ونظام حماية متطور..."),
        ("النمذجة المالية", "مثال: العائد المتوقع على الاستثمار (ROI) هو 15% سنوياً..."),
        ("تحليل SWOT", "مثال: القوة تكمن في قلة المنافسين، والضعف في نقص الكوادر المحلية..."),
        ("قرار الاستثمار", "مثال: بناءً على الأرقام، المشروع مجدٍ اقتصادياً ويُنصح بالبدء...")
    ],
    "🎓 تقرير ختامي لتدريب": [
        ("الأهداف والمنهجية", "تزويد 50 متدرباً بمهارات الذكاء الاصطناعي التوليدي..."),
        ("نتائج القبلي والبعدي", "ارتفع مستوى المهارة من 30% قبل التدريب إلى 85% بعده..."),
        ("تقييم المدرب واللوجستيات", "حصل المدرب على تقييم 4.9/5 من قبل المشاركين..."),
        ("توصيات الاستدامة", "إنشاء مجموعة متابعة (WhatsApp) لتبادل الخبرات المستمرة...")
    ],
    "🔍 متابعة وتقييم (M&E)": [
        ("قياس KPIs", "تم تحقيق 95% من مؤشر الحضور النوعي للمستفيدين..."),
        ("جودة المخرجات", "تطابق جميع المخرجات مع معايير الجودة (ISO)..."),
        ("رضا المستفيدين", "أظهرت الاستبيانات رضا بنسبة 92% عن سرعة الاستجابة...")
    ],
    "🚑 تقييم احتياجات": [
        ("وصف الاحتياج الراهن", "تفتقر المنطقة لمركز صحي متكامل يخدم 5000 نسمة..."),
        ("الفئات المتضررة", "الأطفال والنساء في القرى النائية هم الأكثر تضرراً..."),
        ("خارطة التدخل", "المقترح البدء بإنشاء عيادة متنقلة كمرحلة أولى...")
    ],
    "🏛️ حوكمة وامتثال": [
        ("الالتزام باللوائح", "تمت مراجعة جميع العقود وتطابقها مع قانون العمل..."),
        ("الثغرات المرصودة", "يوجد ضعف في نظام الأرشفة الرقمية للبيانات المالية..."),
        ("إجراءات التصحيح", "اعتماد نظام (ERP) جديد لضبط العمليات الإدارية...")
    ],
    "💰 أداء مالي": [("بيان المصروفات", "تحليل شامل لبنود الصرف خلال الربع الأول..."), ("انحرافات الميزانية", "أسباب تجاوز الميزانية في بند النقل...")],
    "🏗️ فني وهندسي": [("المواصفات الفنية", "مطابقة الخرسانة للمواصفات القياسية..."), ("السلامة المهنية", "تقرير حول الالتزام بارتداء أدوات السلامة...")],
    "🌍 أثر بيئي": [("الأثر الحيوي", "دراسة تأثير المشروع على المياه الجوفية..."), ("المسؤولية المجتمعية", "مستوى تقبل المجتمع المحلي للمشروع...")],
    "📝 تحليل مناقصات": [("التقييم الفني", "مقارنة العروض الفنية للموردين..."), ("توصية الترسية", "اختيار المورد الأنسب بناءً على السعر والجودة...")],
    "⚠️ إدارة مخاطر": [("سجل المخاطر", "تحديد المخاطر الأمنية والمالية المحتملة..."), ("خطط الاستجابة", "خطة الطوارئ في حال تعطل الإمدادات...")],
    "🌟 استراتيجي سنوي": [("المنجز العام", "حصاد عام كامل من المبادرات والنجاحات..."), ("أهداف العام القادم", "الرؤية المستقبلية للتوسع والنمو...")]
}

# --- القائمة الجانبية ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3135/3135715.png", width=70)
    st.markdown("### الإعدادات والتحليل")
    rtype = st.selectbox("🎯 نوع التقرير الاستراتيجي:", list(STRATEGY_DATA.keys()))
    st.write("---")
    uploaded_file = st.file_uploader("📂 ارفع مرجعاً أو مسودة (اختياري)", type=['pdf', 'docx', 'txt'])
    if uploaded_file: st.success("تم الربط مع الوثيقة")

# --- الواجهة الرئيسية ---
st.markdown('<div class="header-box"><h1 class="header-title">منصة المنصور الاستراتيجية AI</h1><p>النظام المؤسسي المتكامل لصناعة التقارير السيادية</p></div>', unsafe_allow_html=True)

# 1. بيانات الغلاف
st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown('<span class="q-label">🛡️ أولاً: بيانات الغلاف الرسمي</span>', unsafe_allow_html=True)
c1, c2 = st.columns(2)
p_name = c1.text_input("عنوان التقرير (اسم المشروع) *")
p_ref = c2.text_input("الرقم المرجعي (Ref No.)")
p_agency = c1.text_input("الجهة المُعِدّة (المؤسسة)")
p_donor = c2.text_input("الجهة الموجه إليها (المانح/العميل)")
p_loc = c1.text_input("المكان والنطاق الجغرافي")
p_date = c2.text_input("التاريخ", value=datetime.now().strftime('%Y-%m-%d'))
st.markdown('</div>', unsafe_allow_html=True)

# 2. خطاب الشكر والمقدمة
with st.expander("🤝 ثانياً: خطاب الإرسال والشكر والمقدمة"):
    p_thanks = st.text_area("أدخل نص الشكر أو مقدمة التقرير هنا...")

# 3. محاور التقرير (مع الأمثلة وزر التحسين)
st.markdown(f"### 🔍 ثالثاً: محاور {rtype}")
responses = {}
for i, (pillar, hint) in enumerate(STRATEGY_DATA[rtype]):
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown(f'<span class="q-label">{pillar}</span>', unsafe_allow_html=True)
    st.markdown(f'<div class="example-hint">💡 {hint}</div>', unsafe_allow_html=True)
    txt = st.text_area("أدخل البيانات أو الملاحظات هنا...", key=f"v13_{i}", height=120, label_visibility="collapsed")
    
    # زر التحسين المستعاد
    col_b, _ = st.columns([1, 3])
    with col_b:
        st.markdown('<div class="magic-btn">', unsafe_allow_html=True)
        if st.button(f"✨ تحسين الصياغة", key=f"btn13_{i}"):
            if txt:
                with st.spinner("جاري المعالجة..."):
                    res = model.generate_content(f"صغ هذا المحور بأسلوب استشاري رفيع: {txt}")
                    st.info(res.text)
            else: st.warning("أدخل نصاً أولاً")
        st.markdown('</div>', unsafe_allow_html=True)
    responses[pillar] = txt
    st.markdown('</div>', unsafe_allow_html=True)

# 4. الإضافات المخصصة (المستعادة)
st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown('<span class="q-label">➕ رابعاً: إضافة أقسام مخصصة (حسب الحاجة)</span>', unsafe_allow_html=True)
if 'v13_extra' not in st.session_state: st.session_state.v13_extra = []
new_sec = st.text_input("اسم القسم الإضافي الذي تريد إنشاؤه:")
if st.button("أنشئ القسم المخصص الآن"):
    if new_sec: st.session_state.v13_extra.append(new_sec); st.rerun()

for ex in st.session_state.v11_extra: # الربط مع الذاكرة السابقة
    st.markdown(f"**⭐ القسم المخصص: {ex}**")
    responses[ex] = st.text_area(f"أدخل بيانات {ex}...", key=f"ex13_{ex}")
st.markdown('</div>', unsafe_allow_html=True)

# 5. التوقيعات والاعتماد
st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown('<span class="q-label">🖊️ خامساً: هيكل الاعتماد والتوقيع</span>', unsafe_allow_html=True)
v1, v2, v3 = st.columns(3)
p_pre = v1.text_input("أعده:")
p_rev = v2.text_input("راجعه:")
p_app = v3.text_input("اعتمده:")
st.markdown('</div>', unsafe_allow_html=True)

# التوليد النهائي
if st.button("🚀 إصدار ومعالجة التقرير النهائي"):
    if p_name and any(responses.values()):
        with st.spinner("جاري الصهر والتحليل الاستراتيجي..."):
            all_txt = "\n".join([f"{k}: {v}" for k, v in responses.items() if v])
            prompt = f"صغ تقريراً استراتيجياً لـ {p_name}. الجهة: {p_agency}. الموجه لـ {p_donor}. المحاور: {all_txt}. التوقيعات: {p_pre}, {p_rev}, {p_app}. المعايير: فخم، رسمي، قيادي."
            res = model.generate_content(prompt)
            st.markdown(res.text)
            st.session_state['v13_out'] = res.text
    else: st.warning("يرجى تعبئة البيانات الأساسية.")

if 'v13_out' in st.session_state:
    doc = Document()
    doc.add_heading(p_name, 0)
    doc.add_paragraph(st.session_state['v13_out'])
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    st.download_button("💾 تحميل ملف Word الرسمي", bio, f"{p_name}.docx")

st.markdown('<center><a href="https://wa.me/967774575749" style="text-decoration:none; color:#25d366; font-weight:bold;">💬 الدعم الفني المباشر: 774575749</a></center>', unsafe_allow_html=True)
