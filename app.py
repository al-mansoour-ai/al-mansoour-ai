import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document
from datetime import datetime

# 1. إعدادات الهوية البصرية الرسمية (Premium Corporate Style)
st.set_page_config(page_title="منصة المنصور AI - التميز المؤسسي", layout="centered")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
    * { font-family: 'Cairo', sans-serif !important; direction: rtl; text-align: right; }
    .stApp { background-color: #f4f7f9; }
    
    .report-card { 
        background: white; border-top: 10px solid #1e3a8a; 
        padding: 30px; border-radius: 15px; 
        box-shadow: 0 10px 30px rgba(0,0,0,0.05); margin-bottom: 35px;
    }
    .brand-title { color: #1e3a8a; font-weight: 900; font-size: 2.3rem; text-align: center; margin: 0; }
    .section-title { background: #1e3a8a; color: white; padding: 12px 20px; border-radius: 8px; font-weight: 700; margin: 25px 0; }
    .q-label { color: #1e293b; font-weight: 800; border-right: 5px solid #d4af37; padding-right: 12px; margin-top: 25px; display: block; }
    .hint-box { color: #64748b; font-size: 0.85rem; background: #fffbeb; padding: 12px; border-radius: 8px; border: 1px solid #fef3c7; margin: 10px 0; line-height: 1.6; }
    .stButton>button { background: #1e3a8a !important; color: white !important; font-weight: 700 !important; border-radius: 10px !important; width: 100%; height: 55px; border: none !important; }
</style>
""", unsafe_allow_html=True)

# تفعيل Gemini
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")
except:
    st.error("⚠️ يرجى ضبط GEMINI_API_KEY")

# --- بنك المعرفة الاستراتيجي المتعمق (V20) ---
REPORTS_DB = {
    "📑 تقرير إنجاز دوري (إداري)": [
        ("الملخص التنفيذي للأداء", "وصف شامل لمستوى التقدم مقابل الأهداف الاستراتيجية المخططة..."),
        ("الأنشطة الميدانية والعملية", "قائمة مفصلة بكل نشاط تم تنفيذه، مكانه، والفئة المستهدفة..."),
        ("تحليل الموارد والاستهلاك المالي", "مستوى استنزاف الميزانية مقابل المخرجات المحققة (كفاءة الإنفاق)..."),
        ("إدارة الانحرافات والتحديات", "وصف دقيق لأي تأخير أو عائق ميداني والإجراء التصحيحي المتخذ..."),
        ("خطة العمل للفترة القادمة", "تحديد الأولويات القصوى والمهام العاجلة للأسبوع/الشهر القادم...")
    ],
    "💰 دراسة جدوى استثمارية": [
        ("تحليل السوق والمنافسة", "حجم الطلب، الفجوة السوقية، ونقاط قوة وضعف المنافسين المباشرين..."),
        ("الدراسة الفنية واللوجستية", "الموقع الجغرافي، التكنولوجيا المستخدمة، وسلاسل الإمداد المطلوبة..."),
        ("النمذجة المالية والربحية", "تقدير رأس المال، التدفقات النقدية، وفترة استرداد رأس المال (Payback Period)..."),
        ("تحليل المخاطر والحساسية", "تأثير تغير أسعار الصرف أو التكاليف التشغيلية على استمرارية المشروع..."),
        ("تحليل SWOT والتوصية النهائية", "توليف نقاط القوة والفرص لاتخاذ قرار الاستثمار النهائي...")
    ],
    "🎓 تقرير ختامي لبرنامج تدريبي": [
        ("بيانات المدرب والمنهجية", "اسم المدرب، خبرته، والأساليب التدريبية المستخدمة (مجموعات، محاكاة)..."),
        ("إحصائيات الحضور والجندر", "إجمالي المشاركين، عدد الذكور، عدد الإناث، ونسبة الحضور الفعلي..."),
        ("تحليل الأثر المعرفي (Pre/Post)", "مقارنة نتائج الاختبارات قبل التدريب وبعده لقياس نسبة الاستيعاب..."),
        ("تقييم الخدمات اللوجستية", "مدى رضا المشاركين عن القاعة، التغذية، والوسائل التعليمية..."),
        ("توصيات الاستدامة المهنية", "مقترحات لضمان تطبيق المتدربين للمهارات الجديدة في مؤسساتهم...")
    ],
    "🔍 متابعة وتقييم (M&E)": [
        ("مؤشرات الأداء الرئيسية (KPIs)", "تحليل رقمي دقيق للمؤشرات المحققة مقابل المستهدفات المعتمدة..."),
        ("جودة المخرجات والامتثال", "مدى مطابقة النتائج للمعايير الفنية وجودة المنظمات الدولية..."),
        ("تحليل رضا المستفيدين", "خلاصة التغذية الراجعة من الجمهور المستهدف وأثر التدخل في حياتهم..."),
        ("الدروس المستفادة (Lessons Learned)", "ما هي الفرص التي ضاعت وكيف يمكن تجنبها في المشاريع القادمة؟")
    ],
    "🏗️ تقرير فني وهندسي": [
        ("مطابقة المواصفات والاختبارات", "نتائج فحوصات المختبر للمواد المستخدمة ومدى مطابقتها للكود..."),
        ("نسبة الإنجاز المادي والزمني", "مقارنة المنجز في الموقع مع الجدول الزمني المخطط له هندسياً..."),
        ("السلامة والصحة المهنية", "تقرير حول الالتزام بأدوات السلامة والحوادث المرصودة (إن وجدت)..."),
        ("التعديلات والأوامر التغييرية", "وصف لأي تغيير في المخططات الأصلية ومبرراته الفنية والمالية...")
    ],
    "🏛️ حوكمة وامتثال": [
        ("مراجعة الالتزام التنظيمي", "مدى توافق الممارسات الإدارية مع اللوائح الداخلية والقوانين المحلية..."),
        ("نتائج التدقيق والرقابة", "رصد أي ثغرات مالية أو إدارية في نظام الصلاحيات والرقابة..."),
        ("خطة تصحيح المسار", "الإجراءات القانونية والإدارية المتخذة لسد الثغرات المكتشفة...")
    ],
    "🚑 تقييم احتياجات إنسانية": [
        ("وصف الأزمة والاحتياج الراهن", "تحليل دقيق للوضع الميداني والفجوة في الخدمات الأساسية..."),
        ("تحليل الفئات الأكثر تضرراً", "تحديد المستفيدين حسب العمر والجنس ونوع الإعاقة والاحتياج..."),
        ("أولويات التدخل العاجل", "قائمة بالاحتياجات التي تتطلب استجابة فورية (ماء، غذاء، دواء)..."),
        ("خارطة الطريق للاستجابة", "المقترح العملي للجهات المانحة حول كيفية وكلفة التدخل...")
    ],
    "💰 تقرير أداء مالي": [
        ("تحليل التدفقات النقدية", "بيان شامل للإيرادات والمصروفات خلال الفترة المحددة..."),
        ("انحرافات الميزانية المعتمدة", "تحليل الأسباب وراء تجاوز أو انخفاض الصرف في البنود الرئيسية..."),
        ("المخاطر المالية والسيولة", "تقييم قدرة المؤسسة على الوفاء بالتزاماتها المالية قصيرة الأجل...")
    ],
    "🌍 أثر بيئي واجتماعي": [
        ("تحليل البصمة البيئية", "تأثير أنشطة المشروع على التربة، الهواء، والموارد المائية..."),
        ("المسؤولية والقبول الاجتماعي", "مدى قبول المجتمع المحلي للمشروع وتأثيره على العادات المحلية..."),
        ("خطة التخفيف والاستدامة", "الإجراءات المتبعة لتقليل الأضرار وحماية التوازن البيئي...")
    ],
    "📝 تحليل مناقصات وترسية": [
        ("التقييم الفني للمتقدمين", "مقارنة كفاءة الموردين، الخبرة السابقة، والمواصفات المعروضة..."),
        ("التقييم المالي والمفاضلة", "تحليل الأسعار مقابل الجودة وتقديم أفضل قيمة مقابل المال..."),
        ("توصية لجنة البت", "المبررات القانونية والفنية لاختيار المورد الأنسب للترسية...")
    ],
    "⚠️ إدارة مخاطر وطوارئ": [
        ("سجل المخاطر المحدث", "تحديد المخاطر الجديدة (أمنية، سياسية، مالية) التي ظهرت..."),
        ("تحليل الاحتمالية والأثر", "تقييم خطورة كل خطر ومدى تأثيره على استمرارية العمل..."),
        ("فعالية خطط الاستجابة", "مدى نجاح إجراءات الطوارئ التي تم تفعيلها خلال الفترة...")
    ],
    "🌟 تقرير استراتيجي سنوي": [
        ("حصاد الرؤية والأهداف", "ما تم تحقيقه من الرؤية الكبرى للمؤسسة خلال عام كامل..."),
        ("المركز المالي والمؤسسي", "قوة المؤسسة في السوق وعلاقاتها مع الشركاء والمانحين..."),
        ("التوجهات الاستراتيجية القادمة", "خارطة الطريق للسنة الجديدة والأهداف الطموحة المراد تحقيقها...")
    ]
}

# --- بناء الواجهة التنفيذية ---
st.markdown('<div class="report-card">', unsafe_allow_html=True)
st.markdown('<h1 class="brand-title">منصة المنصور الاستراتيجية AI</h1>', unsafe_allow_html=True)
st.markdown('<p style="text-align:center; color:#c5a059; font-weight:700;">المختبر العالمي لصناعة التقارير السيادية V20</p>', unsafe_allow_html=True)

# 1. الإعداد والرفع
st.markdown('<div class="section-title">🎯 الخطوة الأولى: نوع التقرير والمراجع</div>', unsafe_allow_html=True)
rtype = st.selectbox("اختر تخصص التقرير المطلوب لتفعيل المنهجية:", list(REPORTS_DB.keys()))
up_file = st.file_uploader("ارفع وثيقة مرجعية أو صوراً للتحليل (اختياري)", type=['pdf', 'docx', 'txt', 'jpg', 'png'])

# 2. بيانات الغلاف الرسمي
st.markdown('<div class="section-title">🛡️ الخطوة الثانية: بيانات الغلاف الرسمي</div>', unsafe_allow_html=True)
p_name = st.text_input("اسم المشروع / البرنامج التدريبي / الفعالية *")
p_agency = st.text_input("الجهة المُعِدّة (المؤسسة أو المستشار)")
p_donor = st.text_input("الجهة الموجه إليها (العميل أو المانح)")
p_loc = st.text_input("مكان التنفيذ (المدينة/المنطقة)")
p_ref = st.text_input("الرقم المرجعي للمستند (Reference No.)")
p_date = st.text_input("تاريخ الإصدار", value=datetime.now().strftime('%Y-%m-%d'))

# 3. صلب التقرير (الأسئلة العميقة + زر التحسين)
st.markdown(f"### 🔍 الخطوة الثالثة: محاور {rtype}")
user_ans = {}
for i, (pillar, hint) in enumerate(REPORTS_DB[rtype]):
    st.markdown('<div class="report-card">', unsafe_allow_html=True)
    st.markdown(f'<span class="q-label">{pillar}</span>', unsafe_allow_html=True)
    st.markdown(f'<div class="hint-box">📝 {hint}</div>', unsafe_allow_html=True)
    txt = st.text_area("", key=f"v20_{i}", height=120, label_visibility="collapsed")
    
    if st.button(f"✨ تحسين صياغة {pillar}", key=f"btn_v20_{i}"):
        if txt:
            with st.spinner("جاري الصياغة الاستشارية..."):
                res = model.generate_content(f"صغ هذا المحور بأسلوب استشاري رفيع، صوت نشط، وإيجاز: {txt}")
                st.success(res.text)
        else: st.warning("أدخل نصاً أولاً.")
    user_ans[pillar] = txt
    st.markdown('</div>', unsafe_allow_html=True)

# 4. الإضافة المخصصة
st.markdown('<div class="report-card">', unsafe_allow_html=True)
st.markdown('<div class="section-title">➕ الخطوة الرابعة: إضافات مخصصة</div>', unsafe_allow_html=True)
if 'extra_v20' not in st.session_state: st.session_state.extra_v20 = []
new_sec = st.text_input("اسم القسم الإضافي الجديد:")
if st.button("أنشئ القسم الآن"):
    if new_sec and new_sec not in st.session_state.extra_v20:
        st.session_state.extra_v20.append(new_sec); st.rerun()
for ex in st.session_state.extra_v20:
    st.markdown(f"**⭐ القسم المخصص: {ex}**")
    user_ans[ex] = st.text_area(f"بيانات {ex}...", key=f"ex20_{ex}")
    if st.button(f"حذف {ex}"): st.session_state.extra_v20.remove(ex); st.rerun()
st.markdown('</div>', unsafe_allow_html=True)

# 5. التوقيعات والاعتماد
st.markdown('<div class="report-card">', unsafe_allow_html=True)
st.markdown('<div class="section-title">🖊️ الخطوة الخامسة: هيكل الاعتماد</div>', unsafe_allow_html=True)
p_concl = st.text_area("الخاتمة والتوصيات الختامية:")
p_pre = st.text_input("إعداد (الاسم والصفة):")
p_rev = st.text_input("مراجعة:")
p_app = st.text_input("اعتماد:")
st.markdown('</div>', unsafe_allow_html=True)

# التوليد النهائي
if st.button("🚀 توليد ومعالجة التقرير الاستراتيجي الشامل"):
    if p_name and any(user_ans.values()):
        with st.spinner("جاري صهر البيانات وفق معايير ISO 2145..."):
            summary = "\n".join([f"- {k}: {v}" for k, v in user_ans.items() if v])
            prompt = f"صغ تقريراً استراتيجياً لـ {p_name}. النوع: {rtype}. الجهة: {p_agency}. المحاور: {summary}. الخاتمة: {p_concl}. التوقيعات: {p_pre}, {p_rev}, {p_app}. التزم بترقيم ISO 2145، وصوت نشط، لغة فصحى رفيعة."
            res = model.generate_content(prompt)
            st.markdown(res.text)
            st.session_state['v20_out'] = res.text
    else: st.warning("يرجى ملء اسم المشروع والبيانات الأساسية.")

if 'v20_out' in st.session_state:
    doc = Document()
    doc.add_heading(p_name, 0)
    doc.add_paragraph(st.session_state['v20_out'])
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    st.download_button("💾 تحميل ملف Word المعتمد", bio, f"{p_name}.docx")

st.markdown('<center><a href="https://wa.me/967774575749" style="text-decoration:none; color:#25d366; font-weight:bold;">💬 الدعم الفني: 774575749</a></center>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)
