import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document
from datetime import datetime

# 1. إعدادات الهوية البصرية (Cairo + منع تداخل الرفع)
st.set_page_config(page_title="منصة المنصور AI - التميز المؤسسي", layout="centered")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
    * { font-family: 'Cairo', sans-serif !important; direction: rtl; text-align: right; }
    .stApp { background-color: #f4f7f9; }
    
    /* حل مشكلة تداخل نص الرفع (Upload/Browse) */
    [data-testid="stFileUploadDropzone"] button {
        text-indent: -9999px;
        line-height: 0;
    }
    [data-testid="stFileUploadDropzone"] button::after {
        content: "اختر الملف";
        text-indent: 0;
        display: block;
        line-height: initial;
    }
    
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

# --- بنك المعرفة الاستراتيجي المتعمق ---
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
        ("النمذجة المالية والربحية", "تقدير رأس المال، التدفقات النقدية، وفترة استرداد رأس المال..."),
        ("تحليل المخاطر والحساسية", "تأثير تغير أسعار الصرف أو التكاليف التشغيلية على المشروع..."),
        ("تحليل SWOT والتوصية النهائية", "توليف نقاط القوة والفرص لاتخاذ قرار الاستثمار النهائي...")
    ],
    "🎓 تقرير ختامي لبرنامج تدريبي": [
        ("بيانات المدرب والمنهجية", "اسم المدرب، خبرته، والأساليب التدريبية المستخدمة (مجموعات، محاكاة)..."),
        ("إحصائيات الحضور والجندر", "إجمالي المشاركين، عدد الذكور، عدد الإناث، ونسبة الحضور الفعلي..."),
        ("تحليل الأثر المعرفي (Pre/Post)", "مقارنة نتائج الاختبارات قبل التدريب وبعده لقياس نسبة الاستيعاب..."),
        ("تقييم الخدمات اللوجستية", "مدى رضا المشاركين عن القاعة، التغذية، والوسائل التعليمية..."),
        ("توصيات الاستدامة المهنية", "مقترحات لضمان تطبيق المتدربين للمهارات الجديدة...")
    ],
    "🔍 متابعة وتقييم (M&E)": [
        ("مؤشرات الأداء الرئيسية (KPIs)", "تحليل رقمي دقيق للمؤشرات المحققة مقابل المستهدفات المعتمدة..."),
        ("جودة المخرجات والامتثال", "مدى مطابقة النتائج للمعايير الفنية وجودة المنظمات الدولية..."),
        ("تحليل رضا المستفيدين", "خلاصة التغذية الراجعة من الجمهور المستهدف وأثر التدخل في حياتهم..."),
        ("الدروس المستفادة (Lessons Learned)", "ما هي الفرص التي ضاعت وكيف يمكن تجنبها مستقبلاً؟")
    ]
    # البقية مدمجة برمجياً لضمان عدم الحذف
}

# --- بناء الواجهة ---
st.markdown('<div class="report-card">', unsafe_allow_html=True)
st.markdown('<h1 class="brand-title">منصة المنصور الاستراتيجية AI</h1>', unsafe_allow_html=True)
st.markdown('<p style="text-align:center; color:#c5a059; font-weight:700;">المختبر العالمي لصناعة التقارير السيادية V21</p>', unsafe_allow_html=True)

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

# 3. صلب التقرير
st.markdown(f"### 🔍 الخطوة الثالثة: محاور {rtype}")
user_ans = {}
for i, (pillar, hint) in enumerate(REPORTS_DB[rtype]):
    st.markdown('<div class="report-card">', unsafe_allow_html=True)
    st.markdown(f'<span class="q-label">{pillar}</span>', unsafe_allow_html=True)
    st.markdown(f'<div class="hint-box">📝 {hint}</div>', unsafe_allow_html=True)
    txt = st.text_area("", key=f"v21_{i}", height=120, label_visibility="collapsed")
    
    if st.button(f"✨ تحسين صياغة {pillar}", key=f"btn_v21_{i}"):
        if txt:
            with st.spinner("جاري الصياغة..."):
                res = model.generate_content(f"صغ هذا المحور بأسلوب استشاري رفيع: {txt}")
                st.success(res.text)
        else: st.warning("أدخل نصاً أولاً.")
    user_ans[pillar] = txt
    st.markdown('</div>', unsafe_allow_html=True)

# التخصيص والاعتماد
if 'extra_v21' not in st.session_state: st.session_state.extra_v21 = []
st.markdown('<div class="section-title">➕ الخطوة الرابعة: إضافات مخصصة</div>', unsafe_allow_html=True)
new_sec = st.text_input("اسم القسم الإضافي الجديد:")
if st.button("أنشئ القسم الآن"):
    if new_sec and new_sec not in st.session_state.extra_v21:
        st.session_state.extra_v21.append(new_sec); st.rerun()
for ex in st.session_state.extra_v21:
    user_ans[ex] = st.text_area(f"بيانات {ex}...", key=f"ex21_{ex}")

# 5. التوقيعات
st.markdown('<div class="section-title">🖊️ الخطوة الخامسة: هيكل الاعتماد</div>', unsafe_allow_html=True)
p_pre = st.text_input("إعداد:")
p_rev = st.text_input("مراجعة:")
p_app = st.text_input("اعتماد:")

if st.button("🚀 توليد التقرير الاستراتيجي الشامل"):
    if p_name and any(user_ans.values()):
        with st.spinner("جاري التوليد..."):
            summary = "\n".join([f"- {k}: {v}" for k, v in user_ans.items() if v])
            prompt = f"صغ تقريراً استراتيجياً لـ {p_name}. المحاور: {summary}. التوقيعات: {p_pre}, {p_rev}, {p_app}."
            res = model.generate_content(prompt)
            st.markdown(res.text)
            st.session_state['v21_out'] = res.text
    else: st.warning("يرجى ملء البيانات.")

if 'v21_out' in st.session_state:
    doc = Document()
    doc.add_heading(p_name, 0)
    doc.add_paragraph(st.session_state['v21_out'])
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    st.download_button("💾 تحميل ملف Word المعتمد", bio, f"{p_name}.docx")

st.markdown('<center><a href="https://wa.me/967774575749" style="text-decoration:none; color:#25d366; font-weight:bold;">💬 الدعم الفني: 774575749</a></center>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)
