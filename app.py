import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document
from datetime import datetime

# 1. التنسيق البصري المؤسسي (نظام البطاقات المستقلة)
st.set_page_config(page_title="منصة المنصور AI - التميز المؤسسي", layout="centered")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
    * { font-family: 'Cairo', sans-serif !important; direction: rtl; text-align: right; }
    .stApp { background-color: #eff3f6; }
    #MainMenu, footer, header { visibility: hidden; }

    /* حل مشكلة تداخل الرفع نهائياً */
    [data-testid="stFileUploader"] { padding: 15px; background: #fff; border: 2px dashed #1e3a8a; border-radius: 12px; }
    [data-testid="stFileUploader"] section > button { display: none !important; }
    [data-testid="stFileUploader"] section::before {
        content: "📥 اضغط هنا لرفع الوثائق المرجعية";
        color: #1e3a8a; font-weight: 800; display: block; text-align: center; padding: 10px; cursor: pointer;
    }

    /* تصميم البطاقات المنفصلة */
    .module-card { 
        background: white; border-right: 10px solid #1e3a8a; 
        padding: 25px; border-radius: 12px; 
        box-shadow: 0 4px 15px rgba(0,0,0,0.05); margin-bottom: 25px;
    }
    
    .main-title { color: #1e3a8a; font-weight: 900; font-size: 2.2rem; text-align: center; margin-bottom: 0; }
    .main-subtitle { color: #c5a059; text-align: center; font-weight: 700; font-size: 0.9rem; margin-bottom: 35px; }

    .module-header { 
        background: #1e3a8a; color: white; padding: 8px 15px; 
        border-radius: 5px; font-weight: 700; font-size: 1rem; margin-bottom: 20px; display: inline-block;
    }
    
    .q-label { color: #1e293b; font-weight: 800; margin-top: 15px; display: block; border-bottom: 2px solid #f1f5f9; padding-bottom: 5px; }
    .hint-box { color: #64748b; font-size: 0.8rem; background: #fffbeb; padding: 10px; border-radius: 8px; border: 1px solid #fef3c7; margin: 8px 0; line-height: 1.5; }

    .stButton>button { 
        background: #1e3a8a !important; color: white !important; 
        font-weight: 800 !important; height: 55px !important; border-radius: 10px !important; width: 100%; border: none;
    }
</style>
""", unsafe_allow_html=True)

# تفعيل Gemini
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")
except:
    st.error("⚠️ يرجى ضبط GEMINI_API_KEY")

# بنك المعرفة الكامل (12 تخصصاً مع الأمثلة العميقة)
REPORTS_DB = {
    "📑 تقرير إنجاز دوري (إداري)": [
        ("الملخص التنفيذي للأداء", "مثال: تم إنجاز 90% من المهام المخططة للفترة الحالية بنجاح تام..."),
        ("الأنشطة والمنجزات المحققة", "مثال: عقد 3 ورش عمل استهدفت 120 موظفاً، وإصدار 5 أدلة إجرائية..."),
        ("إدارة الانحرافات والتحديات", "مثال: واجهنا تأخراً في التوريد وتم تجاوزه عبر تفعيل المورد البديل..."),
        ("خطة العمل القادمة", "مثال: البدء في المرحلة الميدانية الثانية وتدشين نظام الرقابة الإلكتروني...")
    ],
    "🎓 تقرير برنامج تدريبي": [
        ("بيانات المدرب والمنهجية", "مثال: اسم المدرب، خبرته، واستخدام منهجية التعلم بالتطبيق..."),
        ("إحصائيات الحضور (جندر)", "مثال: إجمالي الحضور 50 مشاركاً (30 ذكور / 20 إناث)..."),
        ("نتائج التقييم القبلي والبعدي", "مثال: تحسن مستوى الاستيعاب من 40% إلى 92% بناءً على الاختبارات..."),
        ("توصيات استدامة الأثر", "مثال: متابعة المتدربين عبر ورش عمل شهرية لضمان نقل المعرفة...")
    ],
    "💰 دراسة جدوى استثمارية": [
        ("تحليل السوق والاحتياج", "مثال: يوجد فجوة سوقية بنسبة 30% في خدمات الطاقة المتجددة..."),
        ("النمذجة المالية والربحية", "مثال: رأس المال المطلوب، العائد المتوقع، وفترة استرداد الاستثمار..."),
        ("تحليل SWOT والمنافسة", "مثال: القوة في الابتكار، والتهديد في تغير السياسات المالية...")
    ],
    "🔍 متابعة وتقييم (M&E)": [
        ("مؤشرات الأداء KPIs", "مثال: الوصول لـ 1000 مستفيد مباشر مقابل 900 مستهدف..."),
        ("رضا المستفيدين والدروس", "مثال: أظهرت الاستبيانات رضا بنسبة 95% عن سرعة الاستجابة...")
    ],
    "🏛️ حوكمة وامتثال": [
        ("الالتزام باللوائح والسياسات", "مثال: مطابقة جميع الإجراءات المالية لقانون العمل واللوائح..."),
        ("نتائج التدقيق والرقابة", "مثال: رصد ثغرة في نظام الأرشفة وتم اعتماد نظام رقمي جديد...")
    ],
    "🏗️ فني وهندسي": [
        ("مطابقة المواصفات والاختبارات", "مثال: مطابقة المواد للكود الهندسي ونتائج فحوصات المختبر..."),
        ("سير العمل وسلامة الموقع", "مثال: نسبة الإنجاز الفعلي مقابل المخطط والالتزام بأدوات السلامة...")
    ],
    "🚑 تقييم احتياجات": [("وصف الاحتياج", "نقص حاد في مياه الشرب..."), ("خارطة التدخل", "إنشاء بئر ارتوازية وتوزيع خزانات مياه...")],
    "💰 أداء مالي": [("بيان المصروفات", "تحليل شامل لبنود الصرف..."), ("انحراف الميزانية", "أسباب تجاوز الميزانية في بنود التشغيل...")],
    "🌍 أثر بيئي": [("الأثر الحيوي", "تأثير المشروع على البيئة..."), ("المسؤولية المجتمعية", "مدى تقبل المجتمع المحلي...")],
    "📝 تحليل مناقصات": [("التقييم الفني والمالي", "مقارنة العروض الفنية والمالية للموردين..."), ("توصية الترسية", "مبررات اختيار المورد الفائز بالترسية...")],
    "⚠️ إدارة مخاطر": [("سجل المخاطر", "تحديد المخاطر الأمنية والمالية..."), ("خطط الاستجابة", "خطة الطوارئ المعتمدة للتعامل مع الأزمات...")],
    "🌟 استراتيجي سنوي": [("المنجز الاستراتيجي العام", "حصاد الإنجازات السنوية مقارنة بالرؤية..."), ("أهداف العام القادم", "خارطة الطريق الاستراتيجية للسنة القادمة...")]
}

# --- بناء الواجهة ---
st.markdown('<div class="brand-header"><h1 class="main-title">منصة المنصور الاستراتيجية AI</h1><p class="main-subtitle">PREMIUM EXECUTIVE REPORTING SYSTEM V23</p></div>', unsafe_allow_html=True)

# بطاقة 1: الإعداد والرفع
st.markdown('<div class="module-card">', unsafe_allow_html=True)
st.markdown('<div class="module-header">📁 الإعداد والوثائق</div>', unsafe_allow_html=True)
rtype = st.selectbox("حدد تخصص التقرير المطلوب:", list(REPORTS_DB.keys()))
up_file = st.file_uploader("", type=['pdf', 'docx', 'txt', 'jpg', 'png'])
st.markdown('</div>', unsafe_allow_html=True)

# بطاقة 2: بيانات الغلاف
st.markdown('<div class="module-card">', unsafe_allow_html=True)
st.markdown('<div class="module-header">🛡️ أولاً: بيانات الغلاف الرسمي</div>', unsafe_allow_html=True)
p_name = st.text_input("عنوان المشروع / النشاط *")
p_agency = st.text_input("الجهة المُعِدّة (المؤسسة)")
p_donor = st.text_input("الجهة المستلمة (العميل)")
p_loc = st.text_input("مكان التنفيذ والتاريخ")
st.markdown('</div>', unsafe_allow_html=True)

# بطاقة 3: صلب التقرير (المحاور والأمثلة)
st.markdown(f"### 🔍 محاور تقرير: {rtype}")
user_ans = {}
for i, (pillar, hint) in enumerate(REPORTS_DB[rtype]):
    st.markdown('<div class="module-card">', unsafe_allow_html=True)
    st.markdown(f'<span class="q-label">{pillar}</span>', unsafe_allow_html=True)
    st.markdown(f'<div class="hint-box">💡 {hint}</div>', unsafe_allow_html=True)
    txt = st.text_area("", key=f"v23_{i}", height=120, label_visibility="collapsed")
    
    if st.button(f"✨ تحسين صياغة {pillar}", key=f"btn_v23_{i}"):
        if txt:
            with st.spinner("جاري التحسين..."):
                res = model.generate_content(f"صغ هذا المحور بأسلوب استشاري رفيع: {txt}")
                st.success(res.text)
        else: st.warning("أدخل نصاً أولاً.")
    user_ans[pillar] = txt
    st.markdown('</div>', unsafe_allow_html=True)

# بطاقة 4: الإضافات المخصصة
st.markdown('<div class="module-card">', unsafe_allow_html=True)
st.markdown('<div class="module-header">➕ إضافات مخصصة</div>', unsafe_allow_html=True)
if 'extra_v23' not in st.session_state: st.session_state.extra_v23 = []
new_sec = st.text_input("اسم القسم الجديد:")
if st.button("أنشئ القسم الآن"):
    if new_sec and new_sec not in st.session_state.extra_v23:
        st.session_state.extra_v23.append(new_sec); st.rerun()
for ex in st.session_state.extra_v23:
    st.markdown(f"**⭐ {ex}**")
    user_ans[ex] = st.text_area(f"بيانات {ex}...", key=f"ex23_{ex}")
    if st.button(f"حذف {ex}"): st.session_state.extra_v23.remove(ex); st.rerun()
st.markdown('</div>', unsafe_allow_html=True)

# بطاقة 5: التوقيعات والاعتماد
st.markdown('<div class="module-card">', unsafe_allow_html=True)
st.markdown('<div class="module-header">🖊️ الاعتماد والتوقيع</div>', unsafe_allow_html=True)
p_pre = st.text_input("إعداد:")
p_rev = st.text_input("مراجعة:")
p_app = st.text_input("اعتماد:")
st.write("---")
# زر التوليد في نهاية الأقسام مباشرة كما طلبت
if st.button("🚀 توليد التقرير النهائي الآن"):
    if p_name and any(user_ans.values()):
        with st.spinner("جاري التوليد الاستراتيجي..."):
            summary = "\n".join([f"- {k}: {v}" for k, v in user_ans.items() if v])
            prompt = f"صغ تقريراً استراتيجياً لـ {p_name}. المحاور: {summary}. التوقيعات: {p_pre}, {p_rev}, {p_app}."
            res = model.generate_content(prompt)
            st.markdown(res.text)
            st.session_state['v23_out'] = res.text
    else: st.warning("يرجى ملء البيانات الأساسية.")
st.markdown('</div>', unsafe_allow_html=True)

if 'v23_out' in st.session_state:
    doc = Document()
    doc.add_heading(p_name, 0)
    doc.add_paragraph(st.session_state['v23_out'])
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    st.download_button("💾 تحميل ملف Word المعتمد", bio, f"{p_name}.docx")

st.markdown('<center><a href="https://wa.me/967774575749" style="text-decoration:none; color:#25d366; font-weight:bold;">💬 الدعم الفني المباشر: 774575749</a></center>', unsafe_allow_html=True)
