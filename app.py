import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document
from datetime import datetime

# 1. الهندسة البصرية المتقدمة (حل مشكلة التداخل + نظام البطاقات)
st.set_page_config(page_title="منصة المنصور AI - الإصدار المستقر", layout="centered")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
    * { font-family: 'Cairo', sans-serif !important; direction: rtl; text-align: right; }
    .stApp { background-color: #f1f5f9; }
    #MainMenu, footer, header { visibility: hidden; }

    /* الحل النهائي والجذري لمشكلة تداخل نصوص الرفع (Upload/Browse) */
    [data-testid="stFileUploadDropzone"] button {
        opacity: 0 !important; /* إخفاء النص الأصلي تماماً */
        position: relative;
        z-index: 2;
    }
    [data-testid="stFileUploadDropzone"] section::after {
        content: "📁 اضغط هنا لاختيار الملفات";
        position: absolute;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        color: #1e3a8a;
        font-weight: 800;
        z-index: 1;
        pointer-events: none; /* جعل النص غير قابل للنقر ليمرر النقرة للزر المخفي */
        width: 100%;
        text-align: center;
    }

    /* تصميم الوحدات المنفصلة (نظام البطاقات) */
    .module-card { 
        background: white; border-right: 12px solid #1e3a8a; 
        padding: 30px; border-radius: 15px; 
        box-shadow: 0 8px 25px rgba(0,0,0,0.06); margin-bottom: 30px;
    }
    
    .brand-header { text-align: center; padding: 20px 0; }
    .brand-title { color: #1e3a8a; font-weight: 900; font-size: 2.3rem; margin: 0; }
    .brand-subtitle { color: #c5a059; font-weight: 700; font-size: 1rem; }

    .module-label { 
        background: #1e3a8a; color: white; padding: 10px 20px; 
        border-radius: 8px; font-weight: 700; font-size: 1.1rem; margin-bottom: 25px; display: inline-block;
    }
    
    .q-text { color: #1e293b; font-weight: 900; font-size: 1.1rem; margin-bottom: 10px; display: block; border-right: 5px solid #c5a059; padding-right: 12px; }
    .hint-style { color: #64748b; font-size: 0.85rem; background: #fffbeb; padding: 12px; border-radius: 10px; border: 1px solid #fef3c7; margin-bottom: 15px; line-height: 1.6; }

    /* الأزرار الملكية */
    .stButton>button { 
        background: linear-gradient(135deg, #1e3a8a 0%, #152e6d 100%) !important; color: white !important; 
        font-weight: 900 !important; height: 60px !important; border-radius: 12px !important; width: 100%; border: none;
    }
</style>
""", unsafe_allow_html=True)

# تفعيل Gemini
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")
except:
    st.error("⚠️ يرجى ضبط مفتاح GEMINI_API_KEY")

# بنك المعرفة الشامل (الـ 12 تخصصاً مع التعمق الاستراتيجي)
REPORTS_BANK = {
    "📑 تقرير إنجاز دوري (إداري)": [
        ("الملخص التنفيذي للأداء العام", "مثال: تم إنجاز 90% من المهام المخططة للفترة الحالية بنجاح..."),
        ("تحليل الأنشطة والمنجزات المحققة", "مثال: عقد 3 ورش عمل، وتوريد 50 وحدة تقنية، وإتمام 4 زيارات ميدانية..."),
        ("إدارة الانحرافات والتحديات", "مثال: واجهنا تأخراً في التوريد وتم تجاوزه عبر تفعيل المورد البديل..."),
        ("خطة العمل للفترة القادمة", "مثال: البدء في المرحلة الميدانية الثانية وتدشين نظام الرقابة الإلكتروني...")
    ],
    "🎓 تقرير برنامج تدريبي ختامي": [
        ("بيانات المدرب ونبذة عنه", "مثال: اسم المدرب، خبرته الاستشارية، والمنهجية المتبعة..."),
        ("إحصائيات الحضور (ذكور/إناث)", "مثال: إجمالي الحضور 50 مشاركاً (30 ذكور / 20 إناث)..."),
        ("نتائج التقييم القبلي والبعدي", "مثال: تحسن مستوى الاستيعاب من 40% إلى 95% بناءً على الاختبارات..."),
        ("توصيات استدامة الأثر المهني", "مثال: عقد جلسات تنشيطية كل 3 أشهر لمتابعة تطبيق المهارات المكتسبة...")
    ],
    "💰 دراسة جدوى استثمارية": [
        ("تحليل الاحتياج السوقي والمنافسة", "مثال: يوجد فجوة سوقية بنسبة 30% في قطاع الخدمات المستهدف..."),
        ("النمذجة المالية وتوقعات الربحية", "مثال: رأس المال المطلوب، العائد المتوقع، وفترة استرداد الاستثمار..."),
        ("تحليل SWOT وقرار الاستثمار", "مثال: القوة في الابتكار، والتهديد في تغير السياسات المالية...")
    ],
    "🔍 متابعة وتقييم (M&E)": [
        ("مؤشرات الأداء الرئيسية KPIs", "مثال: الوصول لـ 1000 مستفيد مباشر مقابل 900 مستهدف..."),
        ("جودة المخرجات والدروس المستفادة", "مثال: أظهرت الاستبيانات رضا بنسبة 95% عن جودة الخدمة المقدمة...")
    ],
    "🏗️ تقرير هندسي وفني": [
        ("مطابقة المواصفات والاختبارات", "مثال: مطابقة المواد للكود الهندسي ونتائج فحوصات المختبر..."),
        ("سير العمل وتقرير السلامة", "مثال: نسبة الإنجاز الفعلي مقابل المخطط والالتزام بأدوات السلامة...")
    ],
    "🏛️ حوكمة وامتثال": [
        ("الالتزام باللوائح والسياسات", "مثال: مطابقة جميع الإجراءات المالية لقانون العمل واللوائح..."),
        ("نتائج التدقيق وإجراءات التصحيح", "مثال: رصد ثغرة في نظام الأرشفة وتم اعتماد نظام رقمي جديد...")
    ],
    "🚑 تقييم احتياجات إنسانية": [("وصف الاحتياج الراهن", "تحليل الوضع الميداني وفجوة الخدمات..."), ("خارطة التدخل المقترحة", "الأولويات العاجلة للاستجابة وكيفية التدخل...")],
    "💰 تقرير أداء مالي": [("بيان المصروفات والسيولة", "تحليل شامل لبنود الصرف والتدفقات النقدية..."), ("انحراف الميزانية المعتمدة", "تحليل الفروقات بين المخطط والفعلي...")],
    "🌍 أثر بيئي واجتماعي": [("تحليل الأثر البيئي والاجتماعي", "تأثير المشروع على المجتمع المحلي والبيئة المحيطة..."), ("خطة التخفيف والاستدامة", "الإجراءات المتبعة لتقليل الأضرار الجانبية...")],
    "📝 تحليل مناقصات وترسية": [("التقييم الفني والمالي", "مقارنة العروض الفنية والمالية واختيار المورد الأنسب..."), ("توصية لجنة الترسية", "مبررات اختيار العرض الفائز بناءً على الجودة والسعر...")],
    "⚠️ إدارة مخاطر وطوارئ": [("سجل المخاطر المحدث", "تحديد المخاطر الجديدة وفعالية خطط الطوارئ المفعّلة..."), ("خطط الاستجابة للأزمات", "كيفية التعامل مع التهديدات الأمنية أو المالية المرصودة...")],
    "🌟 تقرير استراتيجي سنوي": [("حصاد المنجز الاستراتيجي العام", "ما تم تحقيقه من الرؤية الكبرى للمؤسسة خلال عام..."), ("أهداف العام القادم", "خارطة الطريق الاستراتيجية والأهداف الطموحة للسنة الجديدة...")]
}

# --- بناء الواجهة التنفيذية ---
st.markdown('<div class="brand-header"><h1 class="brand-title">منصة المنصور الاستراتيجية AI</h1><p class="brand-subtitle">نظام الصياغة والتحليل المؤسسي الشامل - V24</p></div>', unsafe_allow_html=True)

# بطاقة 1: الإعداد والرفع
st.markdown('<div class="module-card">', unsafe_allow_html=True)
st.markdown('<div class="module-label">📁 الخطوة 1: الوثائق والمراجع</div>', unsafe_allow_html=True)
rtype = st.selectbox("حدد نوع التقرير المطلوب لتفعيل المنهجية:", list(REPORTS_BANK.keys()))
up_file = st.file_uploader("", type=['pdf', 'docx', 'txt', 'jpg', 'png'])
st.markdown('</div>', unsafe_allow_html=True)

# بطاقة 2: الغلاف الرسمي
st.markdown('<div class="module-card">', unsafe_allow_html=True)
st.markdown('<div class="module-label">🛡️ الخطوة 2: بيانات الغلاف الرسمي</div>', unsafe_allow_html=True)
p_name = st.text_input("اسم المشروع / البرنامج التدريبي / الفعالية *")
p_agency = st.text_input("الجهة المُعِدّة (المؤسسة أو المستشار)")
p_donor = st.text_input("الجهة الموجه إليها (العميل أو المانح)")
p_loc = st.text_input("مكان التنفيذ (المدينة/المنطقة)")
p_date = st.text_input("التاريخ", value=datetime.now().strftime('%Y-%m-%d'))
st.markdown('</div>', unsafe_allow_html=True)

# بطاقة 3: محاور التقرير (نظام البطاقات المستقلة)
st.markdown(f"### 🔍 الخطوة 3: محاور تقرير {rtype}")
user_ans = {}
for i, (pillar, hint) in enumerate(REPORTS_BANK[rtype]):
    st.markdown('<div class="module-card">', unsafe_allow_html=True)
    st.markdown(f'<span class="q-text">{pillar}</span>', unsafe_allow_html=True)
    st.markdown(f'<div class="hint-style">💡 {hint}</div>', unsafe_allow_html=True)
    txt = st.text_area("", key=f"v24_{i}", height=120, label_visibility="collapsed")
    
    if st.button(f"✨ تحسين صياغة {pillar}", key=f"btn_v24_{i}"):
        if txt:
            with st.spinner("جاري الصياغة..."):
                res = model.generate_content(f"صغ هذا المحور بأسلوب استشاري رفيع: {txt}")
                st.success(res.text)
        else: st.warning("أدخل نصاً أولاً.")
    user_ans[pillar] = txt
    st.markdown('</div>', unsafe_allow_html=True)

# بطاقة 4: التخصيص والاعتماد
st.markdown('<div class="module-card">', unsafe_allow_html=True)
st.markdown('<div class="module-label">🖊️ الخطوة 4: الاعتماد والتوقيع</div>', unsafe_allow_html=True)
p_pre = st.text_input("إعداد:")
p_rev = st.text_input("مراجعة:")
p_app = st.text_input("اعتماد:")
st.write("---")
if st.button("🚀 توليد ومعالجة التقرير الاستراتيجي الشامل"):
    if p_name and any(user_ans.values()):
        with st.spinner("جاري التوليد..."):
            summary = "\n".join([f"- {k}: {v}" for k, v in user_ans.items() if v])
            prompt = f"صغ تقريراً استراتيجياً لـ {p_name}. المحاور: {summary}. التوقيعات: {p_pre}, {p_rev}, {p_app}."
            res = model.generate_content(prompt)
            st.markdown(res.text)
            st.session_state['v24_out'] = res.text
    else: st.warning("يرجى ملء البيانات الأساسية.")
st.markdown('</div>', unsafe_allow_html=True)

if 'v24_out' in st.session_state:
    doc = Document()
    doc.add_heading(p_name, 0)
    doc.add_paragraph(st.session_state['v24_out'])
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    st.download_button("💾 تحميل ملف Word المعتمد", bio, f"{p_name}.docx")

st.markdown('<center><a href="https://wa.me/967774575749" style="text-decoration:none; color:#25d366; font-weight:bold;">💬 الدعم الفني المباشر: 774575749</a></center>', unsafe_allow_html=True)
