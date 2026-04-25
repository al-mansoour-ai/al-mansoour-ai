import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document
from datetime import datetime

# --- 1. الهندسة البصرية الملكية (Premium Gold UI) ---
st.set_page_config(page_title="منصة المنصور AI - الإصدار الشامل", layout="centered")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
    
    html, body, [data-testid="stAppViewContainer"] {
        direction: rtl !important;
        text-align: right !important;
        background-color: #050a14 !important;
    }

    * { font-family: 'Cairo', sans-serif !important; }

    /* حل مشكلة تداخل نصوص الرفع نهائياً */
    [data-testid="stFileUploadDropzone"] button { opacity: 0 !important; position: relative; z-index: 2; }
    [data-testid="stFileUploadDropzone"] section::after {
        content: "📁 اضغط هنا لإرفاق الوثائق المرجعية";
        position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%);
        color: #d4af37; font-weight: 800; z-index: 1; width: 100%; text-align: center;
    }

    /* تصميم البطاقات الذهبية المستقلة */
    .gold-card { 
        background: #0f172a; border: 1px solid rgba(212, 175, 55, 0.2);
        border-right: 10px solid #d4af37; padding: 25px; border-radius: 15px; 
        margin-bottom: 30px; box-shadow: 0 10px 30px rgba(0,0,0,0.5);
    }
    
    .gold-label { color: #d4af37; font-weight: 900; font-size: 1.2rem; margin-bottom: 15px; display: block; }
    .hint-style { 
        color: #cbd5e1; font-size: 0.85rem; background: rgba(212, 175, 55, 0.05); 
        padding: 12px; border-radius: 8px; border: 1px solid rgba(212, 175, 55, 0.15); 
        margin-bottom: 15px; line-height: 1.6;
    }

    .stButton>button { 
        background: linear-gradient(135deg, #d4af37 0%, #a68a2d 100%) !important; 
        color: #050a14 !important; font-weight: 900 !important; height: 55px !important; 
        border-radius: 12px !important; border: none !important;
    }

    /* تنسيق الحقول */
    input, textarea { background-color: #1e293b !important; color: #ffffff !important; border: 1px solid rgba(212, 175, 55, 0.2) !important; }
    #MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# --- 2. محرك الذكاء الاصطناعي ---
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")
except:
    st.error("⚠️ يرجى ضبط مفتاح API في الإعدادات")

# البنك الكامل (12 تخصصاً - محمي من الحذف)
REPORTS_DB = {
    "📑 تقرير إنجاز دوري": [
        ("الملخص التنفيذي للأداء", "مثال: تم إنجاز 90% من المهام بنجاح تام خلال هذه الفترة..."),
        ("تحليل الأنشطة والمنجزات", "مثال: عقد 3 ورش عمل وتوريد 50 وحدة تقنية..."),
        ("إدارة التحديات والحلول", "مثال: واجهنا تأخراً في الإمداد وتم حله عبر المورد البديل..."),
        ("خطة العمل القادمة", "مثال: البدء في المرحلة الميدانية الثانية وتدشين نظام الرقابة...")
    ],
    "🎓 تقرير تدريبي ختامي": [
        ("بيانات المدرب والمنهجية", "مثال: المنهجية التشاركية، خبرة المدرب الاستشارية..."),
        ("إحصائيات الحضور (جندر)", "مثال: إجمالي الحضور 50 (30 ذكور / 20 إناث)..."),
        ("نتائج التقييم القبلي والبعدي", "مثال: تحسن مستوى الاستيعاب من 40% إلى 95%..."),
        ("توصيات استدامة الأثر", "مثال: عقد جلسات تنشيطية كل 3 أشهر...")
    ],
    "💰 دراسة جدوى استثمارية": [("تحليل السوق", "حجم الفجوة السوقية..."), ("النموذجة المالية", "العائد المتوقع وفترة الاسترداد..."), ("تحليل SWOT", "نقاط القوة والتهديدات...")],
    "🔍 متابعة وتقييم (M&E)": [("مؤشرات الأداء KPIs", "تحليل رقمي للمستهدفات..."), ("جودة المخرجات والرضا", "أظهرت الاستبيانات رضا بنسبة 95%...")],
    "🏗️ تقرير هندسي وفني": [("مطابقة المواصفات", "نتائج فحص المواد..."), ("سير العمل والسلامة", "نسبة الإنجاز والالتزام بالأمن...")],
    "🏛️ حوكمة وامتثال": [("الالتزام باللوائح", "تطابق الإجراءات مع القانون..."), ("نتائج التدقيق", "رصد الثغرات وإجراءات التصحيح...")],
    "🚑 تقييم احتياجات": [("وصف الاحتياج", "تحليل الوضع الميداني..."), ("خارطة التدخل", "الأولويات العاجلة للاستجابة...")],
    "💰 أداء مالي": [("بيان المصروفات", "تحليل بنود الصرف..."), ("انحراف الميزانية", "أسباب الفروقات المالية...")],
    "🌍 أثر بيئي واجتماعي": [("تحليل الأثر", "تأثير المشروع على المجتمع والبيئة..."), ("خطة التخفيف", "تقليل الأضرار الجانبية...")],
    "📝 تحليل مناقصات": [("التقييم الفني والمالي", "مقارنة عروض الموردين..."), ("توصية الترسية", "مبررات اختيار العرض الفائز...")],
    "⚠️ إدارة مخاطر": [("سجل المخاطر المحدث", "المخاطر الجديدة والتعامل معها..."), ("خطط الاستجابة", "فعالية خطط الطوارئ...")],
    "🌟 استراتيجي سنوي": [("المنجز الاستراتيجي", "حصاد الرؤية السنوية..."), ("أهداف العام القادم", "خارطة الطريق للسنة الجديدة...")]
}

# --- 3. الواجهة التنفيذية ---
st.markdown('<h1 style="text-align:center; color:#d4af37; font-weight:900;">منصة المنصور الاستراتيجية AI</h1>', unsafe_allow_html=True)
st.markdown('<p style="text-align:center; color:#8a99af; font-weight:bold;">نظام الصياغة الشامل - الإصدار الذهبي V27</p>', unsafe_allow_html=True)

# 1. الإعداد والرفع
st.markdown('<div class="gold-card">', unsafe_allow_html=True)
st.markdown('<span class="gold-label">📁 الخطوة 1: الوثائق والنوع</span>', unsafe_allow_html=True)
rtype = st.selectbox("حدد التخصص المطلوب:", list(REPORTS_DB.keys()))
up_file = st.file_uploader("", type=['pdf', 'docx', 'txt'])
st.markdown('</div>', unsafe_allow_html=True)

# 2. بيانات الغلاف
st.markdown('<div class="gold-card">', unsafe_allow_html=True)
st.markdown('<span class="gold-label">🛡️ الخطوة 2: بيانات الغلاف الرسمي</span>', unsafe_allow_html=True)
p_name = st.text_input("عنوان المشروع أو الفعالية *")
p_agency = st.text_input("الجهة المُعِدّة")
p_donor = st.text_input("الجهة الموجه إليها")
p_loc = st.text_input("مكان التنفيذ والتاريخ")
st.markdown('</div>', unsafe_allow_html=True)

# 3. المحاور الديناميكية (مع زر التحسين)
st.markdown(f"### 🔍 محاور تقرير: {rtype}")
user_ans = {}
for i, (pillar, hint) in enumerate(REPORTS_DB[rtype]):
    st.markdown('<div class="gold-card">', unsafe_allow_html=True)
    st.markdown(f'<span class="gold-label">{pillar}</span>', unsafe_allow_html=True)
    st.markdown(f'<div class="hint-style">💡 {hint}</div>', unsafe_allow_html=True)
    txt = st.text_area("أدخل البيانات هنا:", key=f"v27_txt_{i}", height=120)
    
    if st.button(f"✨ تحسين صياغة {pillar}", key=f"btn_v27_{i}"):
        if txt:
            with st.spinner("جاري التنسيق الاستشاري..."):
                res = model.generate_content(f"صغ هذا المحور بأسلوب استشاري فخم وصوت نشط: {txt}")
                st.success(res.text)
        else: st.warning("أدخل نصاً أولاً")
    user_ans[pillar] = txt
    st.markdown('</div>', unsafe_allow_html=True)

# 4. إضافة قسم مخصص (المحرك المستعاد)
st.markdown('<div class="gold-card">', unsafe_allow_html=True)
st.markdown('<span class="gold-label">➕ إضافة أقسام مخصصة</span>', unsafe_allow_html=True)
if 'custom_v27' not in st.session_state: st.session_state.custom_v27 = []
new_sec = st.text_input("اسم القسم الجديد الذي تود إضافته:")
if st.button("أنشئ القسم المخصص الآن"):
    if new_sec and new_sec not in st.session_state.custom_v27:
        st.session_state.custom_v27.append(new_sec); st.rerun()

for ex in st.session_state.custom_v27:
    st.markdown(f"**⭐ {ex}**")
    user_ans[ex] = st.text_area(f"أدخل بيانات {ex}...", key=f"ex27_{ex}")
    if st.button(f"حذف {ex}"): st.session_state.custom_v27.remove(ex); st.rerun()
st.markdown('</div>', unsafe_allow_html=True)

# 5. التوقيعات والتوليد
st.markdown('<div class="gold-card">', unsafe_allow_html=True)
st.markdown('<span class="gold-label">🖊️ التوقيعات والاعتماد</span>', unsafe_allow_html=True)
p_pre = st.text_input("إعداد:")
p_rev = st.text_input("مراجعة:")
p_app = st.text_input("اعتماد:")
st.write("---")
if st.button("🚀 توليد ومعالجة التقرير الاستراتيجي النهائي"):
    if p_name and any(user_ans.values()):
        with st.spinner("جاري صياغة الوثيقة السيادية..."):
            summary = "\n".join([f"- {k}: {v}" for k, v in user_ans.items() if v])
            full_prompt = f"صغ تقريراً استراتيجياً فخماً لـ ({p_name}). النوع: {rtype}. المحاور: {summary}. التواقيع: {p_pre}, {p_rev}, {p_app}. التزم بترقيم ISO 2145."
            res = model.generate_content(full_prompt)
            st.markdown(f'<div style="background: white; color: #050a14; padding: 30px; border-radius: 10px; border-right: 15px solid #d4af37; line-height: 2;">{res.text}</div>', unsafe_allow_html=True)
            st.session_state['v27_out'] = res.text
    else: st.warning("يرجى ملء البيانات الأساسية.")
st.markdown('</div>', unsafe_allow_html=True)

if 'v27_out' in st.session_state:
    doc = Document()
    doc.add_heading(p_name, 0)
    doc.add_paragraph(st.session_state['v27_out'])
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    st.download_button("💾 تحميل ملف Word المعتمد", bio, f"{p_name}.docx")

st.markdown('<center><a href="https://wa.me/967774575749" style="text-decoration:none; color:#d4af37; font-weight:900;">💬 الدعم الفني المباشر (المنصور): 774575749</a></center>', unsafe_allow_html=True)
