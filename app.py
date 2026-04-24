import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document
from datetime import datetime

# 1. إعدادات الهوية البصرية (التركيز على النقاء البصري ومنع التداخل)
st.set_page_config(page_title="منصة المنصور AI", layout="centered")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
    
    /* توحيد الخط والاتجاه */
    * { font-family: 'Cairo', sans-serif !important; direction: rtl; text-align: right; }
    .stApp { background-color: #f4f7f9; }
    
    /* إلغاء أي تأثيرات جانبية تسبب الخطوط العمودية */
    [data-testid="stSidebar"] { display: none; }
    .main .block-container { padding-top: 2rem; max-width: 800px; }

    /* حاوية المحتوى الرئيسي */
    .report-card { 
        background: white; border-top: 8px solid #0f172a; 
        padding: 25px; border-radius: 12px; 
        box-shadow: 0 4px 20px rgba(0,0,0,0.05); margin-bottom: 20px;
        position: relative; z-index: 1;
    }
    
    .brand-header { text-align: center; margin-bottom: 30px; }
    .brand-title { color: #0f172a; font-weight: 900; font-size: 1.8rem; }
    .brand-subtitle { color: #c5a059; font-weight: 700; font-size: 0.85rem; }

    /* العناوين الرسمية */
    .section-title { 
        background: #0f172a; color: white; padding: 10px 15px; 
        border-radius: 6px; font-weight: 700; font-size: 1rem; margin: 20px 0; 
    }
    
    .q-label { color: #1e293b; font-weight: 800; border-right: 4px solid #c5a059; padding-right: 10px; margin-top: 20px; display: block; }
    .hint-box { color: #64748b; font-size: 0.8rem; background: #fffbeb; padding: 12px; border-radius: 8px; border: 1px solid #fef3c7; margin: 10px 0; line-height: 1.5; }

    /* تحسين الأزرار */
    .stButton>button { 
        background: #0f172a !important; color: white !important; 
        font-weight: 700 !important; border-radius: 8px !important; width: 100%; height: 48px; border: none !important;
    }
    .improve-btn button { 
        background: #fdfaf3 !important; color: #856404 !important; border: 1px dashed #c5a059 !important; 
        height: 35px !important; font-size: 0.8rem !important; width: auto !important;
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

# بنك الـ 12 تخصصاً كاملاً (بدون أي نقص)
REPORTS_DATABASE = {
    "📑 تقرير إنجاز دوري": [
        ("الملخص التنفيذي للأداء", "مثال: تم إنجاز 90% من الأنشطة المخططة..."),
        ("تحليل الأنشطة المنفذة", "مثال: عقد 3 ورش تدريبية وتوريد معدات..."),
        ("إدارة التحديات والحلول", "مثال: واجهنا صعوبات لوجستية وتم حلها عبر..."),
        ("الخطة القادمة", "مثال: تدشين المرحلة الثانية في الأسبوع القادم...")
    ],
    "💰 دراسة جدوى استثمارية": [
        ("تحليل الاحتياج السوقي", "مثال: عجز بنسبة 30% في قطاع التغليف..."),
        ("النمذجة المالية", "مثال: العائد المتوقع 15% خلال سنتين..."),
        ("تحليل SWOT", "مثال: القوة في التكنولوجيا والتهديد في الصرف...")
    ],
    "🎓 تقرير ختامي لتدريب": [("الأهداف والمنهجية", "إكساب المشاركين مهارات القيادة..."), ("نتائج التقييم", "تحسن الأداء من 40% إلى 90%...")],
    "🔍 متابعة وتقييم (M&E)": [("قياس KPIs", "تحقيق 95% من مؤشرات الحضور..."), ("جودة المخرجات", "تطابق المخرجات مع معايير ISO...")],
    "🏛️ حوكمة وامتثال": [("الالتزام باللوائح", "تطابق العقود مع قانون العمل..."), ("إجراءات التصحيح", "اعتماد نظام إداري جديد...")],
    "🏗️ فني وهندسي": [("المواصفات الفنية", "مطابقة المواد للكود الهندسي..."), ("السلامة المهنية", "الالتزام بأدوات السلامة...")],
    "🚑 تقييم احتياجات": [("وصف الاحتياج", "نقص حاد في مياه الشرب..."), ("خارطة التدخل", "إنشاء عيادة متنقلة...")],
    "💰 أداء مالي": [("بيان المصروفات", "تحليل بنود الصرف..."), ("انحراف الميزانية", "أسباب تجاوز الميزانية...")],
    "🌍 أثر بيئي": [("الأثر الحيوي", "تأثير المشروع على البيئة..."), ("المسؤولية المجتمعية", "مدى تقبل المجتمع...")],
    "📝 تحليل مناقصات": [("التقييم الفني", "مقارنة عروض الموردين..."), ("توصية الترسية", "اختيار المورد الأنسب...")],
    "⚠️ إدارة مخاطر": [("سجل المخاطر", "تحديد المخاطر الأمنية..."), ("خطط الاستجابة", "خطة الطوارئ المعتمدة...")],
    "🌟 استراتيجي سنوي": [("المنجز العام", "حصاد النجاحات السنوية..."), ("أهداف العام القادم", "الرؤية المستقبلية...")],
}

# --- جسم التطبيق المنظم ---
st.markdown('<div class="brand-header"><h1 class="brand-title">منصة المنصور الاستراتيجية AI</h1><p class="brand-subtitle">نظام الصياغة والتحليل المؤسسي الشامل - 2026</p></div>', unsafe_allow_html=True)

# 1. الإعداد العام (بدون سيدبار)
st.markdown('<div class="report-card">', unsafe_allow_html=True)
st.markdown('<div class="section-title">🎯 إعداد نوع التقرير والوثائق</div>', unsafe_allow_html=True)
rtype = st.selectbox("اختر تخصص التقرير:", list(REPORTS_DATABASE.keys()))
uploaded_file = st.file_uploader("ارفع مرجعاً أو مسودة (اختياري)", type=['pdf', 'docx', 'txt'])
st.markdown('</div>', unsafe_allow_html=True)

# 2. بيانات الغلاف
st.markdown('<div class="report-card">', unsafe_allow_html=True)
st.markdown('<div class="section-title">🛡️ أولاً: بيانات الغلاف الرسمي</div>', unsafe_allow_html=True)
p_name = st.text_input("عنوان التقرير (اسم المشروع) *")
p_agency = st.text_input("الجهة المُعِدّة (المؤسسة)")
p_donor = st.text_input("الجهة الموجه إليها (العميل)")
p_loc = st.text_input("المكان والنطاق الجغرافي")
p_ref = st.text_input("الرقم المرجعي (Ref No.)")
p_date = st.text_input("التاريخ", value=datetime.now().strftime('%Y-%m-%d'))
st.markdown('</div>', unsafe_allow_html=True)

# 3. المحاور (استعادة الأسئلة وزر التحسين)
st.markdown(f"### 🔍 محاور تقرير {rtype}")
user_ans = {}
for i, (pillar, hint) in enumerate(REPORTS_DATABASE[rtype]):
    st.markdown('<div class="report-card">', unsafe_allow_html=True)
    st.markdown(f'<span class="q-label">{pillar}</span>', unsafe_allow_html=True)
    st.markdown(f'<div class="hint-box">💡 {hint}</div>', unsafe_allow_html=True)
    txt = st.text_area("", key=f"v18_{i}", height=120, label_visibility="collapsed")
    
    if st.button(f"✨ تحسين صياغة {pillar}", key=f"btn_v18_{i}"):
        if txt:
            with st.spinner("جاري التحسين..."):
                res = model.generate_content(f"صغ هذا المحور بأسلوب استشاري رفيع: {txt}")
                st.info(res.text)
        else: st.warning("أدخل نصاً أولاً")
    user_ans[pillar] = txt
    st.markdown('</div>', unsafe_allow_html=True)

# 4. الإضافة المخصصة (المستقرة)
st.markdown('<div class="report-card">', unsafe_allow_html=True)
st.markdown('<div class="section-title">➕ ثالثاً: إضافة أقسام مخصصة</div>', unsafe_allow_html=True)
if 'extra_v18' not in st.session_state: st.session_state.extra_v18 = []
new_sec = st.text_input("اكتب اسم القسم الجديد:")
if st.button("أنشئ القسم المخصص الآن"):
    if new_sec and new_sec not in st.session_state.extra_v18:
        st.session_state.extra_v18.append(new_sec); st.rerun()
for ex in st.session_state.extra_v18:
    st.markdown(f"**⭐ القسم المخصص: {ex}**")
    user_ans[ex] = st.text_area(f"بيانات {ex}...", key=f"ex18_{ex}")
    if st.button(f"حذف {ex}"): st.session_state.extra_v18.remove(ex); st.rerun()
st.markdown('</div>', unsafe_allow_html=True)

# 5. التوقيعات
st.markdown('<div class="report-card">', unsafe_allow_html=True)
st.markdown('<div class="section-title">🖊️ رابعاً: هيكل الاعتماد والتوقيع</div>', unsafe_allow_html=True)
p_concl = st.text_area("الخاتمة والتوصيات الاستراتيجية:")
p_pre = st.text_input("إعداد:")
p_rev = st.text_input("مراجعة:")
p_app = st.text_input("اعتماد:")
st.markdown('</div>', unsafe_allow_html=True)

# التوليد
if st.button("🚀 إصدار ومعالجة التقرير النهائي"):
    if p_name and any(user_ans.values()):
        with st.spinner("جاري التوليد..."):
            summary = "\n".join([f"- {k}: {v}" for k, v in user_ans.items() if v])
            prompt = f"صغ تقريراً استراتيجياً لـ {p_name}. المحاور: {summary}. التوقيعات: {p_pre}, {p_rev}, {p_app}."
            res = model.generate_content(prompt)
            st.markdown(res.text)
            st.session_state['v18_out'] = res.text
    else: st.warning("يرجى ملء اسم المشروع.")

if 'v18_out' in st.session_state:
    doc = Document()
    doc.add_heading(p_name, 0)
    doc.add_paragraph(st.session_state['v18_out'])
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    st.download_button("💾 تحميل ملف Word المعتمد", bio, f"{p_name}.docx")

st.markdown('<center><a href="https://wa.me/967774575749" style="text-decoration:none; color:#25d366; font-weight:bold;">💬 الدعم: 774575749</a></center>', unsafe_allow_html=True)
