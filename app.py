import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document
from datetime import datetime

# 1. التنسيق البصري السيادي (Cairo + Premium Blue/Gold)
st.set_page_config(page_title="منصة المنصور AI - النظام المتكامل", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
    html, body, [class*="st-"], * { font-family: 'Cairo', sans-serif !important; direction: rtl; text-align: right; }
    .stApp { background-color: #f0f4f8; }
    #MainMenu, footer, header { visibility: hidden; }
    .main-box { 
        background: white; border-top: 12px solid #1e3a8a; padding: 50px; 
        border-radius: 20px; box-shadow: 0 20px 50px rgba(0,0,0,0.1); margin-top: -60px; 
    }
    .brand-title { color: #1e3a8a; font-weight: 900; font-size: 3rem; text-align: center; }
    .tagline { color: #d4af37; text-align: center; font-weight: 700; margin-bottom: 40px; font-size: 1.2rem; }
    .section-header { background: #1e3a8a; color: white; padding: 15px 25px; border-radius: 12px; margin: 30px 0; font-weight: 900; font-size: 1.4rem; }
    .q-label { color: #0f172a; font-weight: 800; border-right: 6px solid #d4af37; padding-right: 15px; display: block; margin-top: 30px; font-size: 1.1rem; }
    .hint { color: #64748b; font-size: 0.85rem; margin-bottom: 10px; background: #f8fafc; padding: 10px; border-radius: 8px; }
    .stButton>button { background: linear-gradient(135deg, #1e3a8a 0%, #152e6d 100%) !important; color: white !important; font-weight: 900 !important; height: 65px !important; border-radius: 15px !important; font-size: 1.3rem !important; }
    .magic-btn button { background: #f0f9ff !important; color: #1e3a8a !important; border: 1px dashed #1e3a8a !important; height: 35px !important; font-size: 0.8rem !important; margin-top: -10px; }
    .whatsapp-btn { background: #25d366; color: white !important; padding: 15px 30px; border-radius: 50px; text-decoration: none; font-weight: 900; display: inline-flex; align-items: center; justify-content: center; gap: 10px; }
</style>
""", unsafe_allow_html=True)

# 2. بنك المعرفة الاستراتيجي (12 تخصصاً)
REPORT_STRATEGY = {
    "📑 تقرير إنجاز دوري": ["الملخص التنفيذي للأداء العام", "تحليل الأنشطة والمهام المنفذة", "إدارة الانحرافات عن الخطة الزمنية", "استهلاك الموارد والميزانية", "التحديات والحلول التصحيحية", "خطة العمل للفترة القادمة"],
    "💰 دراسة جدوى استثمارية": ["تحليل الفجوة والاحتياج السوقي", "المواصفات الفنية والمتطلبات", "النمذجة المالية وتوقعات الدخل", "تحليل الحساسية ونقطة التعادل", "تحليل SWOT (المنافسة والفرص)", "توصية الاستثمار النهائية"],
    "🎓 تقرير ختامي لتدريب": ["الأهداف والمنهجية التدريبية", "تحليل نتائج التقييم القبلي والبعدي", "كفاءة المدرب والمادة العلمية", "تفاعل المشاركين واللوجستيات", "توصيات استدامة الأثر المهني"],
    "🔍 تقرير متابعة وتقييم (M&E)": ["قياس مؤشرات الأداء (KPIs)", "جودة المخرجات والامتثال المعياري", "تحليل رضا المستفيدين", "الدروس المستفادة والنمو المؤسسي", "توصيات التطوير الاستراتيجي"],
    "🚑 تقرير تقييم احتياجات": ["وصف الأزمة والاحتياج الراهن", "تحديد الفئات الأكثر تضرراً", "أولويات الاستجابة العاجلة", "تحليل الموارد المتاحة والفجوة", "خارطة طريق التدخل المقترحة"],
    "🏛️ تقرير حوكمة وامتثال": ["الالتزام باللوائح والسياسات", "نتائج التدقيق والرقابة الدورية", "الثغرات المرصودة في النظام", "إجراءات التصحيح والتطوير"],
    "💰 تقرير أداء مالي": ["بيان الدخل والمصروفات الفعلية", "تحليل الانحرافات عن الميزانية", "حالة التدفق النقدي والسيولة", "التوصيات لرفع كفاءة الإنفاق"],
    "🏗️ تقرير فني وهندسي": ["مطابقة المواصفات الفنية والمواد", "نتائج اختبارات الجودة الميدانية", "المعوقات والحلول الهندسية", "تقرير السلامة المهنية والموقع"],
    "🌍 تقرير أثر بيئي واجتماعي": ["تحليل الأثر البيئي والحيوي", "المسؤولية المجتمعية والرضا المحلي", "إجراءات التخفيف من الأضرار", "خطة الاستدامة البيئية"],
    "📝 تقرير تحليل مناقصات": ["التقييم الفني للمتقدمين", "التقييم المالي والمقارنة", "تحليل مخاطر الموردين", "توصية الترسية النهائية"],
    "⚠️ تقرير إدارة مخاطر": ["سجل المخاطر المرصودة", "تحليل الاحتمالية والأثر", "خطط الاستجابة والطوارئ", "مسؤوليات المتابعة والتحكم"],
    "🌟 تقرير استراتيجي سنوي": ["الرؤية والرسالة والمنجز العام", "تحليل الأداء الاستراتيجي السنوي", "الوضع المالي الموحد", "الأهداف الاستراتيجية للعام القادم"]
}

# 3. الهيكل التشغيلي للمنصة
st.markdown('<div class="main-box">', unsafe_allow_html=True)
st.markdown('<h1 class="brand-title">منصة المنصور الاستراتيجية AI</h1>', unsafe_allow_html=True)
st.markdown('<p class="tagline">نظام توليد الوثائق الرسمية المتكامل - GLOBAL CONSULTING STANDARD</p>', unsafe_allow_html=True)

try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")
except:
    st.error("⚠️ يرجى ضبط المفتاح في الإعدادات")

# اختيار النوع
rtype = st.selectbox("🎯 حدد نوع التقرير الاستراتيجي المراد توليده:", list(REPORT_STRATEGY.keys()))

# أ. صفحة الغلاف والبيانات الرسمية
st.markdown('<div class="section-header">🛡️ أولاً: صفحة الغلاف والبيانات التعريفية</div>', unsafe_allow_html=True)
c1, c2 = st.columns(2)
p_name = c1.text_input("عنوان التقرير (اسم المشروع) *")
p_ref = c2.text_input("الرقم المرجعي للمستند (Reference No.)")
p_agency = c1.text_input("الجهة المُعِدّة (اسم مؤسستك)")
p_donor = c2.text_input("الجهة الموجه إليها (العميل/المانح)")
p_loc = c1.text_input("المكان والنطاق الجغرافي")
p_date = c2.text_input("تاريخ الإصدار", value=datetime.now().strftime('%Y-%m-%d'))

# ب. الشكر والتقديم
st.markdown('<div class="section-header">🤝 ثانياً: خطاب الإرسال والشكر</div>', unsafe_allow_html=True)
p_thanks = st.text_area("كلمة شكر وتقدير للشركاء والداعمين (اختياري):", placeholder="مثال: نتقدم بخالص الشكر لـ...")

# ج. المحاور العميقة
st.markdown(f'<div class="section-header">🔍 ثالثاً: المحاور الاستراتيجية لـ {rtype}</div>', unsafe_allow_html=True)
user_responses = {}
for i, pillar in enumerate(REPORT_STRATEGY[rtype]):
    st.markdown(f'<span class="q-label">{pillar}</span>', unsafe_allow_html=True)
    st.markdown(f'<p class="hint">💡 مثال: أدخل البيانات الخاصة بـ {pillar} لتمكين الذكاء الاصطناعي من تحليلها...</p>', unsafe_allow_html=True)
    txt = st.text_area("", key=f"v5_{i}_{rtype}", height=120, label_visibility="collapsed")
    
    # زر التحسين لكل سؤال
    cb, cs = st.columns([1, 4])
    with cb:
        if st.button(f"✨ تحسين {i+1}", key=f"btn_v5_{i}"):
            if txt:
                with st.spinner("جاري الصياغة..."):
                    res = model.generate_content(f"صغ هذه الفقرة بأسلوب استشاري رفيع، صوت نشط، وإيجاز: {txt}")
                    st.code(res.text)
            else: st.warning("أدخل نصاً")
    user_responses[pillar] = txt

# د. التخصيص الكامل
st.markdown('<div class="section-header">➕ رابعاً: إضافات وتخصيص</div>', unsafe_allow_html=True)
st.info("إذا كان لديك بيانات أخرى وتريد إضافتها، اضغط على الزر وخصص قسماً جديداً:")
if 'extra_sec' not in st.session_state: st.session_state.extra_sec = []
new_sec = st.text_input("اسم القسم الإضافي:")
if st.button("إضافة القسم المخصص"):
    if new_sec: st.session_state.extra_sec.append(new_sec); st.rerun()

for es in st.session_state.extra_sec:
    st.markdown(f"**⭐ {es}**")
    user_responses[es] = st.text_area(f"بيانات {es}...", key=f"es_{es}")
    if st.button(f"حذف {es}"): st.session_state.extra_sec.remove(es); st.rerun()

# هـ. الاعتماد والتوقيع
st.markdown('<div class="section-header">🖊️ خامساً: هيكل الاعتماد والتوقيعات</div>', unsafe_allow_html=True)
col_v1, col_v2, col_v3 = st.columns(3)
p_pre = col_v1.text_input("أعده (الاسم والصفة)")
p_rev = col_v2.text_input("راجعه (الاسم والصفة)")
p_app = col_v3.text_input("اعتمده (الاسم والصفة)")

# و. التوليد النهائي
st.write("---")
if st.button("🚀 معالجة وتوليد التقرير السيادي الكامل"):
    if p_name and any(user_responses.values()):
        with st.spinner("جاري صياغة الوثيقة وفق معايير الجودة العالمية..."):
            summary_data = "\n".join([f"- {k}: {v}" for k, v in user_responses.items() if v])
            full_prompt = f"""
            بصفتك مستشاراً دولياً رفيع المستوى، صغ تقريراً سيادياً من نوع {rtype}.
            الغلاف: {p_name}، المرجع {p_ref}، الجهة {p_agency}، الموجه لـ {p_donor}، المكان {p_loc}، التاريخ {p_date}.
            خطاب الشكر: {p_thanks}
            المعطيات الفنية: {summary_data}
            هيكل التوقيع: المعد {p_pre}، المراجع {p_rev}، المعتمد {p_app}.
            
            الشروط:
            1. ابدأ بملخص تنفيذي مبهر.
            2. استخدم الترقيم ISO 2145 (1. ثم 1.1).
            3. استخدم نبرة قيادية (Executive Tone) وصوت نشط.
            4. أضف توصيات استراتيجية قابلة للتنفيذ في النهاية.
            5. صمم صفحة التوقيعات في آخر المستند بشكل رسمي.
            """
            response = model.generate_content(full_prompt)
            st.markdown("### 📄 المعاينة النهائية للمستند:")
            st.markdown(response.text)
            st.session_state['final_result'] = response.text
    else: st.warning("يرجى تعبئة البيانات الأساسية.")

# ز. التصدير
if 'final_result' in st.session_state:
    doc = Document()
    doc.add_heading(f"Report: {p_name}", 0)
    doc.add_paragraph(st.session_state['final_result'])
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    st.download_button("💾 تحميل الوثيقة المعتمدة (Word)", bio, f"{p_name}.docx")

st.markdown('<center><a href="https://wa.me/967774575749" class="whatsapp-btn">🟢 للدعم الفني والتطوير: 774575749</a></center>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)
st.markdown("<center style='color:#64748b; font-size:0.75rem; margin-top:20px;'>🛡️ شبكة المنصور الدولية للاستشارات والذكاء الاصطناعي | 2026</center>", unsafe_allow_html=True)
