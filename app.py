import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document
from datetime import datetime

# 1. التنسيق البصري الفخم (Branding & UI)
st.set_page_config(page_title="منصة المنصور AI - الإصدار المستقر والنهائي", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
    * { font-family: 'Cairo', sans-serif !important; direction: rtl; text-align: right; }
    .stApp { background-color: #f1f5f9; }
    #MainMenu, footer, header { visibility: hidden; }

    .main-card { 
        background: white; border-top: 12px solid #0f172a; 
        padding: 45px; border-radius: 15px; 
        box-shadow: 0 15px 35px rgba(0,0,0,0.1); margin-top: -60px; 
    }
    .brand-title { color: #0f172a; font-weight: 900; font-size: 2.5rem; text-align: center; margin: 0; }
    .brand-subtitle { color: #c5a059; text-align: center; font-weight: 700; font-size: 1rem; margin-bottom: 40px; }

    .section-header { 
        background: #0f172a; color: white; padding: 12px 25px; 
        border-radius: 8px; font-weight: 700; margin: 20px 0;
    }
    .q-label { color: #1e293b; font-weight: 800; border-right: 5px solid #c5a059; padding-right: 12px; margin-top: 15px; display: block; }
    .hint-box { color: #64748b; font-size: 0.82rem; background: #f8fafc; padding: 10px; border-radius: 6px; border: 1px solid #e2e8f0; margin-bottom: 10px; }

    .stButton>button { 
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%) !important; color: white !important; 
        font-weight: 800 !important; height: 60px !important; font-size: 1.2rem !important;
        border-radius: 12px !important; border: none !important; width: 100%; transition: 0.3s;
    }
    .magic-btn button { 
        background: #fdfaf3 !important; color: #856404 !important; border: 1px dashed #c5a059 !important; 
        height: 35px !important; font-size: 0.8rem !important; margin-top: -10px; width: auto !important;
    }
    .whatsapp-btn { background: #25d366; color: white !important; padding: 12px 30px; border-radius: 50px; text-decoration: none; font-weight: 700; display: inline-flex; align-items: center; gap: 10px; margin-top: 20px; }
</style>
""", unsafe_allow_html=True)

# 2. بنك المعلومات الاستراتيجي (الـ 12 تخصصاً كاملاً)
STRATEGY_BANK = {
    "📑 تقرير إنجاز دوري": ["الملخص التنفيذي للأداء", "تحليل الأنشطة والمنجزات", "إدارة الانحرافات والجدول الزمني", "التحديات والحلول التصحيحية", "الأولويات الاستراتيجية القادمة"],
    "💰 دراسة جدوى استثمارية": ["تحليل الاحتياج السوقي", "المواصفات الفنية والمتطلبات", "النمذجة المالية والربحية", "تحليل الحساسية والمخاطر", "تحليل المنافسة (SWOT)", "توصية الاستثمار النهائية"],
    "🎓 تقرير ختامي لتدريب": ["منهجية وأهداف التدريب", "تحليل نتائج القبلي والبعدي", "تقييم المدرب والمادة العلمية", "تفاعل المشاركين واللوجستيات", "توصيات استدامة الأثر"],
    "🔍 تقرير متابعة وتقييم (M&E)": ["قياس مؤشرات الأداء (KPIs)", "جودة المخرجات والامتثال", "تحليل رضا المستفيدين", "الدروس المستفادة والنمو"],
    "🚑 تقرير تقييم احتياجات": ["وصف الأزمة والاحتياج الراهن", "ديموغرافيا الفئات المتضررة", "الأولويات العاجلة للاستجابة", "خارطة التدخل المقترحة"],
    "🏛️ تقرير حوكمة وامتثال": ["الالتزام باللوائح والسياسات", "نتائج التدقيق والرقابة", "الثغرات المرصودة في النظام", "خطة تحسين مستوى الشفافية"],
    "💰 تقرير أداء مالي": ["بيان الدخل والمصروفات", "تحليل انحرافات الميزانية", "التدفق النقدي والسيولة", "توصيات رفع كفاءة الإنفاق"],
    "🏗️ تقرير فني وهندسي": ["مطابقة المواصفات والمواد", "نتائج اختبارات الجودة", "المعوقات والحلول الهندسية", "تقرير السلامة المهنية"],
    "🌍 تقرير أثر بيئي واجتماعي": ["تحليل الأثر البيئي والحيوي", "المسؤولية المجتمعية والرضا", "إجراءات التخفيف", "خطة الاستدامة البيئية"],
    "📝 تقرير تحليل مناقصات": ["التقييم الفني للمتقدمين", "التقييم المالي والمقارنة", "تحليل مخاطر الموردين", "توصية الترسية النهائية"],
    "⚠️ تقرير إدارة مخاطر": ["سجل المخاطر المرصودة", "تحليل الاحتمالية والأثر", "خطط الاستجابة والطوارئ", "مسؤوليات المتابعة"],
    "🌟 تقرير استراتيجي سنوي": ["الرؤية والمنجز العام", "تحليل الأداء السنوي الشامل", "الوضع المالي الموحد", "الأهداف الاستراتيجية القادمة"]
}

# 3. الهيكل التشغيلي (بدون هدم)
st.markdown('<div class="main-card">', unsafe_allow_html=True)
st.markdown('<h1 class="brand-title">منصة المنصور AI للتقارير</h1>', unsafe_allow_html=True)
st.markdown('<p class="brand-subtitle">نظام الصياغة والتحليل المؤسسي المتكامل - الإصدار السيادي 2026</p>', unsafe_allow_html=True)

try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")
except:
    st.error("⚠️ يرجى ضبط المفتاح في الإعدادات")

rtype = st.selectbox("🎯 حدد تخصص التقرير لضبط المنهجية تلقائياً:", list(STRATEGY_BANK.keys()))

# أ. الغلاف والبيانات التعريفية
with st.expander("🛡️ أولاً: صفحة الغلاف والبيانات الرسمية", expanded=True):
    c1, c2 = st.columns(2)
    p_name = c1.text_input("عنوان التقرير (اسم المشروع) *")
    p_ref = c2.text_input("الرقم المرجعي (Ref No.)")
    p_agency = c1.text_input("الجهة المُعِدّة (المؤسسة)")
    p_donor = c2.text_input("الجهة الموجه إليها (المانح/العميل)")
    p_loc = c1.text_input("المكان والنطاق الجغرافي")
    p_date = c2.text_input("التاريخ", value=datetime.now().strftime('%Y-%m-%d'))

# ب. الشكر والمقدمة
with st.expander("🤝 ثانياً: خطاب الإرسال والشكر والمقدمة"):
    p_thanks = st.text_area("كلمة شكر وتقديم للتقرير:", placeholder="مثال: نتقدم بخالص الشكر للجهات الداعمة...")

# ج. صلب التقرير (المحاور الاستراتيجية + زر التحسين)
st.markdown(f'<div class="section-header">🔍 ثالثاً: المحاور الاستراتيجية لـ {rtype}</div>', unsafe_allow_html=True)
responses = {}
for i, pillar in enumerate(STRATEGY_BANK[rtype]):
    st.markdown(f'<span class="q-label">{pillar}</span>', unsafe_allow_html=True)
    st.markdown(f'<div class="hint-box">💡 مثال احترافي: ابدأ بذكر أهم المخرجات في {pillar} ثم التفاصيل...</div>', unsafe_allow_html=True)
    txt = st.text_area("", key=f"v8_{i}_{rtype}", height=120, label_visibility="collapsed")
    
    # زر التحسين تحت كل سؤال مباشرة
    c_btn, c_sp = st.columns([1, 4])
    with c_btn:
        st.markdown('<div class="magic-btn">', unsafe_allow_html=True)
        if st.button(f"✨ تحسين الصياغة", key=f"btn_v8_{i}"):
            if txt:
                with st.spinner("جاري المعالجة..."):
                    res = model.generate_content(f"صغ هذا المحور بأسلوب استشاري فخم وصوت نشط: {txt}")
                    st.code(res.text)
            else: st.warning("أدخل نصاً")
        st.markdown('</div>', unsafe_allow_html=True)
    responses[pillar] = txt

# د. الخاتمة والملاحق والاعتماد
with st.expander("📌 رابعاً: الخاتمة والتوصيات والملاحق"):
    p_concl = st.text_area("الخاتمة والتوصيات الاستراتيجية:", placeholder="خلاصة التقرير والتوصيات القابلة للتنفيذ...")
    p_appendix = st.text_area("الملاحق (صور، روابط، جداول):", placeholder="أدخل روابط الصور أو كشف الملاحق...")

# هـ. التخصيص الكامل
with st.expander("➕ خامساً: إضافة بيانات وأقسام مخصصة"):
    st.info("إذا كان لديك بيانات أخرى وتريد إضافتها، اضغط على الزر وخصص قسماً جديداً:")
    if 'v8_extra' not in st.session_state: st.session_state.v8_extra = []
    new_sec = st.text_input("اسم القسم الإضافي:")
    if st.button("خصص القسم الآن"):
        if new_sec: st.session_state.v8_extra.append(new_sec); st.rerun()
    for ex in st.session_state.v8_extra:
        st.markdown(f"**⭐ قسم مخصص: {ex}**")
        responses[ex] = st.text_area(f"بيانات {ex}...", key=f"ex8_{ex}")
        if st.button(f"حذف {ex}"): st.session_state.v8_extra.remove(ex); st.rerun()

# و. هيكل التوقيعات
with st.expander("🖊️ سادساً: هيكل الاعتماد والتوقيع"):
    v1, v2, v3 = st.columns(3)
    p_pre = v1.text_input("أعده:")
    p_rev = v2.text_input("راجعه:")
    p_app = v3.text_input("اعتمده:")

st.write("---")
if st.button("🚀 توليد ومعالجة التقرير الاستراتيجي الشامل"):
    if p_name and any(responses.values()):
        with st.spinner("جاري صهر البيانات وفق معايير الجودة العالمية..."):
            all_data = "\n".join([f"- {k}: {v}" for k, v in responses.items() if v])
            full_prompt = f"""
            بصفتك مستشاراً دولياً، صغ تقريراً استراتيجياً متكاملاً.
            الغلاف: {p_name}، المرجع {p_ref}، الجهة {p_agency}، الموجه لـ {p_donor}.
            خطاب التقديم: {p_thanks}
            المحاور الفنية: {all_data}
            الخاتمة والتوصيات: {p_concl}
            الملاحق: {p_appendix}
            التوقيعات: المعد {p_pre}، المراجع {p_rev}، المعتمد {p_app}.
            المعايير: ISO 2145، صوت نشط، نبرة قيادية.
            """
            res = model.generate_content(full_prompt)
            st.markdown(res.text)
            st.session_state['v8_out'] = res.text
    else: st.warning("يرجى ملء البيانات.")

if 'v8_out' in st.session_state:
    doc = Document()
    doc.add_heading(f"Report: {p_name}", 0)
    doc.add_paragraph(st.session_state['v8_out'])
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    st.download_button("💾 تحميل المستند الرسمي (Word)", bio, f"{p_name}.docx")

st.markdown('<center><a href="https://wa.me/967774575749" class="whatsapp-btn">💬 تواصل معنا للدعم الاستشاري: 774575749</a></center>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)
st.markdown("<center style='color:#64748b; font-size:0.75rem; margin-top:20px;'>🛡️ حقوق المنصة محفوظة لشبكة المنصور الدولية للاستشارات | 2026</center>", unsafe_allow_html=True)
