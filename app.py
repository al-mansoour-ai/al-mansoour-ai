import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document

# 1. التنسيق البصري المؤسسي (فخامة رسمية)
st.set_page_config(page_title="منصة المنصور AI - الإصدار الاحترافي V40", layout="centered")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700;900&display=swap');
    html, body, [class*="st-"] { font-family: 'Cairo', sans-serif; direction: rtl; text-align: right; }
    .stApp { background-color: #f8fafc; }
    #MainMenu, footer, header { visibility: hidden; }
    .main-box {
        background: white; border-top: 10px solid #1e3a8a; 
        padding: 30px; border-radius: 15px; 
        box-shadow: 0 10px 25px rgba(0,0,0,0.05); margin-top: -50px;
    }
    .brand-title { color: #1e3a8a; font-weight: 900; font-size: 2.2rem; text-align: center; margin-bottom: 5px; }
    .tagline { background: #1e3a8a; color: #fbbf24; padding: 5px 15px; border-radius: 20px; font-size: 0.8rem; display: table; margin: 0 auto 25px auto; font-weight: bold; }
    .section-title { color: #1e3a8a; font-weight: 700; border-right: 4px solid #fbbf24; padding-right: 10px; margin-top: 20px; background: #f0f4f8; padding: 5px 10px; }
    .hint { color: #64748b; font-size: 0.8rem; margin-bottom: 5px; font-style: italic; }
    .stButton>button {
        background: linear-gradient(90deg, #1e3a8a, #d4af37) !important;
        color: white !important; font-weight: bold !important; height: 45px !important; border-radius: 10px !important; border: none !important; width: 100%;
    }
    .magic-btn button { background: #f0fdf4 !important; color: #166534 !important; border: 1px dashed #166534 !important; height: 30px !important; font-size: 0.75rem !important; margin-top: -10px; }
</style>
""", unsafe_allow_html=True)

# 2. مصفوفة التقارير الشاملة (8 تخصصات)
GLOBAL_REPORTS = {
    "📑 تقرير الإنجاز الدوري": ["الملخص التنفيذي ومستوى الإنجاز", "تحليل الانحرافات عن الخطة", "إدارة التحديات والمخاطر", "الخطوات التصحيحية القادمة"],
    "🎓 تقرير ختامي لتدريب": ["نتائج التقييم القبلي والبعدي", "كفاءة المادة العلمية", "تفاعل المشاركين واللوجستيات", "توصيات استدامة الأثر"],
    "💰 تقرير الأداء المالي": ["تحليل المصروفات مقابل الميزانية", "تحليل انحرافات التكلفة", "المخاطر المالية والامتثال", "التوصيات المالية للفترة القادمة"],
    "📊 تقرير المتابعة والتقييم": ["قياس مؤشرات الأداء (KPIs)", "جودة المخرجات ورضا المستفيدين", "الدروس المستفادة", "التوصيات الاستراتيجية"],
    "🚑 تقرير تقييم الاحتياجات": ["تحليل الوضع الراهن والفجوة", "تحديد الفئات الأكثر احتياجاً", "الأولويات العاجلة", "خارطة طريق التدخل"],
    "🏛️ تقرير الحوكمة والامتثال": ["الالتزام باللوائح والسياسات", "نتائج الرقابة والتدقيق", "الثغرات المرصودة", "إجراءات تطوير الأداء"],
    "🌍 تقرير الأثر البيئي": ["تحليل الأثر البيئي والحيوي", "المسؤولية المجتمعية", "إجراءات التخفيف من الآثار", "استدامة الموارد"],
    "🏗️ تقرير فني وهندسي": ["المواصفات الفنية ومطابقة المواد", "نتائج اختبارات الجودة", "المعوقات الإنشائية", "الحلول الهندسية المنفذة"]
}

# 3. الهيكل الرئيسي
st.markdown('<div class="main-box">', unsafe_allow_html=True)
st.markdown('<h1 class="brand-title">المنصور AI للتقارير الاستراتيجية</h1>', unsafe_allow_html=True)
st.markdown('<div class="tagline">إصدار 2026 | معايير IBCS & ISO 2145</div>', unsafe_allow_html=True)

# تفعيل الذكاء الاصطناعي
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")
except:
    st.error("⚠️ يرجى ضبط المفتاح في الإعدادات")

# واجهة المدخلات
rtype = st.selectbox("🎯 اختر تخصص التقرير لضبط المنهجية:", list(GLOBAL_REPORTS.keys()))
p_name = st.text_input("اسم المشروع / البرنامج *")

st.markdown('<p class="section-title">المحاور الاستراتيجية للتقرير</p>', unsafe_allow_html=True)

responses = {}
for i, pillar in enumerate(GLOBAL_REPORTS[rtype]):
    st.write(f"**{pillar}**")
    st.markdown(f'<p class="hint">💡 مثال: أدخل نقاطاً مختصرة حول {pillar} وسيقوم النظام بصياغتها.</p>', unsafe_allow_html=True)
    
    txt = st.text_area("", key=f"area_{i}", height=100, label_visibility="collapsed")
    
    # زر التحسين تحت كل سؤال
    col_empty, col_btn = st.columns([3, 1])
    with col_btn:
        st.markdown('<div class="magic-btn">', unsafe_allow_html=True)
        if st.button(f"✨ تحسين المحور", key=f"btn_{i}"):
            if txt:
                with st.spinner("جاري التحسين..."):
                    res = model.generate_content(f"صغ هذه النقاط بأسلوب استشاري قوي (صوت نشط، جمل قصيرة): {txt}")
                    st.success("تم التحسين! انسخ النص أدناه إذا أعجبك:")
                    st.code(res.text)
            else: st.warning("أدخل نصاً أولاً")
        st.markdown('</div>', unsafe_allow_html=True)
    
    responses[pillar] = txt

st.write("---")
if st.button("🚀 توليد ومعالجة التقرير النهائي بالكامل"):
    if p_name and any(responses.values()):
        with st.spinner("جاري المعالجة النهائية وفق معايير ISO 2145..."):
            full_data = str(responses)
            prompt = f"أنت خبير تطوير مؤسسي. صغ تقريراً من نوع {rtype} للمشروع {p_name}. التزم بالترقيم الدولي، الصوت النشط، الإيجاز، ونبرة قيادية رفيعة. البيانات: {full_data}"
            final_res = model.generate_content(prompt)
            st.markdown("### النتيجة النهائية:")
            st.info(final_res.text)
            st.session_state['final_doc'] = final_res.text
    else: st.warning("يرجى ملء البيانات واسم المشروع")

# التصدير
if 'final_doc' in st.session_state:
    doc = Document()
    doc.add_heading(f"تقرير: {p_name}", 0)
    doc.add_paragraph(st.session_state['final_doc'])
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    st.download_button("💾 تحميل التقرير (Word)", bio, f"{p_name}.docx")

st.markdown('</div>', unsafe_allow_html=True)
st.markdown("<center style='color:#94a3b8; font-size:0.7rem; margin-top:20px;'>🛡️ شبكة المنصور الدولية للاستشارات | 2026</center>", unsafe_allow_html=True)
