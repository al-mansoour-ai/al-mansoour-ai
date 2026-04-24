import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document
from datetime import datetime

# 1. الهندسة البصرية الفخمة (UI/UX)
st.set_page_config(page_title="المنصور AI - النظام الاستشاري المتكامل", layout="centered")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
    
    /* توحيد الخط والاتجاه */
    * { font-family: 'Cairo', sans-serif !important; direction: rtl; text-align: right; }
    .stApp { background-color: #f1f5f9; }
    #MainMenu, footer, header { visibility: hidden; }

    /* الحاوية الرئيسية (Main Card) */
    .main-box { 
        background: white; border-top: 10px solid #0f172a; 
        padding: 40px; border-radius: 12px; 
        box-shadow: 0 10px 30px rgba(0,0,0,0.05); margin-top: -60px; 
    }
    
    .brand-title { color: #0f172a; font-weight: 900; font-size: 2.2rem; text-align: center; margin-bottom: 0; }
    .brand-subtitle { color: #c5a059; text-align: center; font-weight: 700; font-size: 0.9rem; margin-bottom: 35px; text-transform: uppercase; letter-spacing: 1px; }

    /* تنسيق العناوين الرشيقة */
    .section-header { 
        background: #0f172a; color: #f1f5f9; padding: 10px 20px; 
        border-radius: 6px; font-weight: 700; font-size: 1.1rem; 
        margin: 30px 0 15px 0; display: inline-block;
    }

    .q-label { color: #334155; font-weight: 800; font-size: 1rem; border-right: 4px solid #c5a059; padding-right: 12px; margin-top: 25px; display: block; }
    .hint-box { color: #64748b; font-size: 0.8rem; margin-bottom: 12px; background: #f8fafc; padding: 10px; border-radius: 6px; border: 1px solid #e2e8f0; line-height: 1.6; }

    /* الأزرار الملكية المتوازنة */
    .stButton>button { 
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%) !important; color: #ffffff !important; 
        font-weight: 700 !important; height: 55px !important; border-radius: 8px !important; 
        border: none !important; width: 100%; transition: 0.3s;
    }
    .stButton>button:hover { transform: translateY(-2px); box-shadow: 0 5px 15px rgba(15, 23, 42, 0.2); }
    
    .magic-btn button { 
        background: #fdfaf3 !important; color: #856404 !important; border: 1px dashed #c5a059 !important; 
        height: 32px !important; font-size: 0.75rem !important; margin-top: -10px;
    }

    .whatsapp-btn {
        background: #25d366; color: white !important; padding: 12px 25px; border-radius: 50px; 
        text-decoration: none; font-weight: 700; display: inline-flex; align-items: center; gap: 10px; margin-top: 20px;
    }
</style>
""", unsafe_allow_html=True)

# 2. بنك المعرفة الاستراتيجي (أسئلة عميقة + أمثلة واقعية خطوة بخطوة)
REPORT_CONFIG = {
    "📑 تقرير الإنجاز الدوري": [
        ("الملخص التنفيذي للمرحلة", "ابدأ بذكر النتيجة الإجمالية؛ مثال: 'تم بحمد الله إنجاز 95% من أنشطة الربع الأول، حيث تركز العمل على بناء البنية التحتية للبرنامج وتدريب الكوادر الميدانية...'"),
        ("منهجية العمل المتبعة", "اشرح كيف نفذتم المهام؛ مثال: 'اعتمدنا منهجية النزول الميداني المباشر، مع استخدام أدوات الرقابة الرقمية لضمان جودة المخرجات في الوقت الحقيقي...'"),
        ("تحليل الأنشطة والمنجزات", "اسرد ما تم؛ مثال: '1. عقد 4 ورش عمل استهدفت 120 موظفاً. 2. توريد المعدات التقنية لمقر المؤسسة بمواصفات كذا...'"),
        ("إدارة التحديات والحلول", "كن صريحاً؛ مثال: 'واجهنا صعوبة في الوصول للمناطق الجبلية بسبب الطقس، فتم الاستعانة بالفرق المحلية المتطوعة لتسهيل المهمة...'"),
        ("التوصيات وخطة المستقبل", "ما هي الخطوة القادمة؟ مثال: 'نوصي بزيادة الميزانية التشغيلية لبند النقل، والبدء فوراً في تدشين المرحلة الثانية من التوزيع...'")
    ],
    "💰 دراسة جدوى استثمارية": [
        ("فكرة المشروع والاحتياج", "مثال: 'يستهدف المشروع سد فجوة بنسبة 30% في قطاع التغليف بمدينة تعز، نظراً لزيادة الصادرات المحلية وغياب المصانع المتخصصة...'"),
        ("النمذجة والتقديرات المالية", "مثال: 'رأس المال المطلوب كذا، مع توقعات إيرادات شهرية تبدأ من كذا، بناءً على تحليل أسعار السوق المنافسة...'"),
        ("تحليل الحساسية والمخاطر", "مثال: 'في حال ارتفاع أسعار المواد الخام بنسبة 10%، ستتأثر نقطة التعادل لتتأخر لمدة 3 أشهر إضافية، وهذا يتطلب سيولة احتياطية...'"),
        ("الخاتمة وقرار الاستثمار", "مثال: 'بناءً على المعطيات أعلاه، نوصي بالمضي قدماً في الاستثمار لجدواه العالية في ظل غياب المنافسة الحالية...'")
    ]
}

# 3. واجهة المستخدم (التصميم المتوازن)
st.markdown('<div class="main-box">', unsafe_allow_html=True)
st.markdown('<h1 class="brand-title">منصة المنصور الاستراتيجية AI</h1>', unsafe_allow_html=True)
st.markdown('<p class="brand-subtitle">PREMIUM EXECUTIVE REPORTING SYSTEM V6.0</p>', unsafe_allow_html=True)

try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")
except:
    st.error("⚠️ يرجى ضبط المفتاح في الإعدادات")

rtype = st.selectbox("🎯 حدد تخصص التقرير لضبط المنهجية:", list(REPORT_CONFIG.keys()))

# أ. صفحة الغلاف والتعريف
st.markdown('<div class="section-header">🛡️ الغلاف والبيانات التعريفية</div>', unsafe_allow_html=True)
c1, c2 = st.columns(2)
p_name = c1.text_input("عنوان التقرير (اسم المشروع) *")
p_ref = c2.text_input("الرقم المرجعي (Ref No.)")
p_agency = c1.text_input("الجهة المُعِدّة")
p_target = c2.text_input("الجهة الموجه إليها")

# ب. خطاب الشكر والمقدمة
st.markdown('<div class="section-header">🤝 التصدير: الشكر والمقدمة</div>', unsafe_allow_html=True)
p_thanks = st.text_area("كلمة شكر وتقدير وتقديم للتقرير:", placeholder="مثال: نتوجه بخالص الشكر للجهات الداعمة على ثقتهم...")

# ج. صلب التقرير (المحاور الاستراتيجية)
st.markdown(f'<div class="section-header">🔍 صلب التقرير: محاور {rtype}</div>', unsafe_allow_html=True)
responses = {}
for i, (q, hint) in enumerate(REPORT_CONFIG[rtype]):
    st.markdown(f'<span class="q-label">{q}</span>', unsafe_allow_html=True)
    st.markdown(f'<div class="hint-box">📝 {hint}</div>', unsafe_allow_html=True)
    txt = st.text_area("", key=f"v6_{i}_{rtype}", height=120, label_visibility="collapsed")
    
    # زر التحسين الرشيق
    col_btn, col_sp = st.columns([1, 3])
    with col_btn:
        st.markdown('<div class="magic-btn">', unsafe_allow_html=True)
        if st.button(f"✨ تحسين {i+1}", key=f"btn_v6_{i}"):
            if txt:
                with st.spinner("جاري الصياغة..."):
                    res = model.generate_content(f"صغ هذا المحور بأسلوب استشاري رفيع (صوت نشط، إيجاز): {txt}")
                    st.code(res.text)
            else: st.warning("أدخل نصاً")
        st.markdown('</div>', unsafe_allow_html=True)
    responses[q] = txt

# د. الخاتمة والملاحق والاعتماد
st.markdown('<div class="section-header">📌 الخاتمة والملاحق والتوقيعات</div>', unsafe_allow_html=True)
p_concl = st.text_area("الخاتمة النهائية:", placeholder="اكتب خلاصة التقرير هنا...")
p_appendix = st.text_area("الملاحق (صور، روابط، جداول):", placeholder="مثال: مرفق صور النزول الميداني في الرابط أدناه...")

# هـ. التخصيص (إضافة أقسام أخرى)
if 'v6_extra' not in st.session_state: st.session_state.v6_extra = []
st.info("إذا كان لديك بيانات أخرى وتريد إضافتها، اضغط على الزر وخصص قسماً جديداً:")
new_sec = st.text_input("اسم القسم الإضافي:")
if st.button("خصص قسم جديد"):
    if new_sec: st.session_state.v6_extra.append(new_sec); st.rerun()

for ex in st.session_state.v6_extra:
    st.markdown(f"**⭐ قسم مخصص: {ex}**")
    responses[ex] = st.text_area(f"بيانات {ex}...", key=f"ex6_{ex}")
    if st.button(f"حذف {ex}"): st.session_state.v6_extra.remove(ex); st.rerun()

st.write("---")
col_pre, col_rev, col_app = st.columns(3)
p_pre = col_pre.text_input("أعده:")
p_rev = col_rev.text_input("راجعه:")
p_app = col_app.text_input("اعتمده:")

# و. التوليد النهائي
st.write("---")
if st.button("🚀 توليد ومعالجة التقرير الاستراتيجي الشامل"):
    if p_name and any(responses.values()):
        with st.spinner("جاري بناء الوثيقة السيادية..."):
            all_data = "\n".join([f"- {k}: {v}" for k, v in responses.items() if v])
            full_prompt = f"""
            بصفتك مستشاراً عالمياً، صغ تقريراً استراتيجياً متكاملاً.
            الغلاف: {p_name}، المرجع {p_ref}، الجهة {p_agency}، الموجه لـ {p_target}.
            خطاب التقديم: {p_thanks}
            المحاور الفنية: {all_data}
            الخاتمة: {p_concl}
            الملاحق: {p_appendix}
            هيكل الاعتماد: المعد {p_pre}، المراجع {p_rev}، المعتمد {p_app}.
            
            المعايير: ISO 2145، صوت نشط، نبرة قيادية، ترقيم احترافي.
            """
            res = model.generate_content(full_prompt)
            st.markdown(res.text)
            st.session_state['v6_out'] = res.text
    else: st.warning("يرجى ملء البيانات.")

if 'v6_out' in st.session_state:
    doc = Document()
    doc.add_heading(f"Report: {p_name}", 0)
    doc.add_paragraph(st.session_state['v6_out'])
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    st.download_button("💾 تحميل المستند الرسمي (Word)", bio, f"{p_name}.docx")

st.markdown('<center><a href="https://wa.me/967774575749" class="whatsapp-btn">🟢 تواصل معنا: 774575749</a></center>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)
