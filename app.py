import streamlit as st
import google.generativeai as genai
from io import BytesIO
from docx import Document
from datetime import datetime

# 1. الهندسة البصرية وتوحيد الخط (Cairo)
st.set_page_config(page_title="منصة المنصور AI - التميز الاستراتيجي", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');
    
    /* توحيد الخط في كامل المنصة */
    html, body, [class*="st-"], * { font-family: 'Cairo', sans-serif !important; direction: rtl; text-align: right; }
    
    .stApp { background-color: #f4f7f9; }
    #MainMenu, footer, header { visibility: hidden; }
    
    /* الحاوية الرئيسية */
    .main-card { 
        background: white; border-top: 12px solid #1e3a8a; 
        padding: 45px; border-radius: 20px; 
        box-shadow: 0 15px 40px rgba(0,0,0,0.1); margin-top: -60px; 
    }
    
    .brand-title { color: #1e3a8a; font-weight: 900; font-size: 2.8rem; text-align: center; margin-bottom: 5px; }
    .brand-subtitle { color: #d4af37; text-align: center; font-weight: 700; font-size: 1rem; margin-bottom: 40px; letter-spacing: 2px; }
    
    /* تنسيق الحقول والأسئلة */
    .section-header { 
        background: #1e3a8a; color: #ffffff; padding: 15px 25px; 
        border-radius: 12px; font-weight: 700; margin: 35px 0 20px 0;
        box-shadow: 0 4px 10px rgba(30, 58, 138, 0.2); font-size: 1.2rem;
    }
    
    .q-label { color: #1e293b; font-weight: 800; font-size: 1.05rem; border-right: 5px solid #d4af37; padding-right: 12px; margin-top: 25px; display: block; }
    .hint-box { color: #64748b; font-size: 0.85rem; margin-bottom: 10px; background: #f8fafc; padding: 8px; border-radius: 5px; }
    
    /* الأزرار الملكية */
    .stButton>button { 
        background: linear-gradient(135deg, #1e3a8a 0%, #152e6d 100%) !important; color: #ffffff !important; 
        font-weight: 800 !important; height: 60px !important; font-size: 1.2rem !important;
        border-radius: 15px !important; border: none !important; width: 100%; transition: 0.4s all;
    }
    .stButton>button:hover { transform: translateY(-3px); box-shadow: 0 8px 20px rgba(30, 58, 138, 0.3); border-bottom: 5px solid #d4af37 !important; }
    
    /* أزرار التواصل */
    .whatsapp-btn {
        background-color: #25d366; color: white !important; padding: 15px 30px; 
        border-radius: 50px; text-decoration: none; font-weight: 900; 
        display: inline-flex; align-items: center; justify-content: center; gap: 10px; margin-top: 20px;
    }
</style>
""", unsafe_allow_html=True)

# 2. الهيكل العالمي المتطور للتقارير (أسئلة عميقة + أمثلة داخلية)
STRATEGIC_FRAMEWORK = {
    "📑 تقرير إنجاز دوري | Progress Report": [
        ("الملخص التنفيذي للأداء", "مثال: تم تحقيق 85% من الأهداف المخططة للفترة الحالية بنجاح تام..."),
        ("تحليل الأنشطة والمهام المنفذة", "مثال: عقد ورشتي عمل، توريد 50 وحدة تقنية، وإتمام 4 زيارات ميدانية..."),
        ("إدارة الانحرافات عن الخطة", "مثال: تأخر بند التدريب لمدة أسبوع بسبب إجراءات التصاريح وتم تداركه..."),
        ("تحليل الموارد والاستهلاك المالي", "مثال: تم صرف 30% من الميزانية التشغيلية بما يتوافق مع المخرجات..."),
        ("التحديات والحلول التصحيحية", "مثال: واجهنا ضعف استجابة المورد، فتم تفعيل قائمة الموردين البدلاء..."),
        ("الأولويات الاستراتيجية القادمة", "مثال: البدء في مرحلة التقييم النهائي وجمع بيانات المستفيدين...")
    ],
    "💰 دراسة جدوى استثمارية | Feasibility Study": [
        ("تحليل الفجوة والاحتياج السوقي", "مثال: يوجد عجز بنسبة 40% في خدمات الطاقة المتجددة في منطقة تعز..."),
        ("المواصفات الفنية والمتطلبات", "مثال: يتطلب المشروع مساحة 500 متر مربع، و3 فنيين متخصصين، ومولد هجين..."),
        ("النمذجة المالية وتوقعات الدخل", "مثال: العائد المتوقع يبدأ من السنة الثانية بنسبة نمو 15% سنوياً..."),
        ("تحليل الحساسية ونقطة التعادل", "مثال: ستتم استعادة رأس المال (ROI) خلال 18 شهراً من تاريخ التدشين..."),
        ("تحليل SWOT (المنافسة والفرص)", "مثال: قوتنا في التكنولوجيا الحديثة، والتهديد يكمن في تقلب أسعار الصرف..."),
        ("قرار الاستثمار والتوصيات", "مثال: المشروع مجدٍ اقتصادياً ويُنصح بالبدء في المرحلة التمهيدية...")
    ],
    "🎓 تقرير ختامي لتدريب | Capacity Building": [
        ("المنهجية والأهداف التدريبية", "مثال: تزويد المتدربين بمهارات الإدارة الاستراتيجية وكتابة التقارير..."),
        ("تحليل نتائج القبلي والبعدي", "مثال: ارتفع مستوى المعرفة لدى المشاركين من 40% إلى 95% بعد الدورة..."),
        ("تقييم كفاءة المدرب واللوجستيات", "مثال: حصل المحتوى العلمي على تقييم 4.8/5 من قبل المشاركين..."),
        ("تفاعل المشاركين وقصص النجاح", "مثال: قام أحد المتدربين بتطبيق خطة تحسين فورية في قسمه أثناء التدريب..."),
        ("توصيات استدامة الأثر المهني", "مثال: عقد جلسات تنشيطية كل 3 أشهر لمتابعة تطبيق المهارات المكتسبة...")
    ],
    "🔍 تقرير متابعة وتقييم | M&E Report": [
        ("قياس مؤشرات الأداء (KPIs)", "مثال: تم الوصول لـ 500 مستفيد من أصل 450 مستهدف (نسبة 111%)..."),
        ("جودة المخرجات والامتثال المعياري", "مثال: جميع المواد الموزعة تطابق مواصفات الجودة المعتمدة دولياً..."),
        ("تحليل رضا المستفيدين", "مثال: أظهرت المقابلات أن 90% من الجمهور راضٍ عن سرعة الاستجابة..."),
        ("الدروس المستفادة والنمو المؤسسي", "مثال: الاعتماد على المجتمع المحلي قلل من تكاليف النقل بنسبة 20%..."),
        ("توصيات التطوير الاستراتيجي", "مثال: توسيع نطاق المشروع ليشمل المديريات المجاورة في المرحلة القادمة...")
    ]
}

# 3. بناء الواجهة الاستشارية
st.markdown('<div class="main-card">', unsafe_allow_html=True)
st.markdown('<h1 class="brand-title">منصة المنصور الاستراتيجية AI</h1>', unsafe_allow_html=True)
st.markdown('<p class="brand-subtitle">نظام الصياغة والتحليل المؤسسي العالمي - الإصدار 2026</p>', unsafe_allow_html=True)

# تفعيل الذكاء الاصطناعي
try:
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")
except:
    st.error("⚠️ يرجى ضبط مفتاح API في الإعدادات.")

# اختيار نوع التقرير
report_type = st.selectbox("🚀 حدد تخصص التقرير المطلوب لتفعيل المنهجية:", list(STRATEGIC_FRAMEWORK.keys()))

st.markdown('<div class="section-header">🏷️ أولاً: البيانات التعريفية للمستند</div>', unsafe_allow_html=True)
c1, c2 = st.columns(2)
p_name = c1.text_input("اسم المشروع / النشاط الرئيسي *")
p_agency = c2.text_input("الجهة المنفذة (المؤسسة / الشركة)")
p_target = c1.text_input("الجهة الموجه إليها التقرير")
p_loc = c2.text_input("نطاق التنفيذ والتاريخ")

# عرض الأسئلة بناءً على التخصص
st.markdown(f'<div class="section-header">🔍 ثانياً: الأسئلة الجوهرية لـ {report_type}</div>', unsafe_allow_html=True)
responses = {}
for q, example in STRATEGIC_FRAMEWORK[report_type]:
    st.markdown(f'<span class="q-label">{q}</span>', unsafe_allow_html=True)
    st.markdown(f'<div class="hint-box">💡 {example}</div>', unsafe_allow_html=True)
    responses[q] = st.text_area("", key=q, height=120, label_visibility="collapsed")

# القسم المخصص (بدون أمثلة كما طلبت)
st.markdown('<div class="section-header">➕ ثالثاً: التخصيص الإضافي</div>', unsafe_allow_html=True)
st.info("إذا كان لديك بيانات أخرى وتريد إضافتها، اضغط على الزر وخصص قسماً جديداً:")
if 'custom_fields' not in st.session_state: st.session_state.custom_fields = []
new_field = st.text_input("أضف عنوان القسم الجديد هنا:")
if st.button("خصص قسم جديد الآن"):
    if new_field: st.session_state.custom_fields.append(new_field); st.rerun()

for cf in st.session_state.custom_fields:
    st.markdown(f'**⭐ قسم مخصص: {cf}**')
    responses[cf] = st.text_area(f"أدخل بيانات {cf}...", key=f"cf_{cf}")
    if st.button(f"حذف {cf}"): st.session_state.custom_fields.remove(cf); st.rerun()

# توليد التقرير
st.write("---")
if st.button("🔥 توليد التقرير الاستراتيجي النهائي"):
    if p_name and any(responses.values()):
        with st.spinner("جاري التفكير وصهر البيانات بمعايير ISO 2145..."):
            summary = "\n".join([f"- {k}: {v}" for k, v in responses.items() if v])
            prompt = f"""
            أنت مستشار دولي رفيع المستوى. صغ لي {report_type} للمشروع {p_name}.
            البيانات: الجهة {p_agency}، الموجه لـ {p_target}، المكان {p_loc}.
            المعطيات: {summary}
            
            القواعد الذهبية:
            - لغة عربية فصحى متينة وقوية.
            - ترقيم دولي (1, 1.1, 1.2).
            - استخدام 'الصوت النشط' (Active Voice) والإيجاز الاستراتيجي.
            - صياغة ملخص تنفيذي مبهر وتوصيات قابلة للتنفيذ.
            """
            res = model.generate_content(prompt)
            st.markdown("### 🏆 التقرير الاستراتيجي المُولد:")
            st.markdown(res.text)
            st.session_state['final_doc'] = res.text
    else: st.warning("يرجى ملء اسم المشروع والأسئلة الأساسية.")

# التحميل والتواصل
if 'final_doc' in st.session_state:
    doc = Document()
    doc.add_heading(f"تقرير: {p_name}", 0)
    doc.add_paragraph(st.session_state['final_doc'])
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    
    st.download_button("💾 تحميل المستند كملف Word رسمي", bio, f"{p_name}.docx")
    
st.markdown('<center>', unsafe_allow_html=True)
st.markdown(f'<a href="https://wa.me/967774575749" class="whatsapp-btn">🟢 تواصل معنا مباشرة عبر واتساب: 774575749</a>', unsafe_allow_html=True)
st.markdown('</center>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)
st.markdown("<center style='color:#64748b; font-size:0.75rem; margin-top:20px;'>🛡️ حقوق المنصة محفوظة لشبكة المنصور الدولية للاستشارات | 2026</center>", unsafe_allow_html=True)
