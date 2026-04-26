import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import datetime

# 1. المعمارية البصرية الثابتة
st.set_page_config(page_title="المنصور الاستراتيجية", layout="wide")
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;700&display=swap');
    #MainMenu, footer, header {visibility: hidden;}
    .stApp { background-color: #0c0c0c !important; }
    h1, h2, h3, h4, p, span, div, label, li { font-family: 'Cairo', sans-serif !important; text-align: right !important; direction: rtl !important; color: #ffffff; }
    h1, h2, h3 { color: #D4AF37 !important; margin-bottom: 5px; }
    input, textarea, .stSelectbox > div { background-color: #1a1a1a !important; border: 1px solid #D4AF37 !important; color: #ffffff !important; }
    .stButton > button { background-color: #D4AF37 !important; color: #0c0c0c !important; font-weight: 700 !important; width: 100% !important; padding: 15px !important; }
    .package-box { border: 2px solid #D4AF37; padding: 15px; border-radius: 8px; background-color: #111; margin-bottom: 20px; text-align: center; }
</style>
""", unsafe_allow_html=True)

# 2. القاموس المنهجي (The Scientific Database)
# هنا تم وضع 10 أسئلة لكل نموذج بناءً على المنهجيات العالمية
methodology_db = {
    "مسار الرقابة (ISO 19011)": {
        "تقرير النزول الميداني الفني": [
            ("نطاق الزيارة والمرحلة التشغيلية:", "مثال: معاينة صب القواعد في مشروع مجمع المنصور"),
            ("نسبة الإنجاز الفعلي vs المستهدف (%):", "مثال: المستهدف 40%، الفعلي 25%"),
            ("مسببات الانحراف الجذرية (Root Causes):", "مثال: نقص الكادر الفني المتخصص وتأخر التوريدات"),
            ("حالات عدم المطابقة للمواصفات (NCR):", "مثال: استخدام إسمنت مقاوم بدلاً من العادي في الأعمدة"),
            ("كفاءة استخدام الموارد والمعدات:", "مثال: وجود رافعة معطلة بالموقع تستهلك إيجاراً يومياً"),
            ("مستوى الالتزام بمعايير السلامة (HSE):", "مثال: عدم توفر حواجز حماية في المناطق المرتفعة"),
            ("المخاطر الكامنة المرصودة:", "مثال: خطر انهيار التربة في الجهة الشرقية بسبب الأمطار"),
            ("جودة التوثيق الورقي والمكتبي بالموقع:", "مثال: سجل الزيارات غير محدث والخرائط تالفة"),
            ("مدى استجابة المقاول للتوجيهات السابقة:", "مثال: تم تجاهل 3 ملاحظات مسجلة في المحضر السابق"),
            ("التوصية التصحيحية العاجلة (القرار):", "مثال: إيقاف العمل في قطاع (ب) حتى استبدال المواد")
        ],
        "تقرير تدقيق الامتثال الإداري": [
            ("المعيار القانوني المرجعي المحقق:", "مثال: اللائحة التنفيذية لقانون العمل مادة 40"),
            ("تحليل الفجوة الإجرائية (Gap Analysis):", "مثال: عدم وجود توصيف وظيفي معتمد لـ 30% من الطاقم"),
            # يتم إكمال الـ 10 أسئلة هنا بنفس النمط...
        ]
    },
    "مسار الأثر (Kirkpatrick Model)": {
        "تقرير تقييم أثر التدريب والتمكين": [
            ("مستوى رد الفعل والرضا الأولي:", "مثال: تقييم المتدربين للمحتوى بلغ 9.5/10"),
            ("قياس اكتساب المعرفة (Pre/Post Test):", "مثال: تحسن متوسط درجات الاختبار من 40% إلى 85%"),
            ("التغير السلوكي الملموس في بيئة العمل:", "مثال: التزام الموظفين باستخدام نظام الأرشفة الجديد"),
            ("التحسن في مؤشرات الأداء (KPIs):", "مثال: انخفاض زمن الاستجابة للشكاوى بنسبة 40%"),
            ("العائد على الاستثمار المتوقع (ROI):", "مثال: توفير 2000$ شهرياً من هدر الورق"),
            # يتم إكمال الـ 10 أسئلة هنا بنفس النمط...
        ]
    }
}

# 3. واجهة بوابة الباقات (صدر الصفحة)
st.title("المنصور الاستراتيجية")
st.markdown('<div class="package-box">', unsafe_allow_html=True)
st.subheader("💳 بوابة شحن الباقات السيادية")
col_p1, col_p2 = st.columns([2, 1])
with col_p1:
    activation_code = st.text_input("أدخل كود تفعيل الباقة:", type="password")
with col_p2:
    st.write("**تواصل للشحن:** 774575749")
    st.write("**الباقات:** [منجز | خبير | سيادي]")
st.markdown('</div>', unsafe_allow_html=True)

# 4. الطبقة الإدارية (الغلاف)
st.markdown("### 🏛️ أولاً: البيانات الإدارية")
c1, c2 = st.columns(2)
with c1:
    org_name = st.text_input("الجهة المصدرة للتقرير:", placeholder="مثال: مؤسسة شباب اليمن")
    proj_name = st.text_input("اسم المشروع / المهمة:", placeholder="مثال: مشروع الاستجابة الطارئة")
with c2:
    loc_name = st.text_input("النطاق الجغرافي:", placeholder="مثال: تعز - مديرية المظفر")
    user_info = st.text_input("مُعد الوثيقة (الاسم والمنصب):")

st.markdown("---")

# 5. الطبقة المنهجية (المسارات والفرعيات والأمثلة)
st.markdown("### 🔍 ثانياً: الاستنطاق المنهجي")
selected_pillar = st.selectbox("1. حدد المسار الرئيسي:", list(methodology_db.keys()))
selected_report = st.selectbox("2. حدد التقرير التخصصي (الفرعي):", list(methodology_db[selected_pillar].keys()))

st.success(f"المنهجية المطبقة حالياً: {selected_report}")

# توليد الـ 10 أسئلة مع أمثلتها
user_answers = {}
questions_list = methodology_db[selected_pillar][selected_report]

for i, (q_text, q_hint) in enumerate(questions_list):
    user_answers[q_text] = st.text_area(f"{i+1}. {q_text}", placeholder=f"إرشاد: {q_hint}", height=100)

st.markdown("---")

# 6. الطبقة الاعتمادية
st.markdown("### 📝 ثالثاً: الخواتيم والاعتماد")
final_recs = st.text_area("التوصيات والمقترحات الاستراتيجية للإدارة العليا:")
files_list = st.text_input("الملاحق والشواهد المرفقة:")

# 7. المحرك التنفيذي (Gemini Engine)
if st.button("اعتماد وتوليد الوثيقة السيادية"):
    if not activation_code:
        st.error("⚠️ يجب إدخال كود التفعيل لتشغيل المحرك.")
    elif not (org_name and proj_name and user_info):
        st.warning("⚠️ يرجى استكمال بيانات الغلاف.")
    else:
        try:
            genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
            model = genai.GenerativeModel('gemini-1.5-pro')
            
            # صهر البيانات في برومبت منهجي
            data_feed = "\n".join([f"{k} {v}" for k, v in user_answers.items() if v])
            prompt = f"""
            بصفتك مستشاراً تنفيذياً عالمياً، صغ تقرير '{selected_report}' لجهة '{org_name}' حول '{proj_name}'.
            الموقع: {loc_name}. إعداد: {user_info}.
            المنهجية المتبعة: {selected_pillar}.
            البيانات الميدانية:
            {data_feed}
            التوصيات: {final_recs}
            الملاحق: {files_list}
            
            [التعليمات]: صغ التقرير بلغة رصينة، رسمية، ركز على النتائج والأرقام. ابدأ بالغلاف الرسمي ثم الملخص ثم التحليل ثم التوصيات.
            """
            
            with st.spinner("جاري المعالجة المنهجية..."):
                response = model.generate_content(prompt)
                st.info(response.text)
                
                # تصدير Word
                doc = Document()
                doc.add_heading(f"{org_name} | {selected_report}", 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
                doc.add_paragraph(f"المشروع: {proj_name}\nالمكان: {loc_name}\nإعداد: {user_info}\nالتاريخ: {datetime.date.today()}").alignment = WD_ALIGN_PARAGRAPH.RIGHT
                for line in response.text.split('\n'):
                    if line.strip():
                        p = doc.add_paragraph(line.strip())
                        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                
                bio = io.BytesIO()
                doc.save(bio)
                st.download_button("تحميل الوثيقة المعتمدة (Word)", bio.getvalue(), file_name=f"Report_{proj_name}.docx")
        except Exception as e:
            st.error(f"عطل في المحرك: {e}")
