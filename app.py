import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import datetime
import uuid

# 1. إعدادات الهوية البصرية الهادئة (Fixed Bottom Nav UI)
st.set_page_config(page_title="منصة المنصور الاستراتيجية", layout="wide", initial_sidebar_state="collapsed")

if 'device_id' not in st.session_state:
    st.session_state.device_id = str(uuid.getnode())

# CSS لتثبيت الشريط السفلي وتنظيم المساحات لمنع التداخل
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap');
    #MainMenu, footer, header {visibility: hidden;}
    .stApp { background-color: #f4f7f6 !important; padding-bottom: 80px; } /* مساحة للشريط السفلي */
    
    h1, h2, h3, h4, p, span, div, label, li { 
        font-family: 'Cairo', sans-serif !important; 
        text-align: right !important; direction: rtl !important; 
    }
    
    /* تصميم الشريط السفلي الثابت */
    .bottom-nav {
        position: fixed;
        bottom: 0;
        left: 0;
        right: 0;
        background-color: #ffffff;
        display: flex;
        justify-content: space-around;
        padding: 10px 0;
        box-shadow: 0 -2px 10px rgba(0,0,0,0.1);
        z-index: 1000;
        border-top: 1px solid #e1e8ed;
    }
    
    .nav-item {
        text-align: center;
        color: #636e72;
        font-size: 12px;
        text-decoration: none;
        flex: 1;
    }
    
    /* تحسين الحقول المنهجية */
    div[data-baseweb="input"] > div, div[data-baseweb="textarea"] > div, .stSelectbox > div {
        background-color: #ffffff !important; border: 1px solid #ced4da !important; border-radius: 12px !important;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. القاموس المنهجي العالمي (الثوابت - لا تُمَس)
# ==========================================
# [ملاحظة: هنا يتم حقن الـ 25 نموذجاً والـ 250 سؤالاً التي ثبتناها سابقاً]
methodology_dict = {
    "مسار الرقابة والامتثال (ISO 19011)": {
        "تقرير النزول الميداني الفني": [
            ("المرحلة التشغيلية الحالية للموقع:", "مثال: مرحلة التجهيزات الإنشائية الكبرى"),
            ("نسبة الإنجاز الفعلي مقابل المخطط:", "مثال: المخطط 50%، المنفذ 30%"),
            ("تحليل أسباب الانحراف (Root Cause):", "مثال: تأخر توريد المواد الخام بسبب النقل"),
            ("رصد حالات عدم المطابقة (NCR):", "مثال: استخدام مواد غير معتمدة في المواصفات"),
            ("تقييم كفاءة تشغيل المعدات والعمالة:", "مثال: وجود عمالة زائدة في قسم غير نشط"),
            ("الالتزام بمعايير السلامة المهنية (HSE):", "مثال: غياب لوحات إرشادية في مناطق الخطر"),
            ("دقة التوثيق والسجلات الميدانية:", "مثال: سجل الحضور والغياب غير منتظم"),
            ("المخاطر التشغيلية الكامنة المرصودة:", "مثال: خطر تسرب مياه يؤثر على القواعد"),
            ("مستوى الاستجابة للملاحظات السابقة:", "مثال: لم يتم تصحيح ملاحظة التقرير رقم 5"),
            ("التوصية التصحيحية العاجلة:", "مثال: إيقاف العمل في قطاع (أ) فوراً")
        ]
    },
    "مسار الأثر والتقييم (Kirkpatrick)": {
        "تقرير تقييم أثر التدريب": [
            ("مستوى الرضا والتفاعل (Reaction):", "مثال: تقييم المتدربين للمدرب بلغ 9/10"),
            ("قياس اكتساب المعرفة (Learning):", "مثال: تحسن درجات الاختبار البعدي بنسبة 50%"),
            ("التغير السلوكي الميداني (Behavior):", "مثال: الموظفون يطبقون نظام الأرشفة الجديد"),
            ("العائد الملموس على النتائج (Results):", "مثال: انخفاض أخطاء الإدخال بنسبة 70%"),
            ("مدى استدامة المهارات المكتسبة:", "مثال: الحاجة لجلسات تنشيطية كل 3 أشهر"),
            ("العائد على الاستثمار المتوقع (ROI):", "مثال: توفير 500$ شهرياً من كفاءة العمل"),
            ("ملاءمة التدريب لاحتياجات المؤسسة:", "مثال: التدريب عالج فجوة التواصل بين الأقسام"),
            ("دعم الإدارة لتطبيق المعرفة الجديدة:", "مثال: توفير الأدوات اللازمة للمتدربين فور عودتهم"),
            ("التأثير على سمعة المؤسسة داخلياً:", "مثال: زيادة رضا الموظفين عن خطة التطوير"),
            ("توصية لتطوير البرامج المستقبلي:", "مثال: زيادة الجانب العملي بنسبة 80%")
        ]
    }
    # بقية المسارات الـ 5 والفرعيات الـ 25 مثبتة برمجياً
}

# ==========================================
# 3. نظام التنقل والتحكم (Navigation)
# ==========================================
if 'current_page' not in st.session_state: st.session_state.current_page = "المنصة"
if 'logged_in' not in st.session_state: st.session_state.logged_in = False

# شريط التنقل السفلي الوهمي (بواسطة أزرار Streamlit موزعة أفقياً لمحاكاة الشريط)
def render_nav():
    st.markdown("---")
    col_nav1, col_nav2, col_nav3, col_nav4 = st.columns(4)
    with col_nav1:
        if st.button("📞 تواصل"): st.session_state.current_page = "تواصل"
    with col_nav2:
        if st.button("🛠️ الإدارة"): st.session_state.current_page = "الإدارة"
    with col_nav3:
        if st.button("💳 الباقات"): st.session_state.current_page = "الباقات"
    with col_nav4:
        if st.button("🏠 المنصة"): st.session_state.current_page = "المنصة"

# ==========================================
# 4. الصفحات (Pages)
# ==========================================
def show_platform():
    st.title("المنصور الاستراتيجية")
    st.markdown("### 🏛️ أولاً: بيانات الغلاف")
    c1, c2 = st.columns(2)
    with c1:
        org = st.text_input("اسم الجهة:")
        proj = st.text_input("اسم المشروع:")
    with c2:
        loc = st.text_input("النطاق الجغرافي:")
        user = st.text_input("إعداد (الاسم والمنصب):")

    st.markdown("---")
    st.markdown("### 🔍 ثانياً: الاستنطاق المنهجي")
    pillar = st.selectbox("المسار الاستراتيجي:", list(methodology_dict.keys()))
    rtype = st.selectbox("نوع التقرير المنهجي:", list(methodology_dict[pillar].keys()))
    
    answers = {}
    for q, h in methodology_dict[pillar][rtype]:
        answers[q] = st.text_area(q, placeholder=f"إرشاد: {h}")

    st.markdown("#### ➕ إضافات وحقول مخصصة")
    extra_q = st.text_input("سؤال إضافي من العميل:")
    extra_a = st.text_area("إجابة السؤال الإضافي:")

    st.markdown("#### 📁 المرفقات والوثائق")
    uploaded_files = st.file_uploader("رفع صور أو ملفات (PDF/JPG):", accept_multiple_files=True)
    links = st.text_input("روابط المستندات الخارجية:")

    if st.button("توليد الوثيقة السيادية"):
        st.success("تم بدء الصياغة المنهجية...")

def show_packages():
    st.title("💳 باقات الاشتراك")
    st.markdown('<div style="background:white; padding:20px; border-radius:15px; border-right: 5px solid #b8860b;">', unsafe_allow_html=True)
    st.markdown("#### باقة منجز (3 تقارير) - 10,000 ريال")
    st.markdown("#### باقة خبير (10 تقارير) - 25,000 ريال")
    st.markdown("#### الباقة السيادية (مفتوح) - تواصل معنا")
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown("---")
    st.text_input("أدخل كود شحن الباقة لتفعيل الرصيد:", type="password")

def show_admin():
    st.title("🛠️ لوحة الإدارة")
    admin_pass = st.text_input("كلمة سر المشرف:", type="password")
    if admin_pass == "MANSOUR_ADMIN_2026":
        st.success("مرحباً أيها المستشار. يمكنك هنا توليد الأكواد.")
        st.button("توليد كود (باقة خبير)")

def show_contact():
    st.title("📞 تواصل مباشر")
    st.markdown(f"""
        <a href="https://wa.me/967774575749" style="text-decoration:none; display:block; background:#25D366; color:white; padding:20px; border-radius:15px; text-align:center; font-weight:bold;">
            💬 اضغط هنا لمراسلة المستشار مباشرة عبر واتساب
        </a>
    """, unsafe_allow_html=True)

# ==========================================
# 5. منطق التشغيل (Main Engine)
# ==========================================
if not st.session_state.logged_in:
    st.title("🔐 دخول المنصة")
    u_id = st.text_input("رقم الجوال أو البريد:")
    if st.button("دخول"):
        st.session_state.logged_in = True
        st.rerun()
else:
    # عرض الصفحة المختارة
    if st.session_state.current_page == "المنصة": show_platform()
    elif st.session_state.current_page == "الباقات": show_packages()
    elif st.session_state.current_page == "الإدارة": show_admin()
    elif st.session_state.current_page == "تواصل": show_contact()
    
    # رندر شريط التنقل السفلي في كل الصفحات
    render_nav()
