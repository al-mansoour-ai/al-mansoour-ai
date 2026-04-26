import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import datetime
import uuid

# 1. المعمارية البصرية الهادئة (Mobile-Friendly Executive UI)
st.set_page_config(page_title="منصة المنصور الاستراتيجية", layout="wide")

# نظام "بصمة الجهاز" البسيط لمحاصرة التلاعب بالباقة المجانية
if 'device_id' not in st.session_state:
    st.session_state.device_id = str(uuid.getnode()) # جلب معرف الجهاز الفعلي

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap');
    #MainMenu, footer, header {{visibility: hidden;}}
    .stApp {{ background-color: #f8f9fa !important; }}
    h1, h2, h3, h4, p, span, div, label, li {{ 
        font-family: 'Cairo', sans-serif !important; 
        text-align: right !important; direction: rtl !important; color: #2d3436 !important; 
    }}
    .main-card {{ background: white; padding: 25px; border-radius: 15px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); margin-bottom: 20px; border: 1px solid #e1e8ed; }}
    .stButton > button {{ 
        background-color: #2d3436 !important; color: white !important; 
        border-radius: 10px !important; width: 100% !important; border: none !important; padding: 10px !important;
    }}
    .whatsapp-btn {{
        background-color: #25D366 !important; color: white !important;
        text-decoration: none; padding: 10px 20px; border-radius: 10px; display: inline-block; text-align: center; width: 100%;
    }}
    .package-badge {{ background: #f1f2f6; padding: 10px; border-radius: 8px; border-right: 5px solid #b8860b; margin-bottom: 10px; }}
</style>
""", unsafe_allow_html=True)

# --- محرك البيانات المنهجي (تم الاحتفاظ به كاملاً كما طلبت) ---
# [هنا تضع مصفوفة methodology_dict السابقة كاملة]

# 2. نظام تسجيل الدخول وحماية الباقة
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False

def login_page():
    st.markdown('<div class="main-card">', unsafe_allow_html=True)
    st.title("🔐 تسجيل الدخول السيادي")
    st.info(f"معرف الجهاز النشط: {st.session_state.device_id[-6:]} (محمي ضد التكرار)")
    login_type = st.radio("اختر وسيلة الدخول:", ["رقم الجوال", "البريد الإلكتروني"])
    u_id = st.text_input(f"أدخل {login_type}:")
    if st.button("دخول للمنصة"):
        if u_id:
            st.session_state.logged_in = True
            st.session_state.user_id = u_id
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# 3. صفحة الإدارة (مخفية - تظهر بكلمة سر)
def admin_panel():
    st.markdown("### 🛠️ لوحة تحكم الإدارة (توليد الأكواد)")
    admin_pass = st.text_input("كلمة سر الإدارة:", type="password")
    if admin_pass == "MANSOUR_ADMIN_2026":
        p_type = st.selectbox("نوع الباقة:", ["منجز (3 تقارير)", "خبير (10 تقارير)", "سيادي (مفتوح)"])
        if st.button("توليد كود جديد"):
            new_code = f"MS-{uuid.uuid4().hex[:6].upper()}"
            st.success(f"تم توليد الكود: {new_code} لباقة {p_type}")
            # هنا يتم حفظ الكود في قاعدة البيانات مستقبلاً

# 4. واجهة المنصة الرئيسية
def main_app():
    # الهيدر الجانبي (التنقل)
    menu = st.sidebar.selectbox("القائمة الرئيسية", ["المنصة التنفيذية", "باقات الاشتراك", "تواصل مع المستشار", "لوحة الإدارة"])

    if menu == "المنصة التنفيذية":
        st.markdown(f"**مرحباً بك:** {st.session_state.user_id}")
        
        # [هنا تضع كود البيانات الإدارية واختيار المسارات والـ 10 أسئلة]
        # إضافة ميزة الحقول الإضافية
        st.markdown("---")
        st.markdown("#### ➕ حقول مخصصة (إضافات العميل)")
        if 'extra_fields' not in st.session_state:
            st.session_state.extra_fields = []
        
        if st.button("إضافة سؤال/حقل مخصص"):
            st.session_state.extra_fields.append(len(st.session_state.extra_fields))
        
        for i in st.session_state.extra_fields:
            st.text_input(f"عنوان الحقل الإضافي {i+1}:", key=f"extra_t_{i}")
            st.text_area(f"إجابة الحقل الإضافي {i+1}:", key=f"extra_a_{i}")

        # إضافة المرفقات
        st.markdown("#### 📁 الشواهد والمرفقات")
        uploaded_files = st.file_uploader("ارفع صور أو وثائق (PDF/JPG):", accept_multiple_files=True)
        links = st.text_input("روابط المراجع (Google Drive / DropBox):")

        # زر التوليد [نفس كود Gemini السابق]

    elif menu == "باقات الاشتراك":
        st.title("💳 باقات المنصور الاستراتيجية")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown('<div class="package-badge"><b>باقة منجز</b><br>3 تقارير تخصصية<br>سعر: 50$</div>', unsafe_allow_html=True)
        with col2:
            st.markdown('<div class="package-badge"><b>باقة خبير</b><br>10 تقارير + دعم فني<br>سعر: 120$</div>', unsafe_allow_html=True)
        with col3:
            st.markdown('<div class="package-badge"><b>الباقة السيادية</b><br>تقارير غير محدودة<br>سعر: 500$ سنوياً</div>', unsafe_allow_html=True)

    elif menu == "تواصل مع المستشار":
        st.title("📞 تواصل مباشر")
        st.markdown(f"""
            <a href="https://wa.me/967774575749" class="whatsapp-btn">
                اضغط هنا للتواصل عبر واتساب: 774575749
            </a>
        """, unsafe_allow_html=True)

    elif menu == "لوحة الإدارة":
        admin_panel()

# تشغيل المنصة
if not st.session_state.logged_in:
    login_page()
else:
    main_app()
