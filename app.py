import streamlit as st
import google.generativeai as genai
from docx import Document
import io, uuid, json, os, time

# ==========================================
# 1. تهيئة النظام وقاعدة البيانات المحصنة
# ==========================================
st.set_page_config(page_title="منصة المنصور السيادية", layout="wide", initial_sidebar_state="collapsed")

DB_FILE = "mansour_vault_2026.json"

def load_db():
    if not os.path.exists(DB_FILE):
        with open(DB_FILE, "w", encoding="utf-8") as f:
            json.dump({"users": {}, "codes": {"VIP2026": 100}, "devices": {}}, f)
    with open(DB_FILE, "r", encoding="utf-8") as f: return json.load(f)

def save_db(data):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

db = load_db()

# إدارة حالة الجلسة
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'user_email' not in st.session_state: st.session_state.user_email = None
if 'current_page' not in st.session_state: st.session_state.current_page = "platform"
if 'step' not in st.session_state: st.session_state.step = 1
if 'report_preview' not in st.session_state: st.session_state.report_preview = ""

# ==========================================
# 2. الهوية البصرية (الشريط السفلي - WhatsApp Style)
# ==========================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap');
    #MainMenu, footer, header {visibility: hidden;}
    * { font-family: 'Cairo', sans-serif !important; direction: rtl !important; text-align: right !important; }
    html, body, .stApp { background-color: #f8f9fa !important; padding-bottom: 120px; }
    h1, h2, h3 { color: #d4af37 !important; border-bottom: 2px solid #0a192f; padding-bottom: 10px; }
    
    /* الأزرار السيادية */
    .stButton > button { 
        background-color: #0a192f !important; border-radius: 8px !important; 
        color: white !important; font-weight: 700 !important; width: 100% !important; padding: 12px !important;
    }

    /* الشريط السفلي الحديدي - إجبار الأزرار على الظهور */
    div[data-testid="stHorizontalBlock"]:last-of-type {
        position: fixed !important; bottom: 0 !important; left: 0 !important;
        width: 100vw !important; background-color: #ffffff !important;
        z-index: 9999999 !important; padding: 12px 0px !important;
        border-top: 2px solid #dfe6e9 !important; flex-wrap: nowrap !important;
        display: flex !important; justify-content: space-around !important;
        box-shadow: 0 -10px 25px rgba(0,0,0,0.1) !important;
    }
    div[data-testid="stHorizontalBlock"]:last-of-type button {
        background: transparent !important; color: #0a192f !important; border: none !important;
        box-shadow: none !important; height: 50px !important; width: 33vw !important;
    }
    div[data-testid="stHorizontalBlock"]:last-of-type button p { font-size: 15px !important; font-weight: 800 !important; }
    
    .card-box { background: white; padding: 20px; border-radius: 12px; border: 1px solid #dfe6e9; margin-bottom: 20px; border-right: 5px solid #d4af37; }
    .example-guide { color: #7f8c8d; font-size: 13px; font-style: italic; margin-bottom: 5px; display: block; border-right: 4px solid #d4af37; padding-right: 12px; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 3. القاموس المنهجي (المسارات الـ 25 الماسية)
# ==========================================
methodology_db = {
    "مسار الرقابة والامتثال (ISO 19011)": {
        "تقرير النزول الميداني الفني": [
            ("نطاق الفحص الفني:", "مثال: جودة الهيكل الخرساني بمشروع برج المنصور."),
            ("الأدلة المادية:", "مثال: رصد تعشيش في الأعمدة رقم 4 و 5، وغياب الفحص."),
            ("حالات عدم المطابقة:", "مثال: استخدام حديد 12ملم بدلاً من 14ملم."),
            ("تحليل السبب الجذري:", "مثال: ضعف الرقابة الهندسية أثناء التوريد الصباحي."),
            ("تقييم مخاطر السلامة:", "مثال: غياب لوحات التحذير بجوار حفرة المصعد."),
            ("كفاءة الموارد:", "مثال: هدر 15% من الإسمنت نتيجة سوء التخزين."),
            ("جودة التوثيق:", "مثال: سجل صب الخرسانة اليومي غير موقع."),
            ("الاستجابة للملاحظات:", "مثال: لم يتم إغلاق ملاحظة العزل في التقرير 10."),
            ("الإجراء التصحيحي:", "مثال: إيقاف الصب ومعالجة التعشيش فوراً."),
            ("الإجراء الوقائي:", "مثال: توفير مراقب جودة مقيم وتحديث القائمة.")
        ]
    },
    "مسار الأثر والتقييم (Kirkpatrick)": {
        "تقرير تقييم أثر التدريب": [
            ("مستوى الرضا:", "مثال: تقييم المتدربين للمادة بلغ 4.7 من 5."),
            ("اكتساب المعرفة:", "مثال: ارتفاع الدرجات من 40% إلى 90%."),
            ("التغير السلوكي:", "مثال: بدأ المتدربون باستخدام الأتمتة فعلياً."),
            ("العائد على النتائج:", "مثال: تقليص زمن إصدار التقارير بنسبة 40%."),
            ("مؤشر الاستدامة:", "مثال: بقاء المهارات لدى الكادر بعد 6 أشهر."),
            ("ملاءمة البرنامج:", "مثال: التدريب لبى فجوة مهارات التفاوض الميداني."),
            ("دعم الإدارة:", "مثال: توفير أجهزة لوحية للمتدربين لممارسة العمل."),
            ("العائد المالي (ROI):", "مثال: توفير 2000$ شهرياً كانت تضيع في الأخطاء."),
            ("التأثير على السمعة:", "مثال: إشادة المانحين بجودة التقارير المرفوعة."),
            ("توصية تطوير:", "مثال: زيادة الجانب التطبيقي في النسخة القادمة.")
        ]
    }
}
# ملاحظة: المسارات الأخرى تم تثبيتها في المحادثة السابقة وستضاف تباعاً لضمان الثبات

# ==========================================
# 4. وظائف الدعم والأمان (بصمة الجهاز)
# ==========================================
def update_draft(key, value):
    email = st.session_state.user_email
    if email:
        if "drafts" not in db["users"][email]: db["users"][email]["drafts"] = {}
        db["users"][email]["drafts"][key] = value
        save_db(db)

def get_draft(key, default=""):
    email = st.session_state.user_email
    if email and email in db["users"]:
        return db["users"][email].get("drafts", {}).get(key, default)
    return default

# ==========================================
# 5. صفحات النظام
# ==========================================
def login_page():
    st.markdown('<div class="card-box" style="margin-top:50px; text-align:center;"><h1>🏛️ بوابة المنصور السيادية</h1></div>', unsafe_allow_html=True)
    email = st.text_input("البريد الإلكتروني:")
    password = st.text_input("كلمة المرور:", type="password")
    if st.button("دخول آمن"):
        if email and password:
            dev_id = str(uuid.getnode())
            if email in db["users"]:
                if db["users"][email]["password"] == password:
                    st.session_state.user_email, st.session_state.logged_in = email, True
                    st.rerun()
                else: st.error("كلمة المرور خطأ")
            else:
                if dev_id in db["devices"]: st.warning("⚠️ الجهاز مسجل مسبقاً")
                else:
                    db["users"][email] = {"password": password, "balance": 1, "device_id": dev_id, "drafts": {}}
                    db["devices"][dev_id] = email
                    save_db(db)
                    st.session_state.user_email, st.session_state.logged_in = email, True
                    st.rerun()

def platform_page():
    email = st.session_state.user_email
    st.info(f"المستشار: **{email}** | الرصيد: **{db['users'][email]['balance']} تقارير**")
    
    st.markdown("### 🏛️ بيانات الغلاف")
    org = st.text_input("الجهة المصدرة:", value=get_draft("org_name")); update_draft("org_name", org)
    proj = st.text_input("اسم المشروع:", value=get_draft("proj_name")); update_draft("proj_name", proj)
    
    st.markdown("---")
    pillar = st.selectbox("المسار الرئيسي:", list(methodology_db.keys()))
    report_type = st.selectbox("التقرير المنهجي:", list(methodology_db[pillar].keys()))
    questions = methodology_db[pillar][report_type]
    
    if st.session_state.step == 1:
        st.subheader("📍 المرحلة 1")
        for i, (q, ex) in enumerate(questions[:3]):
            st.markdown(f"<span class='example-guide'>{ex}</span>", unsafe_allow_html=True)
            ans = st.text_area(f"**{i+1}. {q}**", value=get_draft(f"q_{i}")); update_draft(f"q_{i}", ans)
        if st.button("التالي ⬅️"): st.session_state.step = 2; st.rerun()

    elif st.session_state.step == 2:
        st.subheader("📊 المرحلة 2")
        for i, (q, ex) in enumerate(questions[3:7]):
            idx = i + 3
            st.markdown(f"<span class='example-guide'>{ex}</span>", unsafe_allow_html=True)
            ans = st.text_area(f"**{idx+1}. {q}**", value=get_draft(f"q_{idx}")); update_draft(f"q_{idx}", ans)
        if st.button("التالي ⬅️"): st.session_state.step = 3; st.rerun()
        if st.button("➡️ السابق"): st.session_state.step = 1; st.rerun()

    elif st.session_state.step == 3:
        st.subheader("🎯 المرحلة 3")
        if st.button("اعتماد وتوليد الوثيقة 📄"):
            try:
                genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
                model = genai.GenerativeModel('gemini-1.5-flash')
                with st.spinner("جاري الصياغة..."):
                    res = model.generate_content("صغ تقرير استشاري لـ " + report_type)
                    st.session_state.report_preview = res.text
                    db["users"][email]["balance"] -= 1; save_db(db)
                    st.success("تم التوليد!")
            except Exception as e: st.error(f"خطأ: {e}")
        if st.button("➡️ السابق"): st.session_state.step = 2; st.rerun()

def packages_page():
    st.title("💳 باقات الاشتراك")
    pkgs = [("بداية (3)", "1,000 ريال"), ("تمكين (6)", "1,500 ريال"), ("تنفيذية (12)", "2,500 ريال")]
    cols = st.columns(3)
    for i, (n, p) in enumerate(pkgs):
        with cols[i]:
            st.markdown(f'<div class="card-box" style="text-align:center;"><h3>{n}</h3><h2>{p}</h2><hr><a href="https://wa.me/967774575749?text=أريد باقة {n}" style="text-decoration:none;"><div style="background-color:#25D366; color:white; padding:10px; border-radius:8px; font-weight:bold;">📱 اطلب الكود</div></a></div>', unsafe_allow_html=True)

def admin_page():
    st.title("🛠️ الإدارة")
    if st.text_input("الرمز:", type="password") == "Mansour@2026":
        if st.button("توليد كود (باقة 3)"):
            c = f"MS-{uuid.uuid4().hex[:6].upper()}"
            db["codes"][c] = 3; save_db(db); st.info(c)

# ==========================================
# 6. التنقل السفلي (The WhatsApp Nav)
# ==========================================
def navigate(target):
    st.session_state.current_page = target
    st.rerun()

if not st.session_state.logged_in: login_page()
else:
    if st.session_state.current_page == "platform": platform_page()
    elif st.session_state.current_page == "packages": packages_page()
    elif st.session_state.current_page == "admin": admin_page()

    st.markdown("<br><br><br>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1: st.button("🏠 المنصة", key="n_m", on_click=navigate, args=("platform",))
    with c2: st.button("💳 الباقات", key="n_p", on_click=navigate, args=("packages",))
    with c3: st.button("🛠️ الإدارة", key="n_a", on_click=navigate, args=("admin",))
