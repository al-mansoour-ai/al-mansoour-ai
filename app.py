import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io, uuid, json, os, time

# ==========================================
# 1. إعدادات المنصة وقاعدة البيانات (الأمان)
# ==========================================
st.set_page_config(page_title="منصة المنصور الاستراتيجية", layout="wide", initial_sidebar_state="collapsed")

DB_FILE = "mansour_strategic_final.json"

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
# 2. الهوية البصرية (حل مشكلة السكرول والشريط)
# ==========================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap');
    #MainMenu, footer, header {visibility: hidden;}
    * { font-family: 'Cairo', sans-serif !important; direction: rtl !important; text-align: right !important; }
    
    /* حل مشكلة التمرير (Scrolling) */
    .stApp { 
        padding-bottom: 180px !important; 
    }
    
    h1, h2, h3 { color: #d4af37 !important; border-bottom: 2px solid #0a192f; padding-bottom: 10px; }
    
    /* تصميم الشريط السفلي الحديدي */
    div[data-testid="stHorizontalBlock"]:last-of-type {
        position: fixed !important; bottom: 0 !important; left: 0 !important;
        width: 100vw !important; background-color: #ffffff !important;
        z-index: 9999999 !important; padding: 12px 0px !important;
        border-top: 2px solid #dfe6e9 !important;
        display: flex !important; flex-direction: row !important;
        justify-content: space-around !important;
        box-shadow: 0 -10px 25px rgba(0,0,0,0.1) !important;
    }
    div[data-testid="stHorizontalBlock"]:last-of-type button {
        background: transparent !important; color: #0a192f !important; border: none !important;
        box-shadow: none !important; height: 60px !important; width: 32vw !important;
    }
    div[data-testid="stHorizontalBlock"]:last-of-type button p { font-size: 14px !important; font-weight: 800 !important; }
    
    .card-box { background: white; padding: 20px; border-radius: 12px; border: 1px solid #dfe6e9; border-right: 6px solid #d4af37; margin-bottom: 20px; }
    .example-guide { color: #7f8c8d; font-size: 13px; font-style: italic; margin-bottom: 8px; display: block; border-right: 4px solid #d4af37; padding-right: 12px; background: #fafafa; padding: 8px; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 3. القاموس السيادي الماسي (25 مساراً - كاملة)
# ==========================================
methodology_db = {
    "الرقابة والامتثال (ISO 19011)": {
        "تقرير النزول الميداني": [("نطاق الفحص:", "مثال: جودة الهيكل الخرساني بمشروع برج المنصور."), ("الأدلة المادية:", "مثال: رصد تعشيش في الأعمدة رقم 4 و 5."), ("حالات عدم المطابقة:", "مثال: استخدام حديد 12ملم بدلاً من 14ملم.")],
        "تدقيق الامتثال الإداري": [("المعيار المرجعي:", "مثال: اللائحة الداخلية رقم 10 للمشتريات."), ("فجوة الصلاحيات:", "مثال: تجاوز المدير المالي لسقف الاعتماد."), ("نتائج المطابقة:", "مثال: وجود عجز بقيمة 2000 ريال.")],
        "الفحص والجرد الدوري": [("نطاق الجرد:", "مثال: جرد أصول مركز التدريب."), ("نسبة المطابقة:", "مثال: مطابقة بنسبة 98%."), ("حالة الأصول:", "مثال: 5 أجهزة خارجة عن الخدمة.")],
        "رقابة الجودة (QA/QC)": [("المعايير المرجعية:", "مثال: مواصفات ISO 9001."), ("نتائج الاختبارات:", "مثال: قوة ضغط العينة 25 ميجاباسكال."), ("نسبة المرفوضات:", "مثال: رفض 5% لعدم مطابقة اللون.")],
        "امتثال الصحة والسلامة": [("سجل الحوادث:", "مثال: إصابة طفيفة في القسم (أ)."), ("توفير معدات PPE:", "مثال: نقص في نظارات الحماية."), ("المخالفات:", "مثال: التدخين داخل الموقع.")]
    },
    "الأثر والتقييم (Kirkpatrick)": {
        "تقييم أثر التدريب": [("مستوى الرضا:", "مثال: تقييم المادة العلمية 4.5 من 5."), ("اكتساب المعرفة:", "مثال: تحسن الدرجات من 40% لـ 85%."), ("العائد المالي:", "مثال: توفير 2000$ شهرياً.")],
        "ختام وتقييم مشروع": [("المخرجات المحققة:", "مثال: حفر 5 آبار ارتوازية."), ("التحول النوعي:", "مثال: انخفاض أمراض المياه بنسبة 60%."), ("استدامة التدخل:", "مثال: تشكيل لجنة مجتمعية.")],
        "المسح القبلي (Baseline)": [("توصيف المشكلة:", "مثال: ارتفاع البطالة بين الخريجين."), ("إحصائيات الفجوة:", "مثال: 70% لا يجدون عملاً ملائماً."), ("تصميم الحل:", "مثال: دبلوم مكثف لـ 3 أشهر.")],
        "قياس العائد (SROI)": [("إجمالي الاستثمارات:", "مثال: 50,000$ من المانح."), ("نسبة العائد الاجتماعي:", "مثال: كل دولار حقق 3 دولار أثر."), ("مدة دوام الأثر:", "مثال: يستمر الأثر لـ 5 سنوات.")],
        "رضا المستفيدين (CSI)": [("مؤشر الرضا الكلي:", "مثال: نسبة الرضا العام 88%."), ("سهولة الوصول:", "مثال: 15% واجهوا صعوبة بالحجز."), ("خطة التحسين:", "مثال: تفعيل تطبيق موبايل.")]
    },
    "الاستراتيجية والمخاطر": {
        "دراسة جدوى ومخاطر": [("وصف الفرصة:", "مثال: إنشاء معمل خياطة."), ("احتمالية الحدوث:", "مثال: عالية (4/5)."), ("قرار الاستثمار:", "مثال: المشروع مجدٍ.")],
        "مراجعة استراتيجية": [("تحقيق الأهداف:", "مثال: إنجاز 70% من الخطة."), ("قوة المنافسة:", "مثال: ظهور منافس دولي جديد."), ("خارطة الطريق:", "مثال: التركيز على الاستدامة.")],
        "تحليل المنافسين": [("المنافسون:", "مثال: شركة (أ) وشركة (ب)."), ("الحصة السوقية:", "مثال: نستحوذ على 30%."), ("خطة الاستحواذ:", "مثال: تقديم ضمان 3 سنوات.")],
        "هندسة القيم (VE)": [("المهمة المستهدفة:", "مثال: بناء مدرسة مسبقة الصب."), ("حجم الوفر:", "مثال: توفير 5,000$ من الكلفة."), ("قرار الاعتماد:", "مثال: اعتماد البديل فوراً.")],
        "تقييم الجاهزية": [("هدف التحول:", "مثال: الأتمتة الكاملة للعمليات."), ("دعم القيادة:", "مثال: المدير العام ملتزم تماماً."), ("قرار الجاهزية:", "مثال: البدء بالمرحلة التجريبية.")]
    },
    "العمليات والإنتاجية": {
        "الإنجاز الدوري": [("المستهدفات:", "مثال: إنتاج 1000 حقيبة."), ("الهدر الزمني:", "مثال: تأخر في التوريد."), ("خطة التصحيح:", "مثال: العمل بنظام الإضافي.")],
        "تحليل الهدر (TIMWOODS)": [("نوع الهدر:", "مثال: هدر في الحركة والانتظار."), ("السبب الجذري:", "مثال: اشتراط توقيع المدير شخصياً."), ("النتيجة المتوقعة:", "مثال: تقليص الزمن لـ 4 ساعات.")],
        "أداء الموردين": [("اسم المورد:", "مثال: مكتب الأمل للتقنية."), ("مطابقة الجودة:", "مثال: التزام بكراسة الشروط."), ("القرار التعاقدي:", "مثال: تجديد العقد بامتياز.")],
        "إدارة الأزمات": [("توصيف الأزمة:", "مثال: اختراق سيبراني للسيرفر."), ("خطة التعافي:", "مثال: استعادة 95% من البيانات."), ("الدروس المستفادة:", "مثال: التحديث الأسبوعي للأمان.")],
        "مؤشرات الأداء (KPIs)": [("القسم المستهدف:", "مثال: قسم المبيعات."), ("Lead Time:", "مثال: 3 أيام للطلب."), ("القرارات:", "مثال: صرف مكافأة تميز.")]
    },
    "العلاقات والصورة المؤسسية": {
        "التغطية الإعلامية": [("الرسالة:", "مثال: إبراز الدور الإنساني."), ("الوصول (Reach):", "مثال: 100,000 مشاهدة."), ("توصية العلاقات:", "مثال: إطلاق حملة ممولة.")],
        "إدارة الفعاليات": [("نوع الفعالية:", "مثال: مؤتمر إطلاق الخطة."), ("إدارة الوقت:", "مثال: الالتزام بالجدول بدقة."), ("التقييم المالي:", "مثال: تمت ضمن الميزانية.")],
        "المسؤولية المجتمعية": [("المبادرة:", "مثال: ترميم 10 فصول ريفية."), ("الأثر الملموس:", "مثال: عودة 300 طالب."), ("ارتباط SDGs:", "مثال: يخدم الهدف 4.")],
        "الأزمات الإعلامية": [("طبيعة الأزمة:", "مثال: شائعة حول تأخر الرواتب."), ("قوة الرد:", "مثال: رد مدعوم بوثائق رسمية."), ("الوقاية:", "مثال: تحديث سياسة التواصل.")],
        "عائد الشراكات": [("الجهة الشريكة:", "مثال: جامعة صنعاء."), ("المنافع:", "مثال: وفر في كلفة التدريب."), ("توصية الاستمرار:", "تجديد لثلاث سنوات.")]
    }
}

# ==========================================
# 4. وظائف الدعم والأمان
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
    st.markdown('<div class="card-box" style="margin-top:50px; text-align:center;"><h1>🏛️ دخول المنصة السيادية</h1></div>', unsafe_allow_html=True)
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
    
    if st.session_state.get('current_report') != report_type:
        st.session_state.current_report, st.session_state.step = report_type, 1
    
    questions = methodology_db[pillar][report_type]
    
    if st.session_state.step == 1:
        st.subheader("📍 المرحلة 1")
        for i, (q, ex) in enumerate(questions):
            st.markdown(f"<span class='example-guide'>{ex}</span>", unsafe_allow_html=True)
            ans = st.text_area(f"**{i+1}. {q}**", value=get_draft(f"q_{report_type}_{i}"), key=f"k1_{i}")
            update_draft(f"q_{report_type}_{i}", ans)
        
        if st.button("اعتماد وتوليد الوثيقة السيادية 📄"):
            if db["users"][email]["balance"] <= 0: st.error("⚠️ رصيدك صفر.")
            else:
                try:
                    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
                    model = genai.GenerativeModel('gemini-1.5-flash')
                    with st.spinner("جاري الصياغة..."):
                        res = model.generate_content(f"صغ تقريراً لـ {report_type}")
                        st.session_state.report_preview = res.text
                        db["users"][email]["balance"] -= 1; save_db(db)
                        st.success("تم التوليد بنجاح!")
                except Exception as e: st.error(f"خطأ تقني: {e}")

    if st.session_state.report_preview:
        st.markdown("### 📄 المعاينة")
        st.info(st.session_state.report_preview)
        doc = Document()
        doc.add_heading(report_type, 0)
        doc.add_paragraph(st.session_state.report_preview)
        bio = io.BytesIO(); doc.save(bio)
        st.download_button("⬇️ تحميل Word", bio.getvalue(), file_name="report.docx")

def packages_page():
    st.title("💳 باقات الاشتراك")
    pkgs = [("البداية (3)", "1,000 ريال"), ("التمكين (6)", "1,500 ريال"), ("التنفيذية (12)", "2,500 ريال")]
    cols = st.columns(3)
    for i, (n, p) in enumerate(pkgs):
        with cols[i]:
            st.markdown(f'<div class="card-box" style="text-align:center;"><h3>{n}</h3><h2>{p}</h2><hr><a href="https://wa.me/967774575749?text=أريد باقة {n}" style="text-decoration:none;"><div style="background-color:#25D366; color:white; padding:10px; border-radius:8px; font-weight:bold;">📱 اطلب الكود</div></a></div>', unsafe_allow_html=True)
    
    code = st.text_input("أدخل الكود:")
    if st.button("تفعيل"):
        if code in db["codes"]:
            val = db["codes"].pop(code)
            db["users"][st.session_state.user_email]["balance"] += val
            save_db(db); st.success("تم التفعيل!"); time.sleep(1); st.rerun()

def admin_page():
    st.title("🛠️ الإدارة")
    if st.text_input("الرمز:", type="password") == "Mansour@2026":
        if st.button("توليد كود (باقة 3)"):
            c = f"MS-{uuid.uuid4().hex[:6].upper()}"
            db["codes"][c] = 3; save_db(db); st.info(c)

# ==========================================
# 6. شريط التنقل السفلي (الهندسة النهائية)
# ==========================================
def navigate(target):
    st.session_state.current_page = target
    st.rerun()

if not st.session_state.logged_in: login_page()
else:
    if st.session_state.current_page == "platform": platform_page()
    elif st.session_state.current_page == "packages": packages_page()
    elif st.session_state.current_page == "admin": admin_page()

    # الأزرار السفلية (إجبار الظهور في صف واحد)
    st.markdown("<br><br><br>", unsafe_allow_html=True)
    c_n1, c_n2, c_n3 = st.columns(3)
    with c_n1: st.button("🏠 المنصة", key="btn_p", on_click=navigate, args=("platform",))
    with c_n2: st.button("💳 الباقات", key="btn_k", on_click=navigate, args=("packages",))
    with c_n3: st.button("🛠️ الإدارة", key="btn_a", on_click=navigate, args=("admin",))
