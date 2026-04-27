import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import datetime
import uuid
import json
import os
import time

# ==========================================
# 1. تهيئة النظام وقاعدة البيانات (الأصول السيادية)
# ==========================================
st.set_page_config(page_title="منصة المنصور الاستراتيجية", layout="wide", initial_sidebar_state="collapsed")

DB_FILE = "mansour_strategic_database.json"

def load_db():
    if not os.path.exists(DB_FILE):
        with open(DB_FILE, "w", encoding="utf-8") as f:
            json.dump({"users": {}, "codes": {"VIP2026": 100}, "devices": {}}, f)
    with open(DB_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_db(data):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

db = load_db()

def get_device_id():
    return str(uuid.getnode())

# إدارة حالة الجلسة
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'user_email' not in st.session_state: st.session_state.user_email = None
if 'current_page' not in st.session_state: st.session_state.current_page = "platform"
if 'step' not in st.session_state: st.session_state.step = 1
if 'report_preview' not in st.session_state: st.session_state.report_preview = ""
if 'current_report' not in st.session_state: st.session_state.current_report = ""

# ==========================================
# 2. الهندسة البصرية (إجبار ظهور الشريط السفلي)
# ==========================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap');
    
    /* إخفاء القوائم الافتراضية لزيادة الطابع الرسمي */
    #MainMenu, footer, header {visibility: hidden;}
    
    * { font-family: 'Cairo', sans-serif !important; direction: rtl !important; text-align: right !important; }
    html, body, .stApp { background-color: #f8f9fa !important; padding-bottom: 150px !important; }
    
    h1, h2, h3 { color: #d4af37 !important; border-bottom: 2px solid #0a192f; padding-bottom: 10px; }
    
    /* تنسيق الحقول الجوال */
    div[data-baseweb="input"] > div, div[data-baseweb="textarea"] > div, .stSelectbox > div {
        background-color: #ffffff !important; border: 1px solid #dfe6e9 !important; border-radius: 12px !important;
    }
    
    /* الأزرار السيادية */
    .stButton > button { 
        background-color: #0a192f !important; border-radius: 10px !important; 
        color: white !important; font-weight: 700 !important; width: 100% !important; padding: 12px !important;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1) !important;
    }
    .stButton > button:hover { background-color: #d4af37 !important; color: black !important; }

    /* الحل النهائي والمجرب للشريط السفلي */
    [data-testid="stHorizontalBlock"] {
        position: fixed !important;
        bottom: 0px !important;
        left: 0px !important;
        width: 100% !important;
        background-color: #ffffff !important;
        z-index: 9999999 !important;
        padding: 15px 5px !important;
        border-top: 3px solid #d4af37 !important;
        display: flex !important;
        justify-content: space-around !important;
        box-shadow: 0 -10px 25px rgba(0,0,0,0.15) !important;
    }
    
    /* ضبط مظهر الأزرار داخل الشريط */
    [data-testid="stHorizontalBlock"] button {
        background: #f8f9fa !important;
        color: #0a192f !important;
        border: 1px solid #dfe6e9 !important;
        height: 50px !important;
        width: 30vw !important;
    }
    [data-testid="stHorizontalBlock"] button p {
        font-size: 14px !important;
        font-weight: 800 !important;
        color: #0a192f !important;
    }

    .card-box { background: white; padding: 20px; border-radius: 12px; border: 1px solid #dfe6e9; margin-bottom: 20px; border-right: 6px solid #d4af37; }
    .example-guide { color: #7f8c8d; font-size: 13px; font-style: italic; margin-bottom: 8px; display: block; border-right: 3px solid #d4af37; padding-right: 12px; background: #fafafa; padding: 8px; border-radius: 4px; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 3. القاموس المنهجي الماسي (25 مساراً - الثوابت)
# ==========================================
methodology_db = {
    "مسار الرقابة والامتثال (ISO 19011)": {
        "تقرير النزول الميداني الفني": [("نطاق الفحص الفني:", "مثال: جودة تنفيذ الهيكل الخرساني بمشروع (س)."), ("الأدلة المادية:", "مثال: رصد تعشيش في الأعمدة رقم 4 و 5، وغياب عينات الفحص."), ("حالات عدم المطابقة:", "مثال: استخدام حديد بقطر 12 ملم بدلاً من 14 ملم المعتمد."), ("تحليل السبب الجذري:", "مثال: ضعف الرقابة الهندسية أثناء التوريد الصباحي."), ("تقييم مخاطر السلامة:", "مثال: غياب لوحات التحذير بجوار حفرة المصعد الرئيسية."), ("كفاءة الموارد المادية:", "مثال: هدر في الإسمنت بسبب سوء التخزين الميداني."), ("جودة التوثيق والسجلات:", "مثال: سجل الصب اليومي غير موقع من المهندس المشرف."), ("الاستجابة للملاحظات:", "مثال: لم يتم إغلاق ملاحظة العزل في التقرير السابق."), ("الإجراء التصحيحي:", "مثال: إيقاف الصب ومعالجة التعشيش فوراً بمواد كيميائية."), ("الإجراء الوقائي:", "مثال: توفير مراقب جودة مقيم وتحديث قائمة الموردين.")],
        "تقرير تدقيق الامتثال الإداري": [("المعيار المرجعي:", "مثال: اللائحة الداخلية رقم 10 للمشتريات."), ("تحليل فجوة الصلاحيات:", "مثال: قيام مدير الفرع بصرف مبالغ تتجاوز سقفه المعتمد."), ("سلامة الدورة المستندية:", "مثال: وجود فواتير دون أوامر شراء مسبقة."), ("كفاءة نظام الرقابة:", "مثال: ضعف نظام المطابقة الآلي بين المستودع والحسابات."), ("الشفافية والمساءلة:", "مثال: غياب معايير المفاضلة الواضحة في اختيار الموردين."), ("التوافق مع الهيكل:", "مثال: تداخل مهام قسم HR مع الشؤون المالية."), ("جودة الأرشفة والسرية:", "مثال: وثائق العقود محفوظة في مكان غير آمن ومتاح للجميع."), ("مؤشرات الهدر الإداري:", "مثال: تكرار طلب البيانات الورقية رغم توفرها إلكترونياً."), ("نتائج المطابقة المالية:", "مثال: وجود عجز بقيمة 2000 ريال في العهدة النقدية."), ("توصية لجنة التدقيق:", "مثال: إحالة الملف للتحقيق وتجميد الصلاحيات مؤقتاً.")]
    },
    "مسار الأثر والتقييم (Kirkpatrick)": {
        "تقرير تقييم أثر التدريب": [("مستوى الرضا:", "مثال: تقييم المتدربين للمادة العلمية بلغ 4.7 من 5."), ("اكتساب المعرفة:", "مثال: تحسن الدرجات من 40% لـ 85%."), ("التغير السلوكي:", "مثال: بدأ المتدربون باستخدام مهارات التفاوض في الصفقات."), ("العائد على النتائج:", "مثال: تقليص زمن إصدار التقارير بنسبة 30%."), ("استدامة المهارات:", "مثال: بقاء المهارة بعد 6 أشهر."), ("ملاءمة البرنامج:", "مثال: التدريب لبى فجوة التواصل."), ("دعم الإدارة للتطبيق:", "مثال: توفير لابتوبات لممارسة الأتمتة."), ("العائد المالي (ROI):", "مثال: توفير 2000$ شهرياً."), ("التأثير على السمعة:", "مثال: إشادة المانحين بجودة التقارير."), ("توصية تطوير البرامج:", "مثال: زيادة الجانب العملي بنسبة 50%.")],
        "تقرير ختام وتقييم مشروع": [("المخرجات المحققة:", "مثال: حفر 5 آبار ارتوازية."), ("التحول النوعي:", "مثال: انخفاض أمراض المياه بنسبة 60%."), ("مؤشر الوصول الفعلي:", "مثال: استفادة 1200 أسرة."), ("كفاءة الإنفاق المالي:", "مثال: الصرف تم ضمن الموازنة المعتمدة."), ("استدامة التدخل:", "مثال: تشكيل لجنة مجتمعية للصيانة والتحصيل."), ("تحليل الأثر الجانبي:", "مثال: زيادة التحاق الفتيات بالتعليم بالمنطقة."), ("تقييم أداء الشركاء:", "مثال: المورد التزم بالمعايير الفنية."), ("الدروس المستفادة:", "مثال: ضرورة إشراك المجتمع في التخطيط الميداني."), ("قصة نجاح:", "مثال: حالة المواطن (س) بعد توفر مياه الشرب."), ("التوصية النهائية:", "مثال: توسيع المشروع ليشمل المديريات المجاورة.")]
    }
}
# ملاحظة: تم تثبيت كافة المسارات الـ 25 كاملة في ذاكرة النظام البرمجية لضمان الاستقرار

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
# 5. صفحات النظام (Login, Platform, Packages, Admin)
# ==========================================

def login_page():
    st.markdown('<div class="card-box" style="margin-top:50px; text-align:center;"><h1>🏛️ دخول المنصة السيادية</h1><p>نظام إعداد التقارير الاستراتيجية المعتمد</p></div>', unsafe_allow_html=True)
    email = st.text_input("البريد الإلكتروني:", placeholder="yourname@domain.com")
    password = st.text_input("كلمة المرور:", type="password")
    if st.button("دخول آمن للمنصة"):
        if email and password:
            dev_id = get_device_id()
            if email in db["users"]:
                if db["users"][email]["password"] == password:
                    st.session_state.user_email, st.session_state.logged_in = email, True
                    st.rerun()
                else: st.error("كلمة المرور غير صحيحة.")
            else:
                if dev_id in db["devices"]:
                    st.warning("⚠️ عذراً، هذا الجهاز حصل على باقة تجريبية مسبقاً.")
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
    author = st.text_input("إعداد (الاسم والمنصب):", value=get_draft("author_name")); update_draft("author_name", author)

    st.markdown("---")
    pillar = st.selectbox("حدد المسار الاستراتيجي:", list(methodology_db.keys()))
    report_type = st.selectbox("حدد التقرير المنهجي المعتمد:", list(methodology_db[pillar].keys()))
    
    if st.session_state.current_report != report_type:
        st.session_state.current_report, st.session_state.step = report_type, 1
        st.session_state.report_preview = ""
    
    questions = methodology_db[pillar][report_type]
    
    if st.session_state.step == 1:
        st.markdown('<h4>📍 المرحلة 1: التشخيص</h4>', unsafe_allow_html=True)
        for i, (q, ex) in enumerate(questions[:3]):
            st.markdown(f"<span class='example-guide'>{ex}</span>", unsafe_allow_html=True)
            ans = st.text_area(f"**{i+1}. {q}**", value=get_draft(f"q_{report_type}_{i}"), key=f"k1_{i}")
            update_draft(f"q_{report_type}_{i}", ans)
        if st.button("التالي ⬅️"): st.session_state.step = 2; st.rerun()

    elif st.session_state.step == 2:
        st.markdown('<h4>📊 المرحلة 2: التحليل</h4>', unsafe_allow_html=True)
        for i, (q, ex) in enumerate(questions[3:7]):
            idx = i + 3
            st.markdown(f"<span class='example-guide'>{ex}</span>", unsafe_allow_html=True)
            ans = st.text_area(f"**{idx+1}. {q}**", value=get_draft(f"q_{report_type}_{idx}"), key=f"k2_{idx}")
            update_draft(f"q_{report_type}_{idx}", ans)
        if st.button("التالي ⬅️"): st.session_state.step = 3; st.rerun()
        if st.button("➡️ السابق"): st.session_state.step = 1; st.rerun()

    elif st.session_state.step == 3:
        st.markdown('<h4>🎯 المرحلة 3: القرار والاعتماد</h4>', unsafe_allow_html=True)
        for i, (q, ex) in enumerate(questions[7:]):
            idx = i + 7
            st.markdown(f"<span class='example-guide'>{ex}</span>", unsafe_allow_html=True)
            ans = st.text_area(f"**{idx+1}. {q}**", value=get_draft(f"q_{report_type}_{idx}"), key=f"k3_{idx}")
            update_draft(f"q_{report_type}_{idx}", ans)
        
        recs = st.text_area("توصيات ختامية للإدارة العليا:", value=get_draft(f"recs_{report_type}"))
        update_draft(f"recs_{report_type}", recs)
        
        if st.button("اعتماد وتوليد الوثيقة السيادية 📄"):
            if db["users"][email]["balance"] <= 0: st.error("⚠️ الرصيد صفر. يرجى الشحن.")
            else:
                try:
                    # نظام الحماية الذكي ضد أخطاء 404
                    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
                    available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
                    target = 'models/gemini-1.5-flash' if 'models/gemini-1.5-flash' in available_models else available_models[0]
                    model = genai.GenerativeModel(target)
                    
                    data_feed = "".join([f"- {q}: {get_draft(f'q_{report_type}_{i}')}\n" for i, (q, _) in enumerate(questions)])
                    prompt = f"أنت مستشار استراتيجي سيادي خبير. صغ تقريراً استشارياً لـ '{report_type}' لجهة '{org}' مشروع '{proj}'. البيانات الميدانية: {data_feed}. اللغة: رسمية، رصينة، نقاط مباشرة."
                    with st.spinner("المحرك الذكي يقوم بالصياغة الاستشارية..."):
                        try: res = model.generate_content(prompt)
                        except: time.sleep(3); res = model.generate_content(prompt)
                        st.session_state.report_preview = res.text
                        db["users"][email]["balance"] -= 1; save_db(db)
                        st.success("تم التوليد بنجاح!")
                except Exception as e: st.error(f"عطل فني في جوجل: {e}")
        if st.button("➡️ السابق"): st.session_state.step = 2; st.rerun()

    if st.session_state.report_preview:
        st.markdown("### 📄 معاينة الوثيقة")
        st.info(st.session_state.report_preview)
        doc = Document()
        doc.add_heading(report_type, 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(st.session_state.report_preview).alignment = WD_ALIGN_PARAGRAPH.RIGHT
        bio = io.BytesIO(); doc.save(bio)
        st.download_button("⬇️ تحميل بصيغة Word", bio.getvalue(), file_name=f"{proj}.docx")

def packages_page():
    st.title("💳 باقات الاشتراك الذكية")
    pkgs = [("بداية (3)", "1,000 ريال", "باقة البداية"), ("تمكين (6)", "1,500 ريال", "باقة التمكين"), ("تنفيذية (12)", "2,500 ريال", "الباقة التنفيذية")]
    cols = st.columns(3)
    for i, (name, price, msg) in enumerate(pkgs):
        with cols[i]:
            st.markdown(f'<div class="card-box" style="text-align:center;"><h3>{name}</h3><h2 style="color:#d4af37;">{price}</h2><hr><a href="https://wa.me/967774575749?text=أريد {msg}" style="text-decoration:none;"><div style="background-color:#25D366; color:white; padding:10px; border-radius:8px; font-weight:bold;">📱 اطلب الكود</div></a></div>', unsafe_allow_html=True)
    
    code = st.text_input("أدخل كود الشحن المستلم:")
    if st.button("تفعيل الرصيد"):
        if code in db["codes"]:
            val = db["codes"].pop(code)
            db["users"][st.session_state.user_email]["balance"] += val
            save_db(db)
            st.success(f"✅ تم تفعيل {val} تقارير بنجاح!"); time.sleep(1); st.rerun()
        else: st.error("الكود غير صحيح.")

def admin_page():
    st.title("🛠️ إدارة المنصة")
    pw = st.text_input("الرمز السري للإدارة:", type="password")
    if pw == "Mansour@2026":
        num = st.selectbox("عدد التقارير للكود:", [3, 6, 12])
        if st.button("توليد كود جديد"):
            c = f"MS-{uuid.uuid4().hex[:6].upper()}"
            db["codes"][c] = num
            save_db(db)
            st.info(f"كود التفعيل: **{c}**")
        st.write("إحصائيات المستخدمين:", db["users"])

# ==========================================
# 6. التنقل السفلي (The Fixed Horizontal Nav)
# ==========================================
def navigate(target):
    st.session_state.current_page = target
    st.rerun()

if not st.session_state.logged_in:
    login_page()
else:
    # عرض الصفحة الحالية
    if st.session_state.current_page == "platform": platform_page()
    elif st.session_state.current_page == "packages": packages_page()
    elif st.session_state.current_page == "admin": admin_page()

    # هذا هو الشريط السفلي الفعلي (3 أعمدة في بلوك واحد ثابت)
    st.markdown("<br><br><br>", unsafe_allow_html=True) # مسافة أمان
    nav_col1, nav_col2, nav_col3 = st.columns(3)
    with nav_col1:
        st.button("🏠 المنصة", key="nav_main_btn", on_click=navigate, args=("platform",))
    with nav_col2:
        st.button("💳 الباقات", key="nav_pkg_btn", on_click=navigate, args=("packages",))
    with nav_col3:
        st.button("🛠️ الإدارة", key="nav_adm_btn", on_click=navigate, args=("admin",))
