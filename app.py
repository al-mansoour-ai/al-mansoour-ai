import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import datetime
import uuid
import json
import os

# ==========================================
# 1. تهيئة النظام وقاعدة البيانات الميدانية
# ==========================================
st.set_page_config(page_title="منصة المنصور الاستراتيجية", layout="wide", initial_sidebar_state="collapsed")

DB_FILE = "mansour_strategic_db.json"

def load_db():
    if not os.path.exists(DB_FILE):
        with open(DB_FILE, "w", encoding="utf-8") as f:
            json.dump({"users": {}, "codes": {"VIP2026": 100}}, f)
    with open(DB_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_db(data):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

db = load_db()

# إدارة الجلسة
if 'user_id' not in st.session_state: st.session_state.user_id = None
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'current_page' not in st.session_state: st.session_state.current_page = "login"
if 'step' not in st.session_state: st.session_state.step = 1
if 'report_preview' not in st.session_state: st.session_state.report_preview = ""
if 'current_report' not in st.session_state: st.session_state.current_report = ""

# ==========================================
# 2. الهندسة البصرية (Cairo Font & Professional UI)
# ==========================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap');
    #MainMenu, footer, header {visibility: hidden;}
    
    * { font-family: 'Cairo', sans-serif !important; direction: rtl !important; text-align: right !important; }
    html, body, .stApp { background-color: #f8f9fa !important; padding-bottom: 100px; }
    
    h1, h2, h3 { color: #d4af37 !important; border-bottom: 2px solid #0a192f; padding-bottom: 10px; }
    
    /* تنسيق الحقول الجوال */
    div[data-baseweb="input"] > div, div[data-baseweb="textarea"] > div, .stSelectbox > div {
        background-color: #ffffff !important; border: 1px solid #dfe6e9 !important; border-radius: 10px !important;
    }
    
    /* الأزرار السيادية */
    .stButton > button { 
        background-color: #0a192f !important; border: 1px solid #0a192f !important;
        border-radius: 10px !important; width: 100% !important; padding: 15px !important;
        color: white !important; font-weight: 700 !important;
    }
    .stButton > button:hover { background-color: #d4af37 !important; color: black !important; border-color: #d4af37 !important; }

    /* الشريط السفلي الأنيق (WhatsApp Style) */
    div[data-testid="stHorizontalBlock"]:last-of-type {
        position: fixed; bottom: 0; left: 0; width: 100vw;
        background-color: #ffffff !important; z-index: 99999;
        padding: 10px 0px; border-top: 1px solid #dfe6e9;
        flex-wrap: nowrap !important; justify-content: space-around !important;
    }
    div[data-testid="stHorizontalBlock"]:last-of-type button { background-color: transparent !important; border: none !important; color: #636e72 !important; height: 50px !important; }
    div[data-testid="stHorizontalBlock"]:last-of-type button:hover { color: #0a192f !important; font-weight: 700 !important; }
    
    .card-box { background: white; padding: 20px; border-radius: 12px; border: 1px solid #dfe6e9; margin-bottom: 20px; border-right: 5px solid #d4af37; }
    .whatsapp-btn { display: block; background-color: #25D366; color: white !important; text-align: center; padding: 12px; border-radius: 8px; text-decoration: none; font-weight: bold; margin-top: 10px; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 3. القاموس المنهجي (المسارات الـ 25 - الثوابت)
# ==========================================
methodology_db = {
    "مسار الرقابة والامتثال (ISO 19011)": {
        "تقرير النزول الميداني الفني": [("المرحلة التشغيلية:", "تجهيزات إنشائية"), ("نسبة الإنجاز:", "50% مخطط"), ("أسباب الانحراف:", "تأخر توريد"), ("حالات عدم المطابقة:", "أنابيب مخالفة"), ("كفاءة المعدات:", "معطلة"), ("السلامة المهنية:", "غياب لوحات"), ("التوثيق الميداني:", "غير محدث"), ("المخاطر المرصودة:", "انهيار تربة"), ("الاستجابة للملاحظات:", "لم تصحح"), ("القرار العاجل:", "إيقاف مؤقت")],
        "تقرير تدقيق الامتثال الإداري": [("المعيار المرجعي:", "لائحة المشتريات"), ("فجوة الصلاحيات:", "تداخل مهام"), ("نظام الأرشفة:", "فقدان عقود"), ("التوصيف الوظيفي:", "مهام غير واضحة"), ("شفافية الإجراءات:", "غياب المحاضر"), ("الجزاءات والمكافآت:", "بدون تقييم"), ("الاتصال الداخلي:", "واتساب"), ("الهيكل التنظيمي:", "قسم غير مفعل"), ("جودة التقارير:", "تفتقر للرقمية"), ("القرار المقترح:", "إعادة هيكلة")],
        "تقرير الفحص والجرد الدوري": [("المخزن:", "المركزي"), ("نسبة المطابقة:", "95%"), ("حالة الأصول:", "تالف"), ("إجراءات السلامة:", "منتهية"), ("الدورة المستندية:", "غير موقعة"), ("الظروف البيئية:", "تكييف معطل"), ("ترميز الأصول:", "بدون Barcode"), ("أسباب العجز:", "خطأ إدخال"), ("أداء الموظفين:", "يحتاج تدريب"), ("الإجراء المحاسبي:", "تسوية")],
        "تقرير رقابة الجودة (QA/QC)": [("المنتج:", "ورشة التدريب"), ("معيار الجودة:", "ISO 9001"), ("الانحرافات:", "ضعف المادة"), ("شكاوى المستفيدين:", "عدم وضوح"), ("أدوات القياس:", "غير دقيقة"), ("نسبة التوالف:", "10%"), ("أداء مسؤولي الجودة:", "تأخر الرفع"), ("تكلفة الجودة الرديئة:", "500$"), ("السبب الجذري:", "غياب المراجعة"), ("التوصية الوقائية:", "نموذج تدقيق")],
        "تقرير امتثال الصحة والسلامة": [("نطاق الفحص:", "ورش الصيانة"), ("سجل الحوادث:", "إصابة واحدة"), ("أدوات الحماية:", "نقص نظارات"), ("امتثال العاملين:", "20% مخالفين"), ("خطة الطوارئ:", "غير معلقة"), ("تخزين المواد:", "قرب حرارة"), ("تراخيص العمل:", "منتهية"), ("الوعي الصحي:", "ضعيف"), ("أسباب المخالفة:", "غياب المشرف"), ("القرار الإلزامي:", "إيقاف العمل")]
    },
    "مسار الأثر والتقييم (Kirkpatrick)": {
        "تقرير تقييم أثر التدريب": [("مستوى الرضا:", "9/10"), ("اكتساب المعرفة:", "تحسن 45%"), ("التغير السلوكي:", "تطبيق الأتمتة"), ("العائد على النتائج:", "انخفاض أخطاء"), ("استدامة المهارات:", "جلسة تنشيطية"), ("الوفر المالي:", "500$ شهرياً"), ("ملاءمة المخرجات:", "عالجت فجوة"), ("دعم الإدارة:", "توفير أجهزة"), ("التأثير على السمعة:", "زيادة الرضا"), ("توصية تطوير:", "زيادة العملي")],
        "تقرير ختام وتقييم مشروع": [("الهدف الاستراتيجي:", "مياه ريفية"), ("الوصول الفعلي:", "1200 أسرة"), ("التحول الملموس:", "توفير 3 ساعات"), ("مؤشرات النجاح:", "تدفق مستمر"), ("كفاءة الموازنة:", "15$ للمستفيد"), ("استدامة التدخل:", "لجنة مجتمعية"), ("الدروس المستفادة:", "التوريد المحلي"), ("أداء الشركاء:", "التزام كامل"), ("قصة نجاح:", "عودة الفتيات"), ("التوصية النهائية:", "تكرار النموذج")],
        "تقرير المسح القبلي": [("المشكلة الأساسية:", "بطالة تقنية"), ("إحصائيات:", "60% بلا عمل"), ("الفجوة:", "نقص التدريب"), ("القدرة والرغبة:", "رغبة عالية"), ("الموارد المتاحة:", "قاعات مجهزة"), ("التهديدات:", "انقطاع الإنترنت"), ("التوقعات:", "توظيف مباشر"), ("تقييم الغير:", "تدخل نظري"), ("أولويات التدخل:", "منهج حصري"), ("تصميم المشروع:", "دبلوم مهني")],
        "تقرير قياس العائد (SROI)": [("الاستثمار:", "20,000$"), ("التغير المحقق:", "تمكين 30 أسرة"), ("القيمة المالية:", "6000$ شهرياً"), ("الأثر البيئي:", "انخفاض الأمراض"), ("مدة بقاء الأثر:", "3 سنوات"), ("الأثر غير المقصود:", "تحسن اجتماعي"), ("نسبة العائد:", "1:4.5"), ("مقارنة البدائل:", "أفضل بـ 20%"), ("شفافية البيانات:", "فواتير فعلية"), ("توصيات التعظيم:", "ربط بالأسواق")],
        "تقرير رضا المستفيدين (CSI)": [("الخدمة:", "إغاثة نقدية"), ("مؤشر الرضا:", "82%"), ("سهولة الوصول:", "صعوبة تسجيل"), ("تعامل الفريق:", "احترافي"), ("سرعة التقديم:", "تأخير أسبوع"), ("جودة الخدمة:", "مطابقة"), ("الشكاوى:", "زيادة المبالغ"), ("وضوح الآلية:", "الرقم لا يرد"), ("معدل الولاء:", "90%"), ("خطة التحسين:", "رد آلي")]
    },
    "مسار الاستراتيجية والمخاطر": {
        "دراسة جدوى ومصفوفة مخاطر": [("الفرصة:", "مركز صيانة"), ("الاستثمار:", "60 ألف $"), ("الميزة السيادية:", "صيانة محلية"), ("نقطة التعادل:", "24 شهر"), ("أخطر 3 مخاطر:", "تذبذب العملة"), ("خطة التحوط:", "حساب دولار"), ("القوى العاملة:", "5 موظفين"), ("الأثر الاقتصادي:", "تقليل تعطل"), ("البيئة التشريعية:", "إعفاء جمركي"), ("القرار النهائي:", "مجدٍ")],
        "تقرير المراجعة الاستراتيجية": [("الهدف الاستراتيجي:", "توسع طاقة"), ("تحقيق المؤشرات:", "40% فقط"), ("أسباب الفجوة:", "ميزانية طوارئ"), ("تغيرات خارجية:", "منافس دولي"), ("تقييم المحفظة:", "هدر 60%"), ("نقاط القوة:", "سمعة عالية"), ("نقاط الضعف:", "ضعف الكادر"), ("الشراكات:", "تعثر إداري"), ("إعادة التوجيه:", "تحول للتنمية"), ("التحديث المطلوب:", "تمكين اقتصادي")],
        "تقرير تحليل المنافسين": [("وصف السوق:", "سوق الأدوية"), ("أبرز المنافسين:", "شركة س/ص/ع"), ("الحصة السوقية:", "45%"), ("نقاط القوة لهم:", "أسطول نقل"), ("نقاط الضعف لهم:", "خدمة بطيئة"), ("الفجوة المتاحة:", "أدوية الأرياف"), ("استراتيجية السعر:", "حرق سعري"), ("عوائق الدخول:", "تصاريح"), ("ولاء العملاء:", "للسعر الأقل"), ("خطة الاختراق:", "تسهيلات دفع")],
        "تقرير هندسة القيم (VE)": [("المشروع المستهدف:", "مركز صحي"), ("التكلفة المعتمدة:", "120,000$"), ("وظائف أساسية:", "أساسات"), ("وظائف ثانوية:", "رخام"), ("البديل المقترح:", "طلاء بكتيري"), ("الوفر المتوقع:", "18,000$"), ("تأثير الجودة:", "لا يوجد"), ("تأثير الوقت:", "تسريع أسبوعين"), ("موقف المورد:", "مرحب"), ("القرار الهندسي:", "اعتماد البديل")],
        "تقرير تقييم الجاهزية": [("هدف التقييم:", "تحول رقمي"), ("جاهزية الهاردوير:", "سيرفرات قديمة"), ("جاهزية السوفتوير:", "رخص غير مفعلة"), ("كفاءة الكادر:", "يحتاج تدريب"), ("الاستقرار المالي:", "ميزانية مرصودة"), ("دعم الإدارة:", "قرار ملزم"), ("مقاومة التغيير:", "تخوف موظفين"), ("جاهزية اللوائح:", "قديمة جداً"), ("تحليل المخاطر:", "توقف مؤقت"), ("قرار الجاهزية:", "تأجيل شهرين")]
    },
    "مسار العمليات والإنتاجية": {
        "تقرير الإنجاز الدوري": [("المستهدفات:", "حفر 500 متر"), ("الإنجاز الفعلي:", "450 متر"), ("أسباب الهدر:", "نقص الديزل"), ("كفاءة الموازنة:", "صرف 100%"), ("جودة المخرجات:", "مطابقة"), ("البيروقراطية:", "تأخر مستخلص"), ("أداء الفريق:", "إنتاجية جيدة"), ("تحديث التوريد:", "وصول إسمنت"), ("الاحتياجات القادمة:", "مضخات"), ("خطة التصحيح:", "عمل إضافي")],
        "تقرير تحليل الهدر": [("نوع الهدر:", "انتظار/حركة"), ("موقع الهدر:", "قسم المشتريات"), ("حجم الخسارة:", "3 أيام"), ("السبب الجذري:", "توقيع مركزي"), ("تأثير العميل:", "تأخر تسليم"), ("إجراء كايزن:", "تفويض صلاحيات"), ("تكلفة التحسين:", "لا يوجد"), ("مقاومة التغيير:", "قلق مالي"), ("مدة التطبيق:", "أسبوع"), ("النتيجة المتوقعة:", "ساعتان")],
        "تقرير أداء الموردين": [("اسم المورد:", "شركة الأفق"), ("الالتزام بالوقت:", "تأخير 4 أيام"), ("مطابقة الجودة:", "100%"), ("مرونة السعر:", "أغلى بـ 10%"), ("الاستجابة للطوارئ:", "بطيئة"), ("خدمات ما بعد البيع:", "ممتازة"), ("الوضع المالي:", "مستقر"), ("سجل المخالفات:", "لا يوجد"), ("مخاطر الاعتماد:", "احتكار"), ("القرار التعاقدي:", "تجديد بشرط")],
        "تقرير إدارة الأزمات": [("وصف الأزمة:", "اختراق سيبراني"), ("التوقيت:", "الخميس 10 مساءً"), ("الاستجابة الفورية:", "فصل السيرفر"), ("حجم الخسائر:", "بيانات يوم"), ("خطة استمرارية:", "عمل ورقي"), ("كفاءة الفريق:", "سريعة جداً"), ("التواصل:", "إبلاغ العملاء"), ("تحليل القصور:", "كلمة مرور ضعيفة"), ("عملية التعافي:", "استرجاع نسخة"), ("الوقاية مستقبلاً:", "تفعيل 2FA")],
        "تقرير كفاءة الأداء (KPIs)": [("القسم المستهدف:", "خدمة العملاء"), ("زمن الاستجابة:", "22 دقيقة"), ("إغلاق التذاكر:", "98%"), ("تكلفة العميل:", "4$"), ("سبب التدني:", "نقص كادر"), ("منجز استثنائي:", "تحديث الأسئلة"), ("أداء الموظفين:", "قيادي متميز"), ("مقارنة بالعام:", "تحسن 30%"), ("مستهدف القادم:", "10 دقائق"), ("قرار تحفيزي:", "مكافأة تميز")]
    },
    "مسار العلاقات وصورة المؤسسة": {
        "تقرير التغطية الإعلامية": [("الرسالة المستهدفة:", "إغاثة طارئة"), ("الوصول الكلي:", "مليون ظهور"), ("نبرة الجمهور:", "85% إيجابي"), ("المؤثرون:", "5 ناشطين"), ("جودة المحتوى:", "تصوير جوي"), ("وصول الرسائل:", "تمويل ذاتي"), ("كفاءة الإنفاق:", "نشر عضوي"), ("فجوات التواصل:", "غياب X"), ("تهديدات السمعة:", "إشاعة تأخير"), ("توصية العلاقات:", "حملة توضيح")],
        "تقرير إدارة الفعاليات": [("نوع الفعالية:", "مؤتمر سنوي"), ("لوجستيات القاعة:", "ممتازة"), ("دقة البروتوكول:", "خطأ بسيط"), ("إدارة الوقت:", "تأخير 45 دقيقة"), ("جودة المتحدثين:", "متمكنين"), ("التغطية الداخلية:", "بث مباشر"), ("حجم الحضور:", "200 شخص"), ("إدارة الأزمات:", "تغيير مايك"), ("التقييم المالي:", "ضمن الموازنة"), ("الدروس المستفادة:", "التأكيد المسبق")],
        "تقرير المسؤولية المجتمعية": [("اسم المبادرة:", "شتاء دافئ"), ("الفئة المستهدفة:", "500 يتيم"), ("حجم المساهمة:", "15,000$"), ("الأثر المباشر:", "حماية صحية"), ("تطوع الموظفين:", "20 موظفاً"), ("الشركاء المحليون:", "الوحدة التنفيذية"), ("الصورة الإنسانية:", "تغطية واسعة"), ("ردود الأفعال:", "إشادة رسمية"), ("ارتباط SDGs:", "القضاء على الفقر"), ("المبادرة القادمة:", "بطولة رياضية")],
        "تقرير الأزمات الإعلامية": [("طبيعة الأزمة:", "إشاعة انتهاء"), ("منصة الانتشار:", "واتساب"), ("سرعة الاستجابة:", "3 ساعات"), ("الإجراء الفوري:", "بيان رسمي"), ("الرد الجوهري:", "فحص مخبري"), ("دور الجهات:", "تصريح الصحة"), ("نبرة الجمهور:", "تحول للدعم"), ("أضرار السمعة:", "محدودة"), ("أداء المتحدث:", "واثق ومقنع"), ("الخطة الوقائية:", "طباعة بارزة")],
        "تقرير عائد الشراكات": [("الجهة الشريكة:", "جامعة تعز"), ("الأهداف:", "تدريب 50 خريج"), ("المنافع:", "وفر 3000$"), ("التزاماتنا:", "تدريس مساق"), ("مستوى التنسيق:", "سلس"), ("عوائق التنفيذ:", "تأخر كشوفات"), ("تأثير السمعة:", "جهة راعية"), ("التقييم المالي:", "مربحة جداً"), ("سرية المعلومات:", "التزام تام"), ("التوصية:", "تجديد 3 سنوات")]
    }
}

# ==========================================
# 4. منطق الحفظ التلقائي والحزم البرمجي
# ==========================================
def update_draft(key, value):
    uid = st.session_state.user_id
    if uid:
        if "draft" not in db["users"][uid]: db["users"][uid]["draft"] = {}
        db["users"][uid]["draft"][key] = value
        save_db(db)

def get_draft(key, default=""):
    uid = st.session_state.user_id
    if uid and uid in db["users"]:
        return db["users"][uid].get("draft", {}).get(key, default)
    return default

# ==========================================
# 5. صفحات النظام (Login, Platform, Packages, Admin)
# ==========================================
def login_page():
    st.markdown('<div class="card-box" style="margin-top:50px;">', unsafe_allow_html=True)
    st.title("🏛️ دخول المنصة السيادية")
    uid = st.text_input("رقم الجوال (للدخول واستعادة التقارير):", placeholder="مثال: 774575749")
    if st.button("دخول آمن للمنصة"):
        if uid:
            if uid not in db["users"]:
                db["users"][uid] = {"balance": 1, "draft": {}} # تجربة مجانية
                save_db(db)
            st.session_state.user_id = uid
            st.session_state.logged_in = True
            st.session_state.current_page = "platform"
            st.rerun()

def platform_page():
    uid = st.session_state.user_id
    balance = db["users"][uid]["balance"]
    st.title("المنصور الاستراتيجية")
    st.info(f"المستشار: **{uid}** | الرصيد: **{balance} تقارير**")

    # بيانات الغلاف
    st.markdown("### 🏛️ أولاً: بيانات الغلاف (الإدارية)")
    org = st.text_input("الجهة المصدرة للوثيقة:", value=get_draft("org_name"))
    update_draft("org_name", org)
    loc = st.text_input("النطاق الجغرافي:", value=get_draft("loc_name"))
    update_draft("loc_name", loc)
    proj = st.text_input("اسم المشروع / المهمة:", value=get_draft("proj_name"))
    update_draft("proj_name", proj)
    author = st.text_input("إعداد (الاسم والمنصب):", value=get_draft("author_name"))
    update_draft("author_name", author)

    st.markdown("---")
    pillar = st.selectbox("1. المسار الاستراتيجي:", list(methodology_db.keys()))
    report_type = st.selectbox("2. التقرير التخصصي:", list(methodology_db[pillar].keys()))
    
    if st.session_state.current_report != report_type:
        st.session_state.current_report = report_type
        st.session_state.step = 1
        st.session_state.report_preview = ""
    
    questions = methodology_db[pillar][report_type]
    st.markdown(f"--- المرحلة الحالية: **{st.session_state.step} من 3** ---")
    
    if st.session_state.step == 1:
        st.subheader("📍 المرحلة 1: التشخيص")
        for i, (q, h) in enumerate(questions[:3]):
            ans = st.text_area(q, value=get_draft(f"q_{report_type}_{i}"), placeholder=h, key=f"k1_{report_type}_{i}")
            update_draft(f"q_{report_type}_{i}", ans)
        if st.button("التالي: التحليل ⬅️", key="n1"): st.session_state.step = 2; st.rerun()

    elif st.session_state.step == 2:
        st.subheader("📊 المرحلة 2: التحليل")
        for i, (q, h) in enumerate(questions[3:7]):
            idx = i + 3
            ans = st.text_area(q, value=get_draft(f"q_{report_type}_{idx}"), placeholder=h, key=f"k2_{report_type}_{idx}")
            update_draft(f"q_{report_type}_{idx}", ans)
        if st.button("التالي: القرار ⬅️", key="n2"): st.session_state.step = 3; st.rerun()
        if st.button("➡️ السابق", key="p2"): st.session_state.step = 1; st.rerun()

    elif st.session_state.step == 3:
        st.subheader("🎯 المرحلة 3: القرار")
        for i, (q, h) in enumerate(questions[7:]):
            idx = i + 7
            ans = st.text_area(q, value=get_draft(f"q_{report_type}_{idx}"), placeholder=h, key=f"k3_{report_type}_{idx}")
            update_draft(f"q_{report_type}_{idx}", ans)
        recs = st.text_area("توصيات ختامية:", value=get_draft(f"recs_{report_type}"), key="recs_area")
        update_draft(f"recs_{report_type}", recs)
        
        if st.button("اعتماد وتوليد الوثيقة 📄", key="gen_btn"):
            if balance <= 0: st.error("⚠️ الرصيد صفر.")
            else:
                try:
                    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
                    # المحرك فلاش 1.5 هو الأسرع والأكثر استقراراً حالياً
                    model = genai.GenerativeModel('gemini-1.5-flash')
                    data_feed = "".join([f"- {q}: {get_draft(f'q_{report_type}_{i}')}\n" for i, (q, _) in enumerate(questions)])
                    
                    prompt = f"أنت مستشار استراتيجي سيادي. صغ تقرير '{report_type}' لجهة '{org}' مشروع '{proj}'. البيانات: {data_feed}. التوصيات: {recs}. اللغة: رسمية، رصينة، نقاط مباشرة."
                    with st.spinner("المحرك الذكي يحلل البيانات..."):
                        res = model.generate_content(prompt)
                        st.session_state.report_preview = res.text
                        db["users"][uid]["balance"] -= 1
                        save_db(db)
                        st.success("تم التوليد بنجاح.")
                except Exception as e:
                    st.error(f"خطأ تقني: {e}")
        if st.button("➡️ السابق", key="p3"): st.session_state.step = 2; st.rerun()

    if st.session_state.report_preview:
        st.markdown("### 📄 معاينة الوثيقة")
        st.info(st.session_state.report_preview)
        doc = Document()
        doc.add_heading(report_type, 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
        for line in st.session_state.report_preview.split('\n'):
            if line.strip(): doc.add_paragraph(line.strip()).alignment = WD_ALIGN_PARAGRAPH.RIGHT
        bio = io.BytesIO()
        doc.save(bio)
        st.download_button("⬇️ تحميل Word", bio.getvalue(), file_name=f"{proj}.docx")

def packages_page():
    st.title("💳 باقات الاشتراك")
    plans = [("بداية (3)", "1,000"), ("تمكين (6)", "1,500"), ("تنفيذية (12)", "2,500")]
    cols = st.columns(3)
    for i, (name, price) in enumerate(plans):
        with cols[i]:
            st.markdown(f'<div class="card-box"><h4>باقة {name}</h4><p>السعر: {price} ريال</p></div>', unsafe_allow_html=True)
            st.markdown(f'<a href="https://wa.me/967774575749?text=أريد باقة {name}" class="whatsapp-btn">طلب كود</a>', unsafe_allow_html=True)
    
    code = st.text_input("أدخل كود الشحن:")
    if st.button("تفعيل الرصيد"):
        if code in db["codes"]:
            val = db["codes"].pop(code)
            db["users"][st.session_state.user_id]["balance"] += val
            save_db(db)
            st.success("تم التفعيل!")
        else: st.error("الكود خطأ.")

def admin_page():
    st.title("🛠️ الإدارة")
    pw = st.text_input("الرمز السري:", type="password")
    if pw == "Mansour@2026":
        num = st.number_input("الرصيد:", 1, 100)
        if st.button("توليد كود"):
            c = f"MS-{uuid.uuid4().hex[:4].upper()}"
            db["codes"][c] = num
            save_db(db)
            st.code(c)

# ==========================================
# 6. التنقل السفلي والتحكم
# ==========================================
def nav(p): st.session_state.current_page = p; st.rerun()

if not st.session_state.logged_in:
    login_page()
else:
    if st.session_state.current_page == "platform": platform_page()
    elif st.session_state.current_page == "packages": packages_page()
    elif st.session_state.current_page == "admin": admin_page()

    nav1, nav2, nav3 = st.columns(3)
    with nav1:
        if st.button("🏠 المنصة", key="n_p"): nav("platform")
    with nav2:
        if st.button("💳 الباقات", key="n_k"): nav("packages")
    with nav3:
        if st.button("🛠️ الإدارة", key="n_a"): nav("admin")
