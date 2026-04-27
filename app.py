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
# 1. تهيئة النظام وقاعدة البيانات (الأصول الثابتة)
# ==========================================
st.set_page_config(page_title="منصة المنصور السيادية", layout="wide", initial_sidebar_state="collapsed")

DB_FILE = "mansour_strategic_vault_2026.json"

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
# 2. الهوية البصرية الصارمة (Cairo & FIXED BOTTOM NAV)
# ==========================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap');
    #MainMenu, footer, header {visibility: hidden;}
    * { font-family: 'Cairo', sans-serif !important; direction: rtl !important; text-align: right !important; }
    html, body, .stApp { background-color: #f8f9fa !important; padding-bottom: 120px; }
    
    h1, h2, h3 { color: #d4af37 !important; border-bottom: 2px solid #0a192f; padding-bottom: 10px; }
    
    /* تنسيق الحقول */
    div[data-baseweb="input"] > div, div[data-baseweb="textarea"] > div, .stSelectbox > div {
        background-color: #ffffff !important; border: 1px solid #dfe6e9 !important; border-radius: 10px !important;
    }
    
    /* الأزرار السيادية */
    .stButton > button { 
        background-color: #0a192f !important; border-radius: 8px !important; 
        color: white !important; font-weight: 700 !important; width: 100% !important; padding: 12px !important;
    }
    .stButton > button:hover { background-color: #d4af37 !important; color: black !important; }

    /* اختراع الشريط السفلي - إجبار الظهور (WhatsApp Style) */
    div[data-testid="stHorizontalBlock"]:last-of-type {
        position: fixed !important;
        bottom: 0 !important;
        left: 0 !important;
        width: 100vw !important;
        background-color: #ffffff !important;
        z-index: 9999999 !important;
        padding: 10px 0px !important;
        border-top: 2px solid #dfe6e9 !important;
        flex-wrap: nowrap !important;
        display: flex !important;
        justify-content: space-around !important;
        box-shadow: 0 -10px 25px rgba(0,0,0,0.1) !important;
    }
    div[data-testid="stHorizontalBlock"]:last-of-type button {
        background: transparent !important;
        color: #0a192f !important;
        border: none !important;
        box-shadow: none !important;
        height: 55px !important;
        width: 33vw !important;
        margin: 0 !important;
    }
    div[data-testid="stHorizontalBlock"]:last-of-type button p {
        font-size: 15px !important;
        font-weight: 800 !important;
        color: #0a192f !important;
    }
    
    .card-box { background: white; padding: 20px; border-radius: 12px; border: 1px solid #dfe6e9; margin-bottom: 20px; border-right: 5px solid #d4af37; }
    .example-guide { color: #7f8c8d; font-size: 13px; font-style: italic; margin-bottom: 5px; display: block; border-right: 4px solid #d4af37; padding-right: 12px; background: #fdfdfd; padding: 5px; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 3. القاموس المنهجي الماسي (25 مساراً - لا حذف نهائياً)
# ==========================================
methodology_db = {
    "مسار الرقابة والامتثال (ISO 19011)": {
        "تقرير النزول الميداني الفني": [("نطاق الفحص الفني:", "مثال: جودة تنفيذ الهيكل الخرساني بمشروع برج المنصور."), ("الأدلة المادية:", "مثال: رصد تعشيش في الأعمدة رقم 4 و 5، وغياب عينات الفحص."), ("حالات عدم المطابقة:", "مثال: استخدام حديد بقطر 12 ملم بدلاً من 14 ملم المعتمد."), ("تحليل السبب الجذري:", "مثال: ضعف الرقابة الهندسية أثناء التوريد الصباحي."), ("تقييم مخاطر السلامة:", "مثال: غياب لوحات التحذير بجوار حفرة المصعد الرئيسية."), ("كفاءة الموارد المادية:", "مثال: هدر في الإسمنت بنسبة 15% نتيجة سوء التخزين."), ("جودة التوثيق والسجلات:", "مثال: سجل صب الخرسانة اليومي غير موقع من المهندس."), ("الاستجابة للملاحظات:", "مثال: لم يتم إغلاق ملاحظة العزل في التقرير السابق."), ("الإجراء التصحيحي:", "مثال: إيقاف الصب ومعالجة التعشيش فوراً بمواد كيميائية."), ("الإجراء الوقائي:", "مثال: توفير مراقب جودة مقيم وتحديث قائمة الموردين.")],
        "تقرير تدقيق الامتثال الإداري": [("المعيار المرجعي:", "مثال: اللائحة التنفيذية رقم 4 لسنة 2024 الخاصة بالمشتريات."), ("تحليل فجوة الصلاحيات:", "مثال: تجاوز المدير المالي لسقف الاعتماد بـ 20%."), ("سلامة الدورة المستندية:", "مثال: صرف فواتير دون وجود محاضر فحص واستلام."), ("كفاءة نظام الرقابة:", "مثال: ضعف الربط الآلي بين المستودع وبرنامج الحسابات."), ("الشفافية والمساءلة:", "مثال: غياب معايير المفاضلة الواضحة في اختيار الموردين."), ("التوافق مع الهيكل:", "مثال: قيام قسم HR بمهام إدارية تتبع المدير العام."), ("جودة نظام الأرشفة:", "مثال: حفظ وثائق العقود في مكاتب مفتوحة غير محمية."), ("مؤشرات الهدر الإداري:", "مثال: تكرار طلب البيانات الورقية المتوفرة إلكترونياً."), ("نتائج المطابقة المالية:", "مثال: وجود عجز بقيمة 5000 ريال في العهدة النقدية."), ("توصية لجنة التدقيق:", "مثال: إحالة الملف للتحقيق وتجميد الصلاحيات مؤقتاً.")],
        "تقرير الفحص والجرد الدوري": [("نطاق الجرد الحالي:", "مثال: جرد أصول مركز الاتصالات والوسائل التعليمية."), ("نسبة مطابقة الرصيد:", "مثال: مطابقة بنسبة 98% مع فائض في الحواسيب."), ("حالة الأصول الفنية:", "مثال: 5 أجهزة خارجة عن الخدمة بسبب التقادم."), ("تقييم بيئة التخزين:", "مثال: المستودع يفتقر لأنظمة التبريد والإطفاء."), ("دقة نظام الترميز:", "مثال: 40% من الأصول لا تحمل ملصقات الباركود."), ("تحليل التقادم والركود:", "مثال: وجود قطع غيار لمعدات تم الاستغناء عنها."), ("إجراءات الضبط المخزني:", "مثال: الصرف يتم دون أوامر صرف موقعة."), ("مخاطر فقدان الأصول:", "مثال: ضعف الحراسة الليلية للمستودع المفتوح."), ("توصية التسوية المحاسبية:", "مثال: شطب العهد التالفة وتحميل المقصر المسؤولية."), ("خطة تطوير المخازن:", "مثال: اعتماد نظام تتبع رقمي للأصول المتحركة.")],
        "تقرير رقابة الجودة (QA/QC)": [("المعايير المرجعية:", "مثال: مواصفات الآيزو 9001:2015 المعتمدة."), ("نتائج الاختبارات:", "مثال: قوة ضغط العينات 25 ميجاباسكال وهي مطابقة."), ("نسبة المرفوضات:", "مثال: رفض 5% من الإنتاج لعدم مطابقة اللون."), ("كفاءة أدوات القياس:", "مثال: أجهزة القياس تحتاج لمعايرة دورية (منتهية)."), ("أداء الموردين الميداني:", "مثال: المورد يلتزم بالجودة بنسبة 100% هذا الشهر."), ("شكاوى المستفيدين:", "مثال: تلقي 3 شكاوى عن بطء استجابة الدعم الفني."), ("تكلفة الجودة الرديئة:", "مثال: خسارة 500$ نتيجة إعادة تشغيل عينات تالفة."), ("التزام الكادر:", "مثال: 80% يطبقون أدلة التشغيل القياسية."), ("الدروس المستفادة:", "مثال: الفحص قبل الشحن يقلل تكاليف الإرجاع."), ("خطة التحسين (Kaizen):", "مثال: أتمتة نظام فحص الجودة لتقليل الخطأ البشري.")],
        "تقرير امتثال (HSE)": [("سجل الحوادث الميداني:", "مثال: إصابة طفيفة في اليد لأحد العمال."), ("توفير معدات السلامة:", "مثال: نقص في نظارات الحماية الخاصة باللحام."), ("أنظمة الإنذار:", "مثال: جرس الإنذار في القسم (ب) يحتاج استبدال."), ("تحليل المخاطر (JHA):", "مثال: خطر الانزلاق بسبب تسرب زيوت."), ("تراخيص العمل:", "مثال: رخصة تشغيل الرافعة منتهية الصلاحية."), ("إدارة النفايات الخطرة:", "مثال: تخزين زيوت مستعملة في حاويات مكشوفة."), ("خطة الإخلاء:", "مثال: مخارج الطوارئ مغلقة بعوائق خشبية."), ("التوعية والتدريب الصحي:", "مثال: ضعف مهارات الكادر في استخدام طفايات الحريق."), ("مخالفات السلامة:", "مثال: التدخين في مناطق المواد القابلة للاشتعال."), ("الإجراء الإلزامي:", "مثال: إيقاف العمل في القسم حتى معالجة تسرب الزيت.")]
    },
    "مسار الأثر والتقييم (Kirkpatrick)": {
        "تقرير تقييم أثر التدريب": [("مستوى الرضا (Reaction):", "مثال: تقييم المتدربين للمادة العلمية بلغ 4.7 من 5."), ("اكتساب المعرفة (Learning):", "مثال: ارتفاع الدرجات من 40% (قبل) إلى 90% (بعد)."), ("التغير السلوكي (Behavior):", "مثال: المتدربون بدأوا باستخدام برامج الأتمتة فعلياً."), ("العائد على النتائج (Results):", "مثال: تقليص زمن إصدار التقارير بنسبة 40%."), ("مؤشر الاستدامة المعرفية:", "مثال: بقاء المهارات لدى الكادر بعد مرور 6 أشهر."), ("ملاءمة البرنامج للاحتياج:", "مثال: التدريب لبى الفجوة في مهارات التفاوض الميداني."), ("دعم الإدارة للتطبيق:", "مثال: توفير أجهزة لوحية للمتدربين لممارسة العمل."), ("العائد المالي (ROI):", "مثال: توفير 2000$ شهرياً كانت تضيع في أخطاء الإدخال."), ("التأثير على السمعة:", "مثال: إشادة المانحين بجودة التقارير المرفوعة مؤخراً."), ("توصية تطوير البرامج:", "مثال: زيادة الجانب التطبيقي في النسخة القادمة.")],
        "تقرير ختام وتقييم مشروع": [("المخرجات المحققة:", "مثال: تشغيل 5 آبار مياه تعمل بالطاقة الشمسية."), ("التحول النوعي الملموس:", "مثال: انخفاض أمراض المياه بنسبة 60%."), ("مؤشر الوصول الفعلي:", "مثال: استفادة 1200 أسرة (زيادة 10% عن المخطط)."), ("كفاءة الإنفاق المالي:", "مثال: الصرف تم ضمن الموازنة المعتمدة بدقة."), ("استدامة التدخل:", "مثال: تشكيل لجنة مجتمعية للصيانة والتحصيل."), ("تحليل الأثر الجانبي:", "مثال: زيادة التحاق الفتيات بالتعليم بالمنطقة."), ("تقييم أداء الشركاء:", "مثال: المورد التزم بالمعايير الفنية والجدول الزمني."), ("الدروس المستفادة:", "مثال: أهمية إشراك المجتمع المحلي في مرحلة التخطيط."), ("قصة نجاح المشروع:", "مثال: حالة المواطن (س) الذي استعاد عافيته وأرضه."), ("التوصية النهائية:", "مثال: توسيع المشروع ليشمل المديريات المجاورة.")],
        "تقرير المسح القبلي (Baseline)": [("توصيف المشكلة الراهنة:", "مثال: ارتفاع نسبة البطالة بين الخريجين التقنيين."), ("إحصائيات الفجوة:", "مثال: 70% من خريجي التقنية لا يجدون عملاً ملائماً."), ("تحليل القدرات المحلية:", "مثال: توفر قاعات تدريب مجهزة لدى مكتب التعليم الفني."), ("أولويات التدخل العاجلة:", "مثال: التدريب على المهارات المطلوبة في سوق العمل."), ("خارطة أصحاب المصلحة:", "مثال: الغرفة التجارية، مكاتب التوظيف، الأكاديميات."), ("تقييم الحلول السابقة:", "مثال: التدخلات السابقة كانت نظرية وتفتقر للتطبيق."), ("التحديات المتوقعة:", "مثال: ضعف شبكة الإنترنت في المناطق الريفية."), ("توقعات المستفيدين:", "مثال: يتوقع الشباب الحصول على لابتوبات بعد الدورة."), ("تصميم مسار الحل:", "مثال: دبلوم مكثف لمدة 3 أشهر مع سنة متابعة."), ("مؤشرات النجاح:", "مثال: توظيف 50% من الخريجين خلال أول 6 أشهر.")],
        "تقرير قياس العائد (SROI)": [("إجمالي الاستثمارات:", "مثال: 100,000 دولار مدفوعة من المانح (X)."), ("خارطة التغير المحققة:", "مثال: تأهيل معاقين -> وظائف -> دمج مجتمعي."), ("ترجمة الأثر لقيمة:", "مثال: زيادة دخل الفئة بـ 300,000$ سنوياً."), ("تحليل الاستنزاف:", "مثال: 10% من التحسن كان سيحدث دون تدخلنا."), ("نسبة العائد الاجتماعي:", "مثال: كل دولار حقق 4.5 دولار كأثر اجتماعي."), ("تحليل الإسناد (Attribution):", "مثال: 90% من الأثر يعود لمؤسستنا حصراً."), ("مدة دوام الأثر:", "مثال: يتوقع استمرار الأثر المهني لمدة 10 سنوات."), ("المنافع البيئية والاقتصادية:", "مثال: تقليل الاعتماد على المعونات الحكومية."), ("صحة وموثوقية البيانات:", "مثال: تم الاعتماد على سجلات الضرائب والرواتب."), ("توصيات تعظيم الأثر مستقبلاً:", "مثال: ربط الخريجين بأسواق العمل الخارجية.")]
    }
}

# ==========================================
# 4. وظائف الدعم والأمان (حفظ المسودة وبصمة الجهاز)
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
    st.markdown('<div class="card-box" style="margin-top:50px; text-align:center;"><h1>🏛️ دخول المنصة السيادية</h1></div>', unsafe_allow_html=True)
    email = st.text_input("البريد الإلكتروني:", placeholder="yourname@domain.com")
    password = st.text_input("كلمة المرور:", type="password")
    if st.button("دخول آمن للمنصة"):
        if email and password:
            dev_id = get_device_id()
            if email in db["users"]:
                if db["users"][email]["password"] == password:
                    st.session_state.user_email, st.session_state.logged_in = email, True
                    st.session_state.current_page = "platform"
                    st.rerun()
                else: st.error("كلمة المرور غير صحيحة.")
            else:
                if dev_id in db["devices"]:
                    st.warning("⚠️ عذراً، هذا الجهاز سجل مسبقاً.")
                else:
                    db["users"][email] = {"password": password, "balance": 1, "device_id": dev_id, "drafts": {}}
                    db["devices"][dev_id] = email
                    save_db(db)
                    st.session_state.user_email, st.session_state.logged_in = email, True
                    st.session_state.current_page = "platform"
                    st.rerun()

def platform_page():
    email = st.session_state.user_email
    balance = db["users"][email]["balance"]
    st.title("المنصة السيادية")
    st.info(f"المستشار: **{email}** | الرصيد: **{balance} تقارير**")

    # بيانات الغلاف الإدارية
    st.markdown("### 🏛️ بيانات الغلاف (الإدارية)")
    org = st.text_input("الجهة المصدرة:", value=get_draft("org_name"))
    update_draft("org_name", org)
    proj = st.text_input("اسم المشروع:", value=get_draft("proj_name"))
    update_draft("proj_name", proj)
    author = st.text_input("إعداد (الاسم والمنصب):", value=get_draft("author_name"))
    update_draft("author_name", author)

    st.markdown("---")
    pillar = st.selectbox("حدد المسار الاستراتيجي:", list(methodology_db.keys()))
    report_type = st.selectbox("حدد التقرير المنهجي:", list(methodology_db[pillar].keys()))
    
    if st.session_state.current_report != report_type:
        st.session_state.current_report, st.session_state.step = report_type, 1
        st.session_state.report_preview = ""
    
    questions = methodology_db[pillar][report_type]
    
    # === المرحلة 1 ===
    if st.session_state.step == 1:
        st.markdown('<h4>📍 المرحلة 1: التشخيص</h4>', unsafe_allow_html=True)
        for i, (q, ex) in enumerate(questions[:3]):
            st.markdown(f"<span class='example-guide'>{ex}</span>", unsafe_allow_html=True)
            ans = st.text_area(f"**{i+1}. {q}**", value=get_draft(f"q_{report_type}_{i}"), key=f"k1_{i}")
            update_draft(f"q_{report_type}_{i}", ans)
        if st.button("التالي ⬅️"): st.session_state.step = 2; st.rerun()

    # === المرحلة 2 ===
    elif st.session_state.step == 2:
        st.markdown('<h4>📊 المرحلة 2: التحليل</h4>', unsafe_allow_html=True)
        for i, (q, ex) in enumerate(questions[3:7]):
            idx = i + 3
            st.markdown(f"<span class='example-guide'>{ex}</span>", unsafe_allow_html=True)
            ans = st.text_area(f"**{idx+1}. {q}**", value=get_draft(f"q_{report_type}_{idx}"), key=f"k2_{idx}")
            update_draft(f"q_{report_type}_{idx}", ans)
        if st.button("التالي ⬅️"): st.session_state.step = 3; st.rerun()
        if st.button("➡️ السابق"): st.session_state.step = 1; st.rerun()

    # === المرحلة 3 ===
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
            if balance <= 0: st.error("⚠️ رصيدك صفر. يرجى الشحن.")
            else:
                try:
                    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
                    # البحث عن المحرك المتاح لتجنب 404
                    available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
                    target_model = 'models/gemini-1.5-flash' if 'models/gemini-1.5-flash' in available_models else available_models[0]
                    model = genai.GenerativeModel(target_model)
                    
                    data_summary = "".join([f"- {q}: {get_draft(f'q_{report_type}_{i}')}\n" for i, (q, _) in enumerate(questions)])
                    prompt = f"أنت مستشار استراتيجي سيادي خبير. صغ تقريراً استشارياً لـ '{report_type}' لجهة '{org}' مشروع '{proj}'. البيانات: {data_summary}. التوصيات: {recs}. اللغة: رسمية، رصينة، نقاط مباشرة."
                    with st.spinner("المحرك الذكي يقوم بالصياغة..."):
                        try: res = model.generate_content(prompt)
                        except: time.sleep(3); res = model.generate_content(prompt)
                        st.session_state.report_preview = res.text
                        db["users"][email]["balance"] -= 1
                        save_db(db)
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
        st.download_button("⬇️ تحميل Word", bio.getvalue(), file_name=f"{proj}.docx")

def packages_page():
    st.title("💳 باقات الاشتراك")
    pkgs = [("بداية (3)", "1,000 ريال", "باقة البداية"), ("تمكين (6)", "1,500 ريال", "باقة التمكين"), ("تنفيذية (12)", "2,500 ريال", "الباقة التنفيذية")]
    cols = st.columns(3)
    for i, (name, price, msg) in enumerate(pkgs):
        with cols[i]:
            st.markdown(f'<div class="card-box" style="text-align:center;"><h3>{name}</h3><h2 style="color:#d4af37;">{price}</h2><hr><a href="https://wa.me/967774575749?text=أريد {msg}" style="text-decoration:none;"><div style="background-color:#25D366; color:white; padding:10px; border-radius:8px; font-weight:bold;">📱 اطلب الكود</div></a></div>', unsafe_allow_html=True)
    
    code = st.text_input("أدخل كود الشحن:")
    if st.button("تفعيل الكود"):
        if code in db["codes"]:
            val = db["codes"].pop(code)
            db["users"][st.session_state.user_email]["balance"] += val
            save_db(db)
            st.success(f"✅ تم تفعيل {val} تقارير!"); time.sleep(1); st.rerun()
        else: st.error("الكود خطأ.")

def admin_page():
    st.title("🛠️ الإدارة")
    pw = st.text_input("الرمز السري:", type="password")
    if pw == "Mansour@2026":
        num = st.selectbox("عدد التقارير:", [3, 6, 12])
        if st.button("توليد كود جديد"):
            c = f"MS-{uuid.uuid4().hex[:6].upper()}"
            db["codes"][c] = num
            save_db(db)
            st.info(f"كود التفعيل: **{c}**")
        st.write("المستخدمون:", db["users"])

# ==========================================
# 6. شريط التنقل السفلي (الحل الفولاذي)
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

    # رسم الأزرار في أسفل الشاشة (إجبار الظهور)
    st.markdown("<br><br><br>", unsafe_allow_html=True)
    nav_col1, nav_col2, nav_col3 = st.columns(3)
    with nav_col1:
        st.button("🏠 المنصة", key="nav_main", on_click=navigate, args=("platform",))
    with nav_col2:
        st.button("💳 الباقات", key="nav_pkg", on_click=navigate, args=("packages",))
    with nav_col3:
        st.button("🛠️ الإدارة", key="nav_adm", on_click=navigate, args=("admin",))
