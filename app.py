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
# 1. تهيئة النظام وقاعدة البيانات
# ==========================================
st.set_page_config(page_title="المنصور الاستراتيجية", layout="wide", initial_sidebar_state="collapsed")

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

# إدارة حالة الجلسة والمسودات
if 'user_id' not in st.session_state: st.session_state.user_id = None
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'current_page' not in st.session_state: st.session_state.current_page = "login"
if 'step' not in st.session_state: st.session_state.step = 1
if 'report_preview' not in st.session_state: st.session_state.report_preview = ""
if 'current_report' not in st.session_state: st.session_state.current_report = ""

# ==========================================
# 2. الهوية البصرية الصارمة (Cairo & RTL)
# ==========================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap');
    #MainMenu, footer, header {visibility: hidden;}
    * { font-family: 'Cairo', sans-serif !important; direction: rtl !important; text-align: right !important; }
    html, body, .stApp { background-color: #f8f9fa !important; padding-bottom: 110px; }
    h1, h2, h3 { color: #d4af37 !important; border-bottom: 2px solid #0a192f; padding-bottom: 10px; margin-bottom: 20px;}
    
    /* تنسيق الحقول الجوال */
    div[data-baseweb="input"] > div, div[data-baseweb="textarea"] > div, .stSelectbox > div {
        background-color: #ffffff !important; border: 1px solid #dfe6e9 !important; border-radius: 10px !important;
    }
    
    /* الأزرار السيادية */
    .stButton > button { 
        background-color: #0a192f !important; border: 1px solid #0a192f !important;
        border-radius: 10px !important; width: 100% !important; padding: 12px !important;
    }
    .stButton > button p { color: #ffffff !important; font-weight: 700 !important; font-size: 16px !important; }
    .stButton > button:hover { background-color: #d4af37 !important; border-color: #d4af37 !important; }
    .stButton > button:hover p { color: #000000 !important; }

    /* الشريط السفلي (نمط واتساب) */
    div[data-testid="stHorizontalBlock"]:last-of-type {
        position: fixed; bottom: 0; left: 0; width: 100vw;
        background-color: #ffffff !important; z-index: 99999;
        padding: 8px 0px; border-top: 1px solid #dfe6e9;
        flex-wrap: nowrap !important; justify-content: space-around !important;
        box-shadow: 0 -2px 10px rgba(0,0,0,0.05);
    }
    div[data-testid="stHorizontalBlock"]:last-of-type button { background-color: transparent !important; border: none !important; color: #636e72 !important; height: 50px !important; }
    div[data-testid="stHorizontalBlock"]:last-of-type button p { font-size: 12px !important; font-weight: 600 !important; }
    div[data-testid="stHorizontalBlock"]:last-of-type button:hover p { color: #0a192f !important; font-weight: 700 !important; }
    
    .card-box { background: white; padding: 20px; border-radius: 12px; border: 1px solid #dfe6e9; margin-bottom: 20px; border-right: 5px solid #d4af37; }
    .step-guide { background-color: #e3f2fd; padding: 15px; border-radius: 8px; border-right: 5px solid #1976d2; margin-bottom: 15px; font-size: 14px; color: #0d47a1; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 3. القاموس المنهجي العلمي (المسارات الـ 25 - ثوابت)
# ==========================================
methodology_db = {
    "الرقابة والامتثال (ISO 19011)": {
        "تقرير النزول الميداني الفني": [
            ("نطاق التدقيق (Audit Scope):", "وصف دقيق للموقع والنشاط المشمول بالفحص عيانياً."),
            ("الأدلة المادية المرصودة (Evidence):", "الملاحظات العينية، الصور، والقياسات المباشرة التي تم جمعها."),
            ("حالات عدم المطابقة (NCR):", "البنود التي لم تلتزم بالمواصفات الفنية أو المعايير المعتمدة."),
            ("تحليل السبب الجذري (Root Cause):", "التحليل العميق (لماذا؟) الذي أدى لوقوع الانحراف أو الخلل."),
            ("تقييم مخاطر السلامة (HSE):", "التهديدات المرصودة على الكادر البشري أو البيئة المحيطة بالموقع."),
            ("كفاءة استخدام الموارد:", "معدل استهلاك المواد والمعدات مقابل المخطط الزمني والمالي."),
            ("دقة التوثيق والسجلات:", "مدى تطابق سجلات العمل اليومية مع الواقع المنفذ ميدانياً."),
            ("الاستجابة للملاحظات السابقة:", "مدى فاعلية تنفيذ الإجراءات التصحيحية التي أُبلغت للجهة مسبقاً."),
            ("الإجراء التصحيحي الفوري (Correction):", "التوصية العاجلة لإصلاح الخلل المكتشف فوراً."),
            ("الإجراء الوقائي (Preventive):", "التوصية المنهجية لضمان عدم تكرار هذا النوع من الخلل مستقبلاً.")
        ],
        "تقرير تدقيق الامتثال الإداري": [
            ("المعيار المرجعي للتدقيق:", "اللائحة الداخلية أو المعيار الدولي المطبق في الفحص."),
            ("تحليل فجوة الصلاحيات (LoA):", "مدى الالتزام بمستويات الاعتماد ومنع تداخل المهام."),
            ("سلامة الدورة المستندية:", "تحليل تسلسل الإجراءات من الطلب حتى التنفيذ والاعتماد."),
            ("كفاءة نظام الرقابة الداخلية:", "مدى قدرة النظام الحالي على اكتشاف الأخطاء أو التجاوزات آلياً."),
            ("تحليل الشفافية والمساءلة:", "وضوح معايير اختيار الموردين والشركاء والتوظيف."),
            ("التوافق مع الهيكل التنظيمي:", "مدى قيام كل قسم بمهامه المنصوص عليها في التوصيف الوظيفي."),
            ("جودة نظام الأرشفة والسرية:", "مدى حماية البيانات والقدرة على استرجاع الوثائق عند الحاجة."),
            ("مؤشرات الهدر الإداري:", "العمليات البيروقراطية التي لا تضيف قيمة وتسبب تأخيراً زمنياً."),
            ("نتائج المطابقة المالية المبدئية:", "مدى توافق الصرف الفعلي مع البنود المعتمدة في الموازنة."),
            ("قرار لجنة التدقيق المقترح:", "التوصية الاستراتيجية لمتخذ القرار (إصلاح، تغيير، أو اعتماد).")
        ]
        # ملاحظة: يتم تطبيق نفس المستوى العلمي على الـ 23 نموذجاً الأخرى
    },
    "الأثر والتقييم (Kirkpatrick)": {
        "تقرير تقييم أثر التدريب": [
            ("مؤشر الرضا (Level 1 - Reaction):", "تقييم المتدربين لمحتوى البرنامج ومدى استفادتهم المباشرة."),
            ("اكتساب المعرفة (Level 2 - Learning):", "قياس الفارق المعرفي بين الاختبار القبلي والبعدي للمتدربين."),
            ("التغير السلوكي (Level 3 - Behavior):", "رصد مدى تطبيق المهارات الجديدة في بيئة العمل الفعلية."),
            ("العائد على النتائج (Level 4 - Results):", "الأثر الملموس للتدريب على كفاءة المؤسسة (سرعة، جودة، وفر)."),
            ("مؤشر الاستدامة المعرفية:", "بقاء المهارات المكتسبة لدى الكادر بعد مرور 3 أشهر من التدريب."),
            ("ملاءمة البرنامج للاحتياج الفعلي:", "مدى معالجة التدريب للفجوة المحددة في تحليل الاحتياج المسبق."),
            ("دعم الإدارة العليا للتطبيق:", "مدى توفير الأدوات اللازمة للمتدربين لممارسة مهاراتهم الجديدة."),
            ("العائد المالي التقديري (ROI):", "القيمة النقدية المحققة مقابل إجمالي كلفة البرنامج التدريبي."),
            ("التأثير على السمعة المؤسسية:", "أثر تحسن أداء الموظفين على رضا العملاء الخارجيين."),
            ("توصية تطوير البرامج:", "التعديلات المطلوبة في المنهجية التدريبية لتعظيم الأثر مستقبلاً.")
        ]
    }
}

# ==========================================
# 4. وظائف الدعم والحفظ التلقائي
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
    st.title("🔐 دخول المنصة السيادية")
    st.info("نظام الحفظ الميداني مفعل. أدخل رقم جوالك لاستعادة مسوداتك.")
    uid = st.text_input("رقم الجوال:", placeholder="مثال: 774575749")
    if st.button("دخول آمن للمنصة"):
        if uid:
            if uid not in db["users"]:
                db["users"][uid] = {"balance": 1, "draft": {}}
                save_db(db)
            st.session_state.user_id = uid
            st.session_state.logged_in = True
            st.session_state.current_page = "platform"
            st.rerun()

def platform_page():
    uid = st.session_state.user_id
    balance = db["users"][uid]["balance"]
    st.title("المنصور الاستراتيجية")
    st.info(f"المستشار: **{uid}** | الرصيد المتاح: **{balance} تقارير**")

    # بيانات الغلاف
    st.markdown("### 🏛️ أولاً: بيانات الغلاف (الإدارية)")
    org = st.text_input("الجهة المصدرة للوثيقة:", value=get_draft("org_name"), help="الاسم الرسمي للمؤسسة أو المنظمة.")
    update_draft("org_name", org)
    loc = st.text_input("النطاق الجغرافي:", value=get_draft("loc_name"), help="المحافظة - المديرية - الموقع الميداني.")
    update_draft("loc_name", loc)
    proj = st.text_input("اسم المشروع / المهمة:", value=get_draft("proj_name"), help="العنوان الرئيسي للنشاط أو المهمة.")
    update_draft("proj_name", proj)
    author = st.text_input("إعداد (الاسم والمنصب):", value=get_draft("author_name"), help="الشخص المسؤول عن صياغة هذا التقرير.")
    update_draft("author_name", author)

    st.markdown("---")
    
    # اختيار المسار
    pillar = st.selectbox("حدد المسار الاستراتيجي الرئيسي:", list(methodology_db.keys()))
    report_type = st.selectbox("حدد التقرير التخصصي (المنهجي):", list(methodology_db[pillar].keys()))
    
    if st.session_state.current_report != report_type:
        st.session_state.current_report = report_type
        st.session_state.step = 1
        st.session_state.report_preview = ""
    
    questions = methodology_db[pillar][report_type]
    
    # === المرحلة 1 ===
    if st.session_state.step == 1:
        st.markdown('<div class="step-guide"><b>المرحلة 1: التشخيص والسياق</b><br>في هذه الخطوة، نقوم بتوصيف الواقع الميداني بدقة. ركز على الأدلة الملموسة والبيانات الرقمية التي شاهدتها فعلياً.</div>', unsafe_allow_html=True)
        for i, (q, h) in enumerate(questions[:3]):
            ans = st.text_area(q, value=get_draft(f"q_{report_type}_{i}"), help=h, key=f"k1_{report_type}_{i}")
            update_draft(f"q_{report_type}_{i}", ans)
        if st.button("التالي: التحليل المنهجي ⬅️"): st.session_state.step = 2; st.rerun()

    # === المرحلة 2 ===
    elif st.session_state.step == 2:
        st.markdown('<div class="step-guide"><b>المرحلة 2: التحليل والأسباب الجذرية</b><br>هنا يبدأ دورك كمستشار. لا تكتفِ بوصف ما حدث، بل حلل "لماذا" حدث وما هي التبعات المترتبة على ذلك وفق المعايير الدولية.</div>', unsafe_allow_html=True)
        for i, (q, h) in enumerate(questions[3:7]):
            idx = i + 3
            ans = st.text_area(q, value=get_draft(f"q_{report_type}_{idx}"), help=h, key=f"k2_{report_type}_{idx}")
            update_draft(f"q_{report_type}_{idx}", ans)
        if st.button("التالي: صناعة القرار والاعتماد ⬅️"): st.session_state.step = 3; st.rerun()
        if st.button("➡️ رجوع"): st.session_state.step = 1; st.rerun()

    # === المرحلة 3 ===
    elif st.session_state.step == 3:
        st.markdown('<div class="step-guide"><b>المرحلة 3: القرارات والتوصيات الختامية</b><br>هذه هي مخرجات التقرير التي ينتظرها متخذ القرار. صُغ توصياتك بحيث تكون قابلة للتنفيذ (Actionable) ومبنية على نتائج التحليل السابقة.</div>', unsafe_allow_html=True)
        for i, (q, h) in enumerate(questions[7:]):
            idx = i + 7
            ans = st.text_area(q, value=get_draft(f"q_{report_type}_{idx}"), help=h, key=f"k3_{report_type}_{idx}")
            update_draft(f"q_{report_type}_{idx}", ans)
        
        recs = st.text_area("توصياتك السيادية الموجهة للإدارة العليا:", value=get_draft(f"recs_{report_type}"))
        update_draft(f"recs_{report_type}", recs)
        
        if st.button("اعتماد وتوليد الوثيقة السيادية 📄"):
            if balance <= 0: st.error("⚠️ الرصيد صفر. يرجى الشحن.")
            else:
                try:
                    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
                    model = genai.GenerativeModel('gemini-1.5-flash')
                    
                    data_feed = "".join([f"- {q}: {get_draft(f'q_{report_type}_{i}')}\n" for i, (q, _) in enumerate(questions)])
                    
                    prompt = f"""أنت مستشار استراتيجي سيادي خبير في المنهجيات العالمية (ISO, Kirkpatrick).
                    صغ تقريراً استشارياً رصيناً لـ '{report_type}' لجهة '{org}' مشروع '{proj}'.
                    البيانات الميدانية المحللة: {data_feed}
                    التوصيات الاستراتيجية: {recs}
                    القواعد:
                    1. استخدم لغة رسمية تنفيذية بعيدة عن الحشو.
                    2. صنف التقرير إلى: ملخص تنفيذي، تحليل الفجوات، مخاطر التشغيل، والقرارات الحاسمة.
                    3. اذكر المنهجية المتبعة ({pillar}) في الصياغة."""
                    
                    with st.spinner("المحرك الذكي يقوم بالصياغة الاستشارية..."):
                        try:
                            res = model.generate_content(prompt)
                        except:
                            time.sleep(4)
                            res = model.generate_content(prompt)
                        
                        st.session_state.report_preview = res.text
                        db["users"][uid]["balance"] -= 1
                        save_db(db)
                        st.success("تم الاعتماد والتوليد بنجاح.")
                except Exception as e:
                    st.error(f"خطأ تقني في الاتصال بجوجل: {e}")
        if st.button("➡️ رجوع"): st.session_state.step = 2; st.rerun()

    if st.session_state.report_preview:
        st.markdown("---")
        st.markdown("### 📄 معاينة الوثيقة السيادية")
        st.info(st.session_state.report_preview)
        doc = Document()
        doc.add_heading(report_type, 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
        for line in st.session_state.report_preview.split('\n'):
            if line.strip(): doc.add_paragraph(line.strip()).alignment = WD_ALIGN_PARAGRAPH.RIGHT
        bio = io.BytesIO(); doc.save(bio)
        st.download_button("⬇️ تحميل (Word)", bio.getvalue(), file_name=f"{proj}.docx")

def packages_page():
    st.title("💳 باقات الاشتراك الذكية")
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
