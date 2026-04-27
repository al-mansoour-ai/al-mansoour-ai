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
# 1. إعدادات المنصة وقاعدة البيانات (الحفظ الدائم)
# ==========================================
st.set_page_config(page_title="المنصور الاستراتيجية", layout="wide", initial_sidebar_state="collapsed")

DB_FILE = "mansour_database.json"

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
# 2. الهوية البصرية الصارمة (Cairo & RTL)
# ==========================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap');
    #MainMenu, footer, header {visibility: hidden;}
    * { font-family: 'Cairo', sans-serif !important; direction: rtl !important; text-align: right !important; }
    html, body, .stApp { background-color: #f8f9fa !important; padding-bottom: 110px; }
    h1, h2, h3 { color: #d4af37 !important; border-bottom: 2px solid #0a192f; padding-bottom: 10px; margin-bottom: 20px;}
    div[data-baseweb="input"] > div, div[data-baseweb="textarea"] > div, .stSelectbox > div {
        background-color: #ffffff !important; border: 1px solid #dfe6e9 !important; border-radius: 10px !important;
    }
    .stButton > button { 
        background-color: #0a192f !important; border: 2px solid #0a192f !important;
        border-radius: 10px !important; width: 100% !important; padding: 12px !important;
    }
    .stButton > button p { color: #ffffff !important; font-weight: 700 !important; font-size: 16px !important; }
    .stButton > button:hover { background-color: #d4af37 !important; border-color: #d4af37 !important; }
    .stButton > button:hover p { color: #000000 !important; }
    
    div[data-testid="stHorizontalBlock"]:last-of-type {
        position: fixed; bottom: 0; left: 0; width: 100vw;
        background-color: #ffffff !important; z-index: 99999;
        padding: 10px 0px; border-top: 1px solid #dfe6e9;
        flex-wrap: nowrap !important; justify-content: space-around !important;
        box-shadow: 0 -2px 10px rgba(0,0,0,0.05);
    }
    div[data-testid="stHorizontalBlock"]:last-of-type button { background-color: transparent !important; border: none !important; color: #636e72 !important; height: 50px !important; }
    div[data-testid="stHorizontalBlock"]:last-of-type button p { font-size: 12px !important; font-weight: 600 !important; }
    div[data-testid="stHorizontalBlock"]:last-of-type button:hover p { color: #0a192f !important; font-weight: 700 !important; }
    .card-box { background: white; padding: 20px; border-radius: 12px; border: 1px solid #dfe6e9; margin-bottom: 20px; border-right: 5px solid #d4af37; }
    .whatsapp-btn { display: block; background-color: #25D366; color: white !important; text-align: center; padding: 10px; border-radius: 8px; text-decoration: none; font-weight: bold; margin-top: 10px; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 3. القاموس المنهجي (المسارات الـ 25 كاملة)
# ==========================================
methodology_db = {
    "مسار الرقابة والامتثال (ISO 19011)": {
        "تقرير النزول الميداني الفني": [("نطاق التدقيق:", "وصف الموقع"), ("الأدلة المادية:", "ملاحظات عينية"), ("حالات عدم المطابقة:", "بنود مخالفة"), ("تحليل السبب الجذري:", "لماذا حدث الخلل؟"), ("تقييم مخاطر السلامة:", "أخطار محتملة"), ("كفاءة استخدام الموارد:", "معدل الاستهلاك"), ("جودة التوثيق:", "دقة السجلات"), ("الاستجابة للملاحظات:", "إغلاق التوصيات"), ("التوصية التصحيحية:", "إجراء فوري"), ("الإجراء الوقائي:", "لمنع التكرار")],
        "تقرير تدقيق الامتثال الإداري": [("المعيار المرجعي:", "اللائحة المتبعة"), ("تحليل فجوة الصلاحيات:", "تداخل المهام"), ("سلامة الدورة المستندية:", "تسلسل الاعتمادات"), ("كفاءة الرقابة الداخلية:", "نظام الضبط"), ("تحليل الشفافية:", "معايير التعاقد"), ("التوافق مع الهيكل:", "تفعيل المهام"), ("كفاءة نظام الأرشفة:", "حماية البيانات"), ("مؤشرات الهدر الإداري:", "بيروقراطية"), ("نتائج المطابقة المالية:", "الفروقات"), ("قرار لجنة التدقيق:", "توصية نهائية")],
        "تقرير الفحص والجرد الدوري": [("نطاق الجرد:", "المخازن المشمولة"), ("نسبة مطابقة الرصيد:", "المطابقة الفعلية"), ("حالة الأصول الفنية:", "جديد، صالح، تالف"), ("تقييم بيئة التخزين:", "الأمان والحرارة"), ("دقة نظام الترميز:", "تتبع والباركود"), ("تحليل التقادم:", "أصول فاقدة للقيمة"), ("إجراءات الضبط:", "أذونات الصرف"), ("مخاطر فقدان الأصول:", "تلف أو سرقة"), ("توصية التسوية:", "معالجة الفروقات"), ("خطة تطوير المستودعات:", "مقترح التحسين")],
        "تقرير رقابة الجودة (QA/QC)": [("معايير الجودة:", "المواصفات القياسية"), ("نتائج الاختبارات:", "القيم الرقمية"), ("نسبة المرفوضات:", "معدل العيوب"), ("كفاءة أدوات القياس:", "دقة الأجهزة"), ("تقييم أداء المزودين:", "جودة التوريد"), ("شكاوى المستفيدين:", "تغذية راجعة"), ("تحليل تكلفة الأخطاء:", "خسارة الجودة"), ("مدى التزام الكادر:", "الوعي بالمعايير"), ("الدروس المستفادة:", "خبرات مستقاة"), ("خطة التحسين (Kaizen):", "إجراءات الرفع")],
        "تقرير امتثال الصحة والسلامة": [("سجل الحوادث:", "إصابات أو وفيات"), ("توفير معدات (PPE):", "نسبة الالتزام"), ("صلاحية أنظمة الإنذار:", "جاهزية الطوارئ"), ("تحليل المخاطر (JHA):", "توصيف الأخطار"), ("تراخيص العمل:", "مدى القانونية"), ("التخلص من النفايات:", "الامتثال البيئي"), ("خطة الإخلاء:", "وضوح المسارات"), ("التوعية والتدريب:", "إسعافات أولية"), ("مخالفات السلامة:", "وصف دقيق"), ("الإجراء الإلزامي:", "قرار السلامة")]
    },
    "مسار الأثر والتقييم (Kirkpatrick)": {
        "تقرير تقييم أثر التدريب": [("مستوى الرضا:", "ملاءمة البيئة"), ("اكتساب المعرفة:", "الفرق المعرفي"), ("التغير السلوكي:", "تطبيق المهارات"), ("العائد على النتائج:", "أثر الإنتاجية"), ("مؤشر الاستدامة:", "بقاء المهارة"), ("تحليل ملاءمة الاحتياج:", "سد الفجوة"), ("عوائق التطبيق:", "أسباب عدم التنفيذ"), ("العائد مالي (ROI):", "القيمة مقابل الكلفة"), ("تحليل السمعة:", "صورة الجهة"), ("توصية تطوير البرامج:", "تحديث المناهج")],
        "تقرير ختام وتقييم مشروع": [("المخرجات المحققة:", "النتائج الكمية"), ("التحول النوعي:", "تغير حياة المستفيد"), ("مؤشر الوصول الفعلي:", "المخطط مقابل المحقق"), ("كفاءة الإنفاق المالي:", "التكلفة الفعلية"), ("مدى استدامة التدخل:", "استمرار الأثر"), ("تحليل الأثر الجانبي:", "نتائج غير مقصودة"), ("تقييم أداء الشركاء:", "التزام المنفذين"), ("الدروس المستفادة:", "خبرات للتوظيف"), ("قصة نجاح:", "حالة واقعية"), ("التوصية النهائية:", "تكرار النموذج")],
        "تقرير المسح القبلي": [("توصيف المشكلة:", "الواقع الراهن"), ("إحصائيات الفجوة:", "بيانات الاحتياج"), ("تحليل القدرات المحلية:", "الموارد المتاحة"), ("أولويات التدخل:", "احتياجات عاجلة"), ("خارطة أصحاب المصلحة:", "المتأثرون"), ("تقييم الحلول السابقة:", "لماذا فشل الغير؟"), ("التحديات المتوقعة:", "عوائق مستقبلية"), ("توقعات المستفيدين:", "طموح المجتمع"), ("تصميم مسار الحل:", "المنهجية المقترحة"), ("مؤشرات النجاح:", "كيف سنقيس الأثر؟")],
        "تقرير قياس العائد (SROI)": [("إجمالي الاستثمارات:", "مدخلات مالية"), ("خارطة التغير:", "كيف حدث التغير؟"), ("ترجمة الأثر لقيمة:", "وكلاء ماليون"), ("تحليل الاستنزاف:", "بدون تدخلنا"), ("نسبة العائد الاجتماعي:", "القيمة مقابل الدولار"), ("تحليل الإسناد:", "دورنا حصراً"), ("مدة دوام الأثر:", "انخفاض القيمة"), ("المنافع البيئية:", "آثار غير مباشرة"), ("صحة وموثوقية البيانات:", "تحليل الشفافية"), ("توصية تعظيم الأثر:", "زيادة القيمة")],
        "تقرير رضا المستفيدين (CSI)": [("الخدمة محل التقييم:", "النشاط المشمول"), ("مؤشر الرضا الكلي:", "النسبة المئوية"), ("سهولة الوصول:", "قنوات التواصل"), ("احترافية الفريق:", "التعامل الميداني"), ("زمن الاستجابة:", "السرعة الفعلية"), ("تحليل الشكاوى:", "نقاط الألم"), ("معدل الولاء (NPS):", "التوصية بالجهة"), ("جودة المخرجات:", "المطابقة للمواصفات"), ("التوقعات المستقبلية:", "تطلعات الجمهور"), ("خطة تحسين التجربة:", "رفع مستوى الرضا")]
    },
    "مسار الاستراتيجية والمخاطر": {
        "دراسة جدوى ومصفوفة مخاطر": [("وصف الفرصة:", "الجدوى التشغيلية"), ("تحليل (PESTEL):", "سياسي واقتصادي"), ("تحديد الأخطار:", "مصفوفة الخطر"), ("احتمالية الحدوث:", "تكرار الخطر"), ("شدة الأثر:", "مدى الخسارة"), ("خطة التحوط:", "تخفيف المخاطر"), ("تحليل المنافسة:", "قوى بورتر"), ("الاستثمار (CAPEX):", "التكاليف التأسيسية"), ("نقطة التعادل:", "الربحية المتوقعة"), ("التوصية النهائية:", "قرار الاستثمار")],
        "تقرير المراجعة الاستراتيجية": [("تحقيق الأهداف:", "النسبة المئوية"), ("الانحراف الاستراتيجي:", "الفجوة الفعلية"), ("تقييم المحفظة:", "وجدوى المشاريع"), ("نقاط القوة التنافسية:", "ميزات الحماية"), ("نقاط الضعف الجوهرية:", "القيود الداخلية"), ("تحليل الفرص الناشئة:", "اتجاهات السوق"), ("تقييم الشراكات:", "فاعلية التعاون"), ("كفاءة توزيع الموارد:", "توجيه الأولويات"), ("مراجعة الرؤية:", "التوافق الكلي"), ("خارطة الطريق المحدثة:", "تعديلات الخطة")],
        "تقرير تحليل المنافسين": [("قائمة المنافسين:", "توصيف المنافسة"), ("تحليل الحصة السوقية:", "مدى الاستحواذ"), ("ميزات المنافسين:", "سر نجاحهم"), ("نقاط ضعف المنافسين:", "فرص الاختراق"), ("الفجوات السوقية:", "شرائح مهملة"), ("تحليل الأسعار:", "التموضع السعري"), ("قوة العلامة:", "ولاء العملاء"), ("استراتيجيات التسويق:", "قنوات الوصول"), ("عوائق التوسع:", "تحديات قانونية"), ("خطة الاستحواذ:", "تحقيق التميز")],
        "تقرير هندسة القيم (VE)": [("المهمة المستهدفة:", "المشروع أو المنتج"), ("تحليل الوظائف:", "الوظيفة الأساسية"), ("تكلفة الوظائف الثانوية:", "الهدر المالي"), ("البدائل المبتكرة:", "خيارات أقل كلفة"), ("حجم الوفر المتوقع:", "نسبة الخفض"), ("تأثير البدائل:", "ضمان الجودة"), ("الجدول الزمني:", "فترة التنفيذ"), ("قبول أصحاب المصلحة:", "موقف العميل"), ("تحليل المخاطر:", "أخطار التغيير"), ("قرار هندسة القيم:", "الاعتماد الفني")],
        "تقرير تقييم الجاهزية": [("هدف التحول:", "ما الذي سنغيره؟"), ("جاهزية التكنولوجيا:", "الأنظمة والأجهزة"), ("كفاءة الكادر البشري:", "مهارات التغيير"), ("الاستقرار المالي:", "ميزانية التحول"), ("دعم القيادة العليا:", "مدى الالتزام"), ("مقاومة التغيير:", "المعوقات النفسية"), ("وضوح اللوائح:", "دوانب قانونية"), ("تحليل المخاطر:", "توقف مؤقت"), ("نضج رقمي:", "مستوى 1-5"), ("خطة الإطلاق:", "قرار البدء")]
    },
    "مسار العمليات والإنتاجية": {
        "تقرير الإنجاز الدوري": [("المستهدفات:", "الأهداف الرقمية"), ("الإنجاز الفعلي:", "المنفذ فعلياً"), ("أسباب الهدر الزمني:", "لماذا تأخرنا؟"), ("كفاءة الموازنة:", "الصرف مقابل الإنجاز"), ("جودة المخرجات:", "المطابقة للمواصفات"), ("العوائق البيروقراطية:", "مشاكل الموافقات"), ("إنتاجية الفريق:", "أداء الموظفين"), ("تحديثات التوريد:", "وصول المواد"), ("الاحتياجات اللوجستية:", "المتطلبات القادمة"), ("خطة التصحيح:", "تدارك التأخير")],
        "تقرير تحليل الهدر": [("تصنيف الهدر:", "انتظار، نقل، عيوب"), ("موقع الهدر:", "القسم المحدد"), ("حجم الخسارة:", "تقدير الكلفة"), ("السبب الجذري:", "لماذا يحدث؟"), ("تأثير العميل:", "تأخر أو جودة"), ("إجراء كايزن:", "مقترح الإزالة"), ("كلفة التحسين:", "الاستثمار المطلوب"), ("مقاومة التغيير:", "تحديات التطبيق"), ("النتيجة المتوقعة:", "الوفر المضاف"), ("المسؤول:", "تحديد المعني")],
        "تقرير أداء الموردين": [("اسم المورد:", "الجهة محل التقييم"), ("الالتزام بمواعيد:", "الدقة في الوقت"), ("مطابقة الجودة:", "جودة التوريدات"), ("التنافسية السعرية:", "الأسعار مقابل السوق"), ("المرونة في التعامل:", "الطلبات الطارئة"), ("خدمات ما بعد البيع:", "الضمان والصيانة"), ("الوضع المالي:", "استقرار المورد"), ("سجل المخالفات:", "مشاكل سابقة"), ("مخاطر الاعتماد:", "احتکار أو ندرة"), ("القرار التعاقدي:", "تجديد أو استبعاد")],
        "تقرير إدارة الأزمات": [("توصيف الأزمة:", "طبيعة الحدث"), ("تاريخ الاستجابة:", "وقت الرصد والتدخل"), ("إجراءات الاحتواء:", "منع التوسع"), ("حجم الأضرار:", "الخسائر الأولية"), ("تفعيل خطة (BCP):", "بدائل التشغيل"), ("كفاءة الفريق:", "سرعة التحرك"), ("تحليل الفشل:", "لماذا حدثت؟"), ("خطة التعافي:", "خطوات العودة"), ("الدروس المستفادة:", "ضمان عدم التكرار"), ("توصيات التدقيق:", "تحديث السياسات")],
        "تقرير كفاءة الأداء (KPIs)": [("القسم المستهدف:", "نطاق القياس"), ("مؤشر زمن التنفيذ:", "المستهدف مقابل المحقق"), ("تكلفة المخرج:", "الصرف المعياري"), ("مؤشر الجودة:", "نسبة الإتقان"), ("تحليل أسباب التدني:", "لماذا لم نحقق؟"), ("المنجزات الاستثنائية:", "نجاحات خاصة"), ("كفاءة القيادة:", "الإدارة والتحفيز"), ("المقارنة بالعام:", "اتجاه الأداء"), ("تحديث المستهدفات:", "الأهداف الجديدة"), ("القرارات الناتجة:", "بناءً على القياس")]
    },
    "مسار العلاقات وصورة المؤسسة": {
        "تقرير التغطية الإعلامية": [("الرسالة الاستراتيجية:", "ماذا نريد أن نقول؟"), ("إحصائيات الوصول:", "عدد المشاهدات"), ("تحليل النبرة:", "إيجابي، سلبي"), ("قائمة الشركاء:", "المساهمون في النشر"), ("جودة المحتوى:", "تصميم وصياغة"), ("وصول الرسائل:", "فهم الجمهور للهدف"), ("كفاءة الإنفاق:", "القيمة مقابل الكلفة"), ("فجوات التواصل:", "فئات لم نصلها"), ("تهديدات السمعة:", "إشاعات مرصودة"), ("توصية العلاقات:", "تحسين الصورة")],
        "تقرير إدارة الفعاليات": [("نوع الفعالية:", "الهدف منها"), ("التنظيم اللوجستي:", "قاعات وضيافة"), ("تطبيق البروتوكول:", "الرسميات والتشريفات"), ("كفاءة إدارة الوقت:", "الالتزام بالجدول"), ("جودة المتحدثين:", "تقييم المحتوى"), ("التغطية الإعلامية:", "البث والتوثيق"), ("تفاعل الحضور:", "المناقشات والأسئلة"), ("المشاكل التقنية:", "سرعة حل الأعطال"), ("التقييم المالي:", "ضمن الموازنة"), ("الدروس المستفادة:", "تحسينات قادمة")],
        "تقرير المسؤولية المجتمعية": [("اسم المبادرة:", "وصف التدخل"), ("الموقع والفئة:", "من استهدفنا؟"), ("حجم المساهمة:", "مالية أو عينية"), ("الأثر الملموس:", "التغير الذي حدث"), ("مشاركة الموظفين:", "العمل التطوعي"), ("الشركاء المحليون:", "المجتمع المدني"), ("أثر السمعة:", "صورة الجهة"), ("ردود الأفعال:", "الإشادات الرسمية"), ("ارتباط (SDGs):", "الهدف الأممي"), ("اقتراح للمبادرة:", "بناءً على الاحتياج")],
        "تقرير الأزمات الإعلامية": [("طبيعة الأزمة:", "التهديد المرصود"), ("منصة الانتشار:", "أين بدأت؟"), ("سرعة الاستجابة:", "وقت أول رد"), ("إجراءات الاحتواء:", "البيانات المباشرة"), ("قوة الرد:", "الأدلة المقدمة"), ("دور الشركاء:", "دعم الحلفاء"), ("نبرة الجمهور:", "التراجع أو الاستمرار"), ("الأضرار المتبقية:", "آثار تحتاج علاج"), ("أداء المتحدث:", "الثبات والإقناع"), ("الخطة الوقائية:", "منع التكرار")],
        "تقرير قياس عائد الشراكات": [("الجهة الشريكة:", "طبيعة التعاون"), ("الأهداف المحددة:", "ما تم الاتفاق عليه"), ("المنافع العائدة:", "مالية أو فنية"), ("التزاماتنا الموفاة:", "ماذا قدمنا؟"), ("التنسيق الإداري:", "سلاسة التواصل"), ("عوائق التنفيذ:", "الخلافات والمشاكل"), ("أثر المكانة:", "السمعة السوقية"), ("التقييم المالي:", "ربحية الشراكة"), ("السرية والخصوصية:", "أمن المعلومات"), ("توصية الاستمرار:", "تجديد أو إنهاء")]
    }
}

# ==========================================
# 4. منطق الحفظ والتعامل مع الأخطاء
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
# 5. صفحات النظام
# ==========================================
def login_page():
    st.markdown('<div class="card-box" style="margin-top:50px;">', unsafe_allow_html=True)
    st.title("🏛️ دخول المنصة السيادية")
    uid = st.text_input("رقم الجوال (للدخول واستعادة التقارير):", placeholder="مثال: 774575749")
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
    st.info(f"المستشار: **{uid}** | الرصيد: **{balance} تقارير**")

    # بيانات الغلاف
    st.markdown("### 🏛️ أولاً: بيانات الغلاف (الإدارية)")
    org = st.text_input("الجهة المصدرة:", value=get_draft("org_name"))
    update_draft("org_name", org)
    loc = st.text_input("النطاق الجغرافي:", value=get_draft("loc_name"))
    update_draft("loc_name", loc)
    proj = st.text_input("اسم المشروع:", value=get_draft("proj_name"))
    update_draft("proj_name", proj)
    author = st.text_input("إعداد (الاسم والمنصب):", value=get_draft("author_name"))
    update_draft("author_name", author)

    st.markdown("---")
    pillar = st.selectbox("1. المسار الاستراتيجي:", list(methodology_db.keys()))
    report_type = st.selectbox("2. التقرير التخصصي المعتمد:", list(methodology_db[pillar].keys()))
    
    if st.session_state.current_report != report_type:
        st.session_state.current_report = report_type
        st.session_state.step = 1
        st.session_state.report_preview = ""
    
    questions = methodology_db[pillar][report_type]
    st.markdown(f"--- المرحلة الحالية: **{st.session_state.step} من 3** ---")
    
    if st.session_state.step == 1:
        st.subheader("📍 المرحلة 1: التشخيص والمطابقة")
        for i, (q, h) in enumerate(questions[:3]):
            ans = st.text_area(q, value=get_draft(f"q_{report_type}_{i}"), placeholder=h, key=f"k1_{report_type}_{i}")
            update_draft(f"q_{report_type}_{i}", ans)
        if st.button("التالي: التحليل الاستراتيجي ⬅️"): st.session_state.step = 2; st.rerun()

    elif st.session_state.step == 2:
        st.subheader("📊 المرحلة 2: تحليل الأسباب الجذرية")
        for i, (q, h) in enumerate(questions[3:7]):
            idx = i + 3
            ans = st.text_area(q, value=get_draft(f"q_{report_type}_{idx}"), placeholder=h, key=f"k2_{report_type}_{idx}")
            update_draft(f"q_{report_type}_{idx}", ans)
        if st.button("التالي: صناعة القرار ⬅️"): st.session_state.step = 3; st.rerun()
        if st.button("➡️ رجوع"): st.session_state.step = 1; st.rerun()

    elif st.session_state.step == 3:
        st.subheader("🎯 المرحلة 3: القرارات والاعتماد")
        for i, (q, h) in enumerate(questions[7:]):
            idx = i + 7
            ans = st.text_area(q, value=get_draft(f"q_{report_type}_{idx}"), placeholder=h, key=f"k3_{report_type}_{idx}")
            update_draft(f"q_{report_type}_{idx}", ans)
        recs = st.text_area("التوصيات الختامية:", value=get_draft(f"recs_{report_type}"))
        update_draft(f"recs_{report_type}", recs)
        
        if st.button("اعتماد وتوليد الوثيقة 📄"):
            if balance <= 0: st.error("⚠️ الرصيد صفر.")
            else:
                try:
                    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
                    # حل مشكلة الـ 429 والـ 404 بذكاء:
                    model = genai.GenerativeModel('gemini-1.5-flash')
                    data_feed = "".join([f"- {q}: {get_draft(f'q_{report_type}_{i}')}\n" for i, (q, _) in enumerate(questions)])
                    prompt = f"أنت مستشار استراتيجي سيادي خبير في {pillar}. صغ تقرير '{report_type}' لجهة '{org}' مشروع '{proj}'. البيانات: {data_feed}. التوصيات: {recs}. اللغة: رسمية، رصينة، نقاط مباشرة."
                    
                    with st.spinner("المحرك الذكي يقوم بالصياغة..."):
                        # محاولة التوليد مع خطة بديلة لو وجد زحام (Retry)
                        try:
                            res = model.generate_content(prompt)
                        except:
                            time.sleep(4) # انتظار 4 ثواني لو فيه ضغط
                            res = model.generate_content(prompt)
                        
                        st.session_state.report_preview = res.text
                        db["users"][uid]["balance"] -= 1
                        save_db(db)
                        st.success("تم التوليد بنجاح!")
                except Exception as e:
                    st.error(f"خطأ تقني: {e}")
        if st.button("➡️ رجوع"): st.session_state.step = 2; st.rerun()

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
