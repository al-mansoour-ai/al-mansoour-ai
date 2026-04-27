# -*- coding: utf-8 -*-
import streamlit as st
import google.generativeai as genai
import json, os, uuid

# ==========================================
# إعداد الصفحة
# ==========================================
st.set_page_config(page_title="منصة المنصور السيادية", layout="wide")

# ==========================================
# البرومبت المركزي (العقل)
# ==========================================
SYSTEM_PROMPT = """
أنت محرك استشاري احترافي لإنتاج التقارير التنفيذية.

المطلوب:
تحويل البيانات إلى تقرير يحتوي على:
1. الملخص التنفيذي
2. تعريف المشكلة
3. التحليل
4. الأسباب الجذرية
5. المخاطر
6. الحلول
7. خطة التنفيذ
8. مؤشرات الأداء
9. الأثر المالي
10. التوصيات

قواعد:
- أسلوب تنفيذي واضح
- بدون حشو
- حلول قابلة للتطبيق
"""

# ==========================================
# قاعدة البيانات
# ==========================================
DB_FILE = "db.json"

def load_db():
    if not os.path.exists(DB_FILE):
        with open(DB_FILE, "w", encoding="utf-8") as f:
            json.dump({"users": {}}, f)
    with open(DB_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

db = load_db()

# ==========================================
# واجهة المستخدم
# ==========================================
st.title("🏛️ منصة المنصور السيادية")

email = st.text_input("البريد الإلكتروني")
project = st.text_input("اسم المشروع")
report_type = st.selectbox("نوع التقرير", ["تقرير فني", "تقرير تدريبي", "تقرير استراتيجي"])

notes = st.text_area("أدخل البيانات أو الملاحظات")

# ==========================================
# زر التوليد
# ==========================================
if st.button("توليد التقرير"):

    if not email or not notes:
        st.warning("أدخل البيانات أولاً")
    else:
        try:
            genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
            model = genai.GenerativeModel('gemini-1.5-flash')

            user_input = f"""
نوع التقرير: {report_type}
المشروع: {project}
البيانات:
{notes}
"""

            final_prompt = SYSTEM_PROMPT + "\n\n" + user_input

            with st.spinner("جاري التوليد..."):
                res = model.generate_content(final_prompt)

            st.success("تم إنشاء التقرير")
            st.write(res.text)

        except Exception as e:
            st.error(f"خطأ: {e}")
