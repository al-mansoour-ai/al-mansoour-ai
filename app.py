# ==========================================
# 1. نظام إدارة الباقات والرصيد (Subscription Logic)
# ==========================================
if "user_balance" not in st.session_state:
    st.session_state.user_balance = 0  # الرصيد الافتراضي

st.sidebar.markdown("### 💳 محفظة الباقات")
st.sidebar.write(f"رصيدك الحالي: **{st.session_state.user_balance} تقارير**")

# حقل إدخال رمز الشحن
activation_code = st.sidebar.text_input("أدخل رمز شحن الباقة:", type="password")
if st.sidebar.button("تفعيل الرمز"):
    # نظام الرموز (يمكنك تغيير هذه الرموز أو توليدها يدوياً للعملاء)
    codes = {
        "MANSOUR_3": 3,   # باقة المنجز
        "EXPERT_10": 10,  # باقة الخبير
        "STRATEGIC_VIP": 100 # الباقة السيادية
    }
    if activation_code in codes:
        st.session_state.user_balance += codes[activation_code]
        st.sidebar.success(f"تم شحن {codes[activation_code]} تقارير بنجاح!")
    else:
        st.sidebar.error("الرمز غير صحيح أو مستخدم.")

st.sidebar.markdown("---")
st.sidebar.markdown("""
**لشراء رموز الشحن:**
1. الكريمي (774575749)
2. ون كاش / جيب
3. تواصل عبر واتساب لطلب الرمز فوراً
""")

# ==========================================
# 2. تعديل زر التوليد ليعمل بنظام الخصم
# ==========================================
if st.button("توليد الوثيقة المؤسسية"):
    if st.session_state.user_balance <= 0:
        st.error("⚠️ رصيدك غير كافٍ. يرجى شحن باقة جديدة للاستمرار في توليد التقارير العالمية.")
        st.markdown(f"[شراء باقة عبر واتساب]({whatsapp_link})")
    elif not (entity_name and project_name and location and author_name):
        st.error("خطأ منهجي: البيانات الإدارية ناقصة.")
    else:
        # هنا يبدأ التوليد (كود Gemini السابق)
        # ... [كود التوليد] ...
        
        # بعد النجاح، يتم خصم تقرير واحد من الرصيد
        st.session_state.user_balance -= 1
        st.success(f"تم التوليد بنجاح. الرصيد المتبقي: {st.session_state.user_balance} تقرير.")
