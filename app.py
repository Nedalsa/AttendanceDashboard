"""
نظام تحليل الدوام والحضور
"""

import streamlit as st
import pandas as pd
from processor import (load_attendance, load_exceptions, run,
                       CONFIG_DEFAULT, JUSTIFICATION_TYPES)
from exporter import build_excel_report, build_template

st.set_page_config(
    page_title="نظام تحليل الدوام",
    page_icon="📊",
    layout="centered",
)
st.markdown("<style>body,.main{direction:rtl}</style>", unsafe_allow_html=True)

# ── session state ─────────────────────────────────────────────────────────────
for key, val in [("detail", None), ("summary", None),
                 ("cfg", CONFIG_DEFAULT.copy()), ("excel_bytes", None)]:
    if key not in st.session_state:
        st.session_state[key] = val

# ─────────────────────────────────────────────────────────────────────────────
st.title("📊 نظام تحليل الدوام والحضور")
st.caption("ارفع ملفي الدوام والاستثناءات ← اضغط معالجة ← حمّل التقرير")
st.divider()

# ═══════════════════════════════ ① UPLOADS ═══════════════════════════════════
st.subheader("① الملفات")

col_a, col_b = st.columns(2)
with col_a:
    att_file = st.file_uploader(
        "ملف الدوام", type=["xlsx","xls"],
        help="أعمدة: رقم الموظف، اسم الموظف، التاريخ (YYYY/MM/DD)، أول دخول، آخر خروج، الحالة، المديرية",
    )
with col_b:
    exc_file = st.file_uploader(
        "ملف الاستثناءات (اختياري)", type=["xlsx","xls"],
        help="شيت ١: العطل الرسمية | شيت ٢: المبررات الفردية — التاريخ بصيغة YYYY/MM/DD",
    )

with st.expander("📥 تحميل قالب ملف الاستثناءات"):
    st.markdown("""
التاريخ يجب أن يكون بصيغة `YYYY/MM/DD` — مثال: `2026/04/01`

**شيت ١ — العطل الرسمية** | **شيت ٢ — المبررات الفردية**

أنواع مقبولة: `غياب مبرر` / `تأخير مبرر` / `خروج مبكر مبرر`
    """)
    st.download_button(
        "تحميل القالب",
        data=build_template(),
        file_name="exceptions_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.divider()

# ═══════════════════════════════ ② CONFIG ════════════════════════════════════
st.subheader("② إعدادات الدوام")

cfg = st.session_state.cfg

c1, c2, c3 = st.columns(3)
with c1:
    cfg["shift_start"] = st.text_input("بداية الدوام", value=cfg["shift_start"])
    cfg["shift_end"]   = st.text_input("نهاية الدوام",  value=cfg["shift_end"])
    cfg["grace_minutes"] = st.number_input(
        "دقائق السماح", min_value=0, max_value=60,
        value=int(cfg["grace_minutes"]), step=5)

with c2:
    cfg["tardiness_base"] = st.radio(
        "احتساب التأخير من",
        options=["shift_start", "after_grace"],
        format_func=lambda x: f"بداية الدوام ({cfg['shift_start']})" if x == "shift_start" else f"بعد وقت السماح ({cfg['shift_start']} + {cfg['grace_minutes']} دقيقة)",
        index=["shift_start","after_grace"].index(cfg["tardiness_base"]),
    )
    cfg["tardiness_rounding"] = st.radio(
        "طريقة احتساب التأخير",
        options=["daily", "total", "none"],
        format_func=lambda x: {
            "daily": "تقريب يومي (كل يوم منفرداً)",
            "total": "تقريب تراكمي (مجموع الشهر)",
            "none":  "بدون تقريب (دقيقة بدقيقة)",
        }[x],
        index=["daily","total","none"].index(cfg["tardiness_rounding"]),
    )
    if cfg["tardiness_rounding"] != "none":
        cfg["tardiness_round_up_to"] = st.selectbox(
            "تقريب التأخير لأعلى (د)", [1,15,30,60],
            index=[1,15,30,60].index(int(cfg["tardiness_round_up_to"])))

with c3:
    cfg["overtime_threshold_minutes"] = st.number_input(
        "حد الإضافي (د)", min_value=0, max_value=120,
        value=int(cfg["overtime_threshold_minutes"]), step=15)
    cfg["overtime_round_down_to"] = st.selectbox(
        "تقريب الإضافي لأسفل (د)", [1,15,30,60],
        index=[1,15,30,60].index(int(cfg["overtime_round_down_to"])))
    cfg["missing_checkin_action"] = st.radio(
        "حاضر بدون تسجيل دخول",
        options=["ignore", "absent"],
        format_func=lambda x: "تنبيه فقط" if x == "ignore" else "احتسابه غائباً",
        index=["ignore","absent"].index(cfg["missing_checkin_action"]),
    )
    cfg["missing_checkout_action"] = st.radio(
        "حاضر بدون تسجيل خروج",
        options=["ignore", "shift_end"],
        format_func=lambda x: "تنبيه فقط" if x == "ignore" else "اعتبار الخروج نهاية الدوام",
        index=["ignore","shift_end"].index(cfg["missing_checkout_action"]),
    )

cfg["deduct_tardiness_from_overtime"] = st.checkbox(
    "طرح التأخير من الإضافي الصافي",
    value=bool(cfg["deduct_tardiness_from_overtime"]),
)
st.session_state.cfg = cfg
st.divider()

# ═══════════════════════════════ ③ PROCESS ═══════════════════════════════════
st.subheader("③ معالجة وتصدير")

if st.button("▶ معالجة البيانات", type="primary", disabled=(att_file is None)):
    with st.spinner("جاري المعالجة..."):
        try:
            df_att = load_attendance(att_file)
            df_hol = pd.DataFrame(columns=["التاريخ","الوصف"])
            df_ind = pd.DataFrame(columns=["رقم الموظف","التاريخ","النوع","ملاحظة"])
            if exc_file:
                df_hol, df_ind = load_exceptions(exc_file)
            detail, summary = run(df_att, df_hol, df_ind, st.session_state.cfg)
            st.session_state.excel_bytes = build_excel_report(
                detail, summary, st.session_state.cfg)
            st.session_state.detail  = detail
            st.session_state.summary = summary
        except Exception as e:
            st.error(f"خطأ: {e}")
            st.stop()

if st.session_state.excel_bytes is not None:
    detail  = st.session_state.detail
    summary = st.session_state.summary

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("الموظفون",      summary["رقم الموظف"].nunique())
    m2.metric("الأشهر",        detail["الشهر"].nunique())
    m3.metric("غياب غير مبرر", int(summary["غياب_غير_مبرر"].sum()))
    m4.metric("معدل الحضور",   f"{summary['معدل_الحضور'].mean()*100:.1f}%")

    st.success("✅ التقرير جاهز")
    months = "_".join(sorted(detail["الشهر"].unique()))
    st.download_button(
        label="⬇ تحميل تقرير Excel",
        data=st.session_state.excel_bytes,
        file_name=f"attendance_report_{months}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )
