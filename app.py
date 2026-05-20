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

st.info(
    "اضبط أوقات الدوام وطريقة احتساب التأخير والخروج المبكر والإضافي. "
    "القيم المختارة هنا ستستخدم مباشرة في التقرير."
)

# ─────────────────────────────
# 1) أوقات الدوام الأساسية
# ─────────────────────────────
with st.expander("① أوقات الدوام الأساسية", expanded=True):
    c1, c2, c3 = st.columns(3)

    with c1:
        cfg["shift_start"] = st.text_input(
            "بداية الدوام",
            value=cfg["shift_start"],
            help="مثال: 08:00"
        )

    with c2:
        cfg["shift_end"] = st.text_input(
            "نهاية الدوام",
            value=cfg["shift_end"],
            help="مثال: 15:00"
        )

    with c3:
        cfg["grace_minutes"] = st.number_input(
            "دقائق السماح",
            min_value=0,
            max_value=120,
            value=int(cfg["grace_minutes"]),
            step=5,
            help="مثال: إذا بداية الدوام 08:00 والسماح 15 دقيقة، لا يحسب التأخير قبل 08:15"
        )

    st.caption(
        f"حسب الإعدادات الحالية: بداية الدوام {cfg['shift_start']}، "
        f"نهاية الدوام {cfg['shift_end']}، والسماح {cfg['grace_minutes']} دقيقة."
    )


# ─────────────────────────────
# 2) إعدادات التأخير
# ─────────────────────────────
with st.expander("② إعدادات التأخير", expanded=True):
    st.markdown("#### احتساب التأخير")

    cfg["tardiness_base"] = st.radio(
        "احتساب التأخير من",
        options=["shift_start", "after_grace"],
        format_func=lambda x: {
            "shift_start": f"من بداية الدوام ({cfg['shift_start']})",
            "after_grace": f"بعد وقت السماح ({cfg['shift_start']} + {cfg['grace_minutes']} دقيقة)",
        }[x],
        index=["shift_start", "after_grace"].index(cfg["tardiness_base"]),
        horizontal=False,
        help="إذا اخترت بعد وقت السماح، فالموظف الذي يدخل 08:20 مع سماح حتى 08:15 يحسب عليه 5 دقائق فقط."
    )

    st.markdown("#### طريقة تجميع التأخير")

    cfg["tardiness_rounding"] = st.radio(
        "طريقة احتساب التأخير",
        options=["daily", "total", "none"],
        format_func=lambda x: {
            "daily": "تقريب يومي: كل يوم منفردا",
            "total": "تقريب تراكمي: مجموع الشهر",
            "none": "بدون تقريب: دقيقة بدقيقة",
        }[x],
        index=["daily", "total", "none"].index(cfg["tardiness_rounding"]),
        horizontal=False,
        help="التقريب اليومي أشد، والتقريب التراكمي أعدل غالبا، وبدون تقريب يعطي الدقائق الفعلية."
    )

    if cfg["tardiness_rounding"] != "none":
        cfg["tardiness_round_up_to"] = st.selectbox(
            "تقريب التأخير لأعلى",
            [1, 15, 30, 60],
            index=[1, 15, 30, 60].index(int(cfg["tardiness_round_up_to"])),
            format_func=lambda x: "بدون تقريب فعلي" if x == 1 else f"{x} دقيقة",
        )
    else:
        cfg["tardiness_round_up_to"] = 1

    st.caption(
        "مثال: إذا كان التأخير 7 دقائق والتقريب 15 دقيقة، يصبح التأخير المحتسب 15 دقيقة."
    )


# ─────────────────────────────
# 3) إعدادات الخروج المبكر
# ─────────────────────────────
with st.expander("③ إعدادات الخروج المبكر", expanded=True):
    c1, c2 = st.columns(2)

    with c1:
        cfg["early_departure_tolerance_minutes"] = st.number_input(
            "سماح الخروج المبكر",
            min_value=0,
            max_value=120,
            value=int(cfg["early_departure_tolerance_minutes"]),
            step=5,
            help="مثال: إذا نهاية الدوام 15:00 والسماح 5 دقائق، فالخروج 14:55 لا يحسب مبكرا."
        )

    with c2:
        cfg["early_departure_round_up_to"] = st.selectbox(
            "تقريب الخروج المبكر لأعلى",
            [1, 15, 30, 60],
            index=[1, 15, 30, 60].index(int(cfg["early_departure_round_up_to"])),
            format_func=lambda x: "بدون تقريب فعلي" if x == 1 else f"{x} دقيقة",
        )

    st.caption(
        "هذا الإعداد مهم حتى لا يتم احتساب خروج مثل 14:59 أو 14:58 كخروج مبكر كبير بسبب التقريب."
    )


# ─────────────────────────────
# 4) إعدادات الإضافي
# ─────────────────────────────
with st.expander("④ إعدادات الإضافي", expanded=True):
    c1, c2 = st.columns(2)

    with c1:
        cfg["overtime_threshold_minutes"] = st.number_input(
            "حد الإضافي",
            min_value=0,
            max_value=180,
            value=int(cfg["overtime_threshold_minutes"]),
            step=15,
            help="مثال: إذا نهاية الدوام 15:00 وحد الإضافي 30 دقيقة، يبدأ الإضافي بعد 15:30."
        )

    with c2:
        cfg["overtime_round_down_to"] = st.selectbox(
            "تقريب الإضافي لأسفل",
            [1, 15, 30, 60],
            index=[1, 15, 30, 60].index(int(cfg["overtime_round_down_to"])),
            format_func=lambda x: "بدون تقريب فعلي" if x == 1 else f"{x} دقيقة",
        )

    cfg["deduct_tardiness_from_overtime"] = st.checkbox(
        "طرح التأخير من الإضافي الصافي",
        value=bool(cfg["deduct_tardiness_from_overtime"]),
        help="إذا كان الموظف متأخرا ولديه إضافي في نفس اليوم، يتم خصم التأخير من الإضافي."
    )

    st.caption(
        "الإضافي يحسب بعد نهاية الدوام وحد الإضافي، ثم يقرب للأسفل حسب الإعداد المختار."
    )


# ─────────────────────────────
# 5) الحالات غير المكتملة
# ─────────────────────────────
with st.expander("⑤ الحالات غير المكتملة", expanded=False):
    c1, c2 = st.columns(2)

    with c1:
        cfg["missing_checkin_action"] = st.radio(
            "حاضر بدون تسجيل دخول",
            options=["ignore", "absent"],
            format_func=lambda x: {
                "ignore": "تنبيه فقط",
                "absent": "احتسابه غائبا",
            }[x],
            index=["ignore", "absent"].index(cfg["missing_checkin_action"]),
            help="الأفضل غالبا: تنبيه فقط، حتى تتم مراجعة الحالة يدويا."
        )

    with c2:
        cfg["missing_checkout_action"] = st.radio(
            "حاضر بدون تسجيل خروج",
            options=["ignore", "shift_end"],
            format_func=lambda x: {
                "ignore": "تنبيه فقط",
                "shift_end": "اعتبار الخروج نهاية الدوام",
            }[x],
            index=["ignore", "shift_end"].index(cfg["missing_checkout_action"]),
            help="الأفضل غالبا: تنبيه فقط، لأن افتراض وقت خروج قد يعطي نتيجة غير دقيقة."
        )

    st.caption(
        "هذه الحالات تظهر في التقرير مع ملاحظة حتى يمكن مراجعتها لاحقا."
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
