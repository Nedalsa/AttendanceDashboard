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
st.markdown("<h3 style='text-align:right'>① الملفات</h3>", unsafe_allow_html=True)


col_a, col_b = st.columns(2)
with col_a:
    att_file = st.file_uploader(
        "ملف الدوام", type=["xlsx","xls"],
        help="أعمدة: رقم الموظف، اسم الموظف، التاريخ (YYYY/MM/DD)، أول دخول، آخر خروج، الحالة، المديرية",
    )
with col_b:
    exc_file = st.file_uploader(
        "ملف الاستثناءات (اختياري)", type=["xlsx","xls"],
        help="شيت واحد: المبررات الفردية — السنة والشهر واليوم في أعمدة منفصلة",
    )

with st.expander("📥 تحميل قالب ملف الاستثناءات"):
    st.markdown("""
التاريخ يجب أن يكون بصيغة `YYYY/MM/DD` — مثال: `2026/04/01`

شيت واحد فقط: **المبررات الفردية**

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
st.markdown("<h3 style='text-align:right'>② إعدادات الدوام</h3>", unsafe_allow_html=True)

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
            "دقائق السماح للدخول",
            min_value=0,
            max_value=120,
            value=int(cfg["grace_minutes"]),
            step=5,
            help="مثال: إذا بداية الدوام 08:00 والسماح للدخول 15 دقيقة، لا يُعدّ الموظف متأخراً قبل 08:15"
        )

    st.caption(
        f"حسب الإعدادات الحالية: بداية الدوام {cfg['shift_start']}، "
        f"نهاية الدوام {cfg['shift_end']}، وسماح الدخول {cfg['grace_minutes']} دقيقة."
    )


# ─────────────────────────────
# 2) إعدادات التأخير
# ─────────────────────────────
with st.expander("② إعدادات التأخير", expanded=True):

    cfg["tardiness_base"] = st.radio(
        "احتساب التأخير من",
        options=["shift_start", "after_grace"],
        format_func=lambda x: {
            "shift_start": f"من بداية الدوام ({cfg['shift_start']}) — الموظف يدفع من أول دقيقة",
            "after_grace": f"من بعد وقت السماح ({cfg['shift_start']} + {cfg['grace_minutes']} د) — يدفع الزيادة فقط",
        }[x],
        index=["shift_start", "after_grace"].index(cfg["tardiness_base"]),
        horizontal=False,
        help="مثال: جاء 08:20 مع سماح 15 د. من بداية الدوام = 20 د تأخير. من بعد السماح = 5 د فقط."
    )

    cfg["tardiness_rounding"] = st.radio(
        "طريقة احتساب التأخير",
        options=["daily", "total", "none"],
        format_func=lambda x: {
            "daily": "تقريب يومي — كل يوم يُقرَّب منفرداً (أشد)",
            "total": "تقريب تراكمي — يجمع الدقائق أولاً ثم يقرّب مرة واحدة (أعدل)",
            "none":  "بدون تقريب — دقيقة بدقيقة (الأدق)",
        }[x],
        index=["daily", "total", "none"].index(cfg["tardiness_rounding"]),
        horizontal=False,
        help="التقريب اليومي أشد على الموظف، التراكمي أعدل، وبدون تقريب يعطي الدقائق الفعلية."
    )

    if cfg["tardiness_rounding"] != "none":
        cfg["tardiness_round_up_to"] = st.selectbox(
            "وحدة تقريب التأخير لأعلى",
            [1, 15, 30, 60],
            index=[1, 15, 30, 60].index(int(cfg["tardiness_round_up_to"])),
            format_func=lambda x: "بدون تقريب فعلي (1 د)" if x == 1 else f"{x} دقيقة",
        )
    else:
        cfg["tardiness_round_up_to"] = 1

    st.caption(
        f"مثال: تأخير 7 دقائق مع تقريب {cfg['tardiness_round_up_to']} د → يُحتسب "
        f"{__import__('math').ceil(7 / cfg['tardiness_round_up_to']) * cfg['tardiness_round_up_to']} دقيقة."
        if cfg["tardiness_rounding"] != "none" else
        "بدون تقريب — الدقائق الفعلية فقط."
    )


# ─────────────────────────────
# 3) إعدادات الخروج المبكر
# ─────────────────────────────
with st.expander("③ إعدادات الخروج المبكر", expanded=True):
    c1, c2 = st.columns(2)

    with c1:
        cfg["early_departure_tolerance_minutes"] = st.number_input(
            "دقائق السماح للخروج المبكر",
            min_value=0,
            max_value=120,
            value=int(cfg["early_departure_tolerance_minutes"]),
            step=5,
            help="مثال: نهاية الدوام 15:00 والسماح 5 دقائق → الخروج 14:55 لا يُحتسب مبكراً."
        )

    with c2:
        cfg["early_departure_round_up_to"] = st.selectbox(
            "وحدة تقريب الخروج المبكر لأعلى",
            [1, 15, 30, 60],
            index=[1, 15, 30, 60].index(int(cfg["early_departure_round_up_to"])),
            format_func=lambda x: "بدون تقريب فعلي (1 د)" if x == 1 else f"{x} دقيقة",
        )

    st.caption(
        f"نهاية الدوام {cfg['shift_end']}، سماح الخروج {cfg['early_departure_tolerance_minutes']} دقيقة — "
        f"أي خروج قبل {cfg['shift_end']} بأكثر من {cfg['early_departure_tolerance_minutes']} دقيقة يُحتسب مبكراً."
    )


# ─────────────────────────────
# 4) إعدادات الإضافي
# ─────────────────────────────
with st.expander("④ إعدادات الإضافي", expanded=True):

    cfg["overtime_base"] = st.radio(
        "احتساب الإضافي من",
        options=["shift_end", "after_threshold"],
        format_func=lambda x: {
            "shift_end":       f"من نهاية الدوام ({cfg['shift_end']}) — الأشمل للموظف",
            "after_threshold": f"من بعد حد الإضافي ({cfg['shift_end']} + {cfg['overtime_threshold_minutes']} د) — الأدق",
        }[x],
        index=["shift_end","after_threshold"].index(cfg.get("overtime_base","shift_end")),
        horizontal=False,
        help="مثال: خرج 16:00 وحد الإضافي 30 د. من نهاية الدوام = 60 د. من بعد الحد = 30 د فقط."
    )

    c1, c2 = st.columns(2)

    with c1:
        cfg["overtime_threshold_minutes"] = st.number_input(
            "حد الإضافي (دقائق الانتظار)",
            min_value=0,
            max_value=180,
            value=int(cfg["overtime_threshold_minutes"]),
            step=15,
            help="الموظف يجب أن يبقى هذه الدقائق الإضافية على الأقل حتى يُحتسب له إضافي."
        )

    with c2:
        cfg["overtime_round_down_to"] = st.selectbox(
            "وحدة تقريب الإضافي لأسفل",
            [1, 15, 30, 60],
            index=[1, 15, 30, 60].index(int(cfg["overtime_round_down_to"])),
            format_func=lambda x: "بدون تقريب فعلي (1 د)" if x == 1 else f"{x} دقيقة",
        )

    cfg["deduct_tardiness_from_overtime"] = st.checkbox(
        "طرح التأخير من الإضافي الصافي",
        value=bool(cfg["deduct_tardiness_from_overtime"]),
        help="مثال: تأخر 30 د وعمل إضافي 60 د → الإضافي الصافي = 30 د فقط."
    )

    st.caption(
        f"الإضافي يبدأ بعد {cfg['overtime_threshold_minutes']} دقيقة من {cfg['shift_end']}، "
        f"ويُقرَّب لأسفل لأقرب {cfg['overtime_round_down_to']} دقيقة."
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
                "ignore": "تنبيه فقط — يظهر في الملاحظة",
                "absent": "احتسابه غائباً",
            }[x],
            index=["ignore", "absent"].index(cfg["missing_checkin_action"]),
            help="الأفضل: تنبيه فقط حتى تراجع الحالة يدوياً قبل الاحتساب."
        )

    with c2:
        cfg["missing_checkout_action"] = st.radio(
            "حاضر بدون تسجيل خروج",
            options=["ignore", "shift_end"],
            format_func=lambda x: {
                "ignore":    "تنبيه فقط — لا يُحسب إضافي",
                "shift_end": f"اعتبار الخروج نهاية الدوام ({cfg['shift_end']})",
            }[x],
            index=["ignore", "shift_end"].index(cfg["missing_checkout_action"]),
            help="الأفضل: تنبيه فقط، لأن افتراض وقت خروج قد يعطي نتيجة غير دقيقة."
        )

    st.caption(
        "هذه الحالات تظهر في عمود الملاحظة بالتقرير حتى يمكن مراجعتها لاحقاً."
    )


st.session_state.cfg = cfg
st.divider()

# ═══════════════════════════════ ③ PROCESS ═══════════════════════════════════
st.markdown("<h3 style='text-align:right'>③ معالجة وتصدير</h3>", unsafe_allow_html=True)

if st.button("▶ معالجة البيانات", type="primary", disabled=(att_file is None)):
    with st.spinner("جاري المعالجة..."):
        try:
            df_att = load_attendance(att_file)
            df_ind = pd.DataFrame(columns=["رقم الموظف","التاريخ","النوع","ملاحظة"])
            if exc_file:
                df_ind = load_exceptions(exc_file)
            detail, summary = run(df_att, df_ind, st.session_state.cfg)
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

st.divider()
st.caption("Developed by **Nidal Al Saqqa** · نظام تحليل الدوام والحضور")
