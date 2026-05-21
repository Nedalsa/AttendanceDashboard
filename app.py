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
st.subheader("② إعدادات الدوام")

# ── initialize session_state keys with defaults ───────────────────────────────
_defaults = {
    "cfg_shift_start":                    CONFIG_DEFAULT["shift_start"],
    "cfg_shift_end":                      CONFIG_DEFAULT["shift_end"],
    "cfg_grace_minutes":                  CONFIG_DEFAULT["grace_minutes"],
    "cfg_tardiness_base":                 CONFIG_DEFAULT["tardiness_base"],
    "cfg_tardiness_rounding":             CONFIG_DEFAULT["tardiness_rounding"],
    "cfg_tardiness_round_up_to":          CONFIG_DEFAULT["tardiness_round_up_to"],
    "cfg_early_tol":                      CONFIG_DEFAULT["early_departure_tolerance_minutes"],
    "cfg_early_round":                    CONFIG_DEFAULT["early_departure_round_up_to"],
    "cfg_overtime_base":                  CONFIG_DEFAULT.get("overtime_base","shift_end"),
    "cfg_ot_threshold":                   CONFIG_DEFAULT["overtime_threshold_minutes"],
    "cfg_ot_round":                       CONFIG_DEFAULT["overtime_round_down_to"],
    "cfg_deduct":                         CONFIG_DEFAULT["deduct_tardiness_from_overtime"],
    "cfg_missing_checkin":                CONFIG_DEFAULT["missing_checkin_action"],
    "cfg_missing_checkout":               CONFIG_DEFAULT["missing_checkout_action"],
}
for k, v in _defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

def _get_cfg():
    """Build cfg dict from session_state — always fresh."""
    return {
        "shift_start":                       st.session_state.cfg_shift_start,
        "shift_end":                         st.session_state.cfg_shift_end,
        "grace_minutes":                     st.session_state.cfg_grace_minutes,
        "tardiness_base":                    st.session_state.cfg_tardiness_base,
        "tardiness_rounding":                st.session_state.cfg_tardiness_rounding,
        "tardiness_round_up_to":             st.session_state.cfg_tardiness_round_up_to,
        "early_departure_tolerance_minutes": st.session_state.cfg_early_tol,
        "early_departure_round_up_to":       st.session_state.cfg_early_round,
        "overtime_base":                     st.session_state.cfg_overtime_base,
        "overtime_threshold_minutes":        st.session_state.cfg_ot_threshold,
        "overtime_round_down_to":            st.session_state.cfg_ot_round,
        "deduct_tardiness_from_overtime":    st.session_state.cfg_deduct,
        "missing_checkin_action":            st.session_state.cfg_missing_checkin,
        "missing_checkout_action":           st.session_state.cfg_missing_checkout,
        "absent_status":                     CONFIG_DEFAULT["absent_status"],
    }

st.info("اضبط أوقات الدوام وطريقة احتساب التأخير والخروج المبكر والإضافي.")

# ① أوقات الدوام
with st.expander("① أوقات الدوام الأساسية", expanded=True):
    c1, c2, c3 = st.columns(3)
    with c1:
        st.text_input("بداية الدوام", key="cfg_shift_start", help="مثال: 08:00")
    with c2:
        st.text_input("نهاية الدوام", key="cfg_shift_end", help="مثال: 15:00")
    with c3:
        st.number_input("دقائق السماح للدخول", key="cfg_grace_minutes",
                        min_value=0, max_value=120, step=5,
                        help="الموظف لا يُعدّ متأخراً خلال هذه الدقائق بعد بداية الدوام")
    st.caption(
        f"بداية الدوام {st.session_state.cfg_shift_start} | "
        f"نهاية الدوام {st.session_state.cfg_shift_end} | "
        f"سماح الدخول {st.session_state.cfg_grace_minutes} دقيقة"
    )

# ② التأخير
with st.expander("② إعدادات التأخير", expanded=True):
    st.radio(
        "احتساب التأخير من",
        options=["shift_start", "after_grace"],
        format_func=lambda x: {
            "shift_start": f"من بداية الدوام ({st.session_state.cfg_shift_start})",
            "after_grace": f"من بعد وقت السماح ({st.session_state.cfg_shift_start} + {st.session_state.cfg_grace_minutes} د)",
        }[x],
        key="cfg_tardiness_base",
        help="من بداية الدوام = أشد | من بعد السماح = أعدل",
    )
    st.radio(
        "طريقة احتساب التأخير",
        options=["daily", "total", "none"],
        format_func=lambda x: {
            "daily": "تقريب يومي — كل يوم منفرداً (أشد)",
            "total": "تقريب تراكمي — مجموع الشهر ثم تقريب (أعدل)",
            "none":  "بدون تقريب — دقيقة بدقيقة (الأدق)",
        }[x],
        key="cfg_tardiness_rounding",
    )
    if st.session_state.cfg_tardiness_rounding != "none":
        st.selectbox(
            "وحدة تقريب التأخير لأعلى",
            [1, 15, 30, 60],
            format_func=lambda x: "بدون تقريب (1 د)" if x == 1 else f"{x} دقيقة",
            key="cfg_tardiness_round_up_to",
        )

# ③ الخروج المبكر
with st.expander("③ إعدادات الخروج المبكر", expanded=True):
    c1, c2 = st.columns(2)
    with c1:
        st.number_input("دقائق السماح للخروج المبكر", key="cfg_early_tol",
                        min_value=0, max_value=120, step=5,
                        help="خروج ضمن هذه الدقائق قبل نهاية الدوام لا يُحتسب مبكراً")
    with c2:
        st.selectbox("وحدة تقريب الخروج المبكر لأعلى",
                     [1, 15, 30, 60],
                     format_func=lambda x: "بدون تقريب (1 د)" if x == 1 else f"{x} دقيقة",
                     key="cfg_early_round")
    st.caption(
        f"نهاية الدوام {st.session_state.cfg_shift_end} | "
        f"سماح الخروج {st.session_state.cfg_early_tol} دقيقة"
    )

# ④ الإضافي
with st.expander("④ إعدادات الإضافي", expanded=True):
    st.radio(
        "احتساب الإضافي من",
        options=["shift_end", "after_threshold"],
        format_func=lambda x: {
            "shift_end":       f"من نهاية الدوام ({st.session_state.cfg_shift_end}) — الأشمل",
            "after_threshold": f"من بعد حد الإضافي ({st.session_state.cfg_shift_end} + {st.session_state.cfg_ot_threshold} د) — الأدق",
        }[x],
        key="cfg_overtime_base",
        help="من نهاية الدوام: خرج 16:00 وحد 30 د → إضافي 60 د. من بعد الحد → 30 د فقط.",
    )
    c1, c2 = st.columns(2)
    with c1:
        st.number_input("حد الإضافي (دقائق الانتظار)", key="cfg_ot_threshold",
                        min_value=0, max_value=180, step=15,
                        help="الموظف يجب أن يبقى هذه الدقائق على الأقل بعد نهاية الدوام")
    with c2:
        st.selectbox("وحدة تقريب الإضافي لأسفل",
                     [1, 15, 30, 60],
                     format_func=lambda x: "بدون تقريب (1 د)" if x == 1 else f"{x} دقيقة",
                     key="cfg_ot_round")
    st.checkbox("طرح التأخير من الإضافي الصافي", key="cfg_deduct",
                help="تأخر 30 د وعمل إضافي 60 د → الإضافي الصافي = 30 د")

# ⑤ الحالات غير المكتملة
with st.expander("⑤ الحالات غير المكتملة", expanded=False):
    c1, c2 = st.columns(2)
    with c1:
        st.radio(
            "حاضر بدون تسجيل دخول",
            options=["ignore", "absent"],
            format_func=lambda x: {
                "ignore": "تنبيه فقط",
                "absent": "احتسابه غائباً",
            }[x],
            key="cfg_missing_checkin",
        )
    with c2:
        st.radio(
            "حاضر بدون تسجيل خروج",
            options=["ignore", "shift_end"],
            format_func=lambda x: {
                "ignore":    "تنبيه فقط — لا يُحسب إضافي",
                "shift_end": f"اعتبار الخروج نهاية الدوام ({st.session_state.cfg_shift_end})",
            }[x],
            key="cfg_missing_checkout",
        )
    st.caption("هذه الحالات تظهر في عمود الملاحظة بالتقرير.")

st.divider()
# ═══════════════════════════════ ③ PROCESS ═══════════════════════════════════
st.subheader("③ معالجة وتصدير")

if st.button("▶ معالجة البيانات", type="primary", disabled=(att_file is None)):
    with st.spinner("جاري المعالجة..."):
        try:
            df_att = load_attendance(att_file)
            df_ind = pd.DataFrame(columns=["رقم الموظف","التاريخ","النوع","ملاحظة"])
            if exc_file:
                df_ind = load_exceptions(exc_file)
            detail, summary = run(df_att, df_ind, _get_cfg())
            st.session_state.excel_bytes = build_excel_report(
                detail, summary, _get_cfg())
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
