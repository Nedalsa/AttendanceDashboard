"""
Attendance Processor — Core logic, no UI dependencies.
"""

import math
import pandas as pd

# defaults 

CONFIG_DEFAULT = {
    "shift_start":                        "08:00",
    "shift_end":                          "15:00",
    "grace_minutes":                      15,
    "tardiness_base":                     "shift_start",
    "tardiness_rounding":                 "daily",
    "tardiness_round_up_to":              15,
    "early_departure_tolerance_minutes":  0,
    "early_departure_round_up_to":        15,
    "overtime_threshold_minutes":         30,
    "overtime_round_down_to":             30,
    "deduct_tardiness_from_overtime":     True,
    "missing_checkin_action":             "ignore",
    "missing_checkout_action":            "ignore",
    "overtime_base":                      "shift_end",
    "absent_status":                      "غائب",
}

JUSTIFICATION_TYPES  = ["غياب مبرر", "تأخير مبرر", "خروج مبكر مبرر"]
BLANK_VALUES         = {"—", "-", "", "nan", "None", "none"}

#  date helper

def normalize_date(val) -> str:
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
    s = str(val).strip()
    if s in BLANK_VALUES:
        return ""
    try:
        if len(s) == 10 and s[4] in ("/", "-"):
            dt = pd.to_datetime(s, errors="coerce", yearfirst=True)
            if not pd.isna(dt):
                return dt.strftime("%Y/%m/%d")
    except Exception:
        pass
    return s

#  time helpers 

def hhmm_to_min(val):
    if not val or str(val).strip() in BLANK_VALUES:
        return None
    try:
        h, m = str(val).strip().split(":")
        return int(h) * 60 + int(m)
    except Exception:
        return None

def min_to_hhmm(minutes):
    if not minutes or minutes <= 0:
        return "00:00"
    h, m = divmod(int(minutes), 60)
    return f"{h:02d}:{m:02d}"

def round_up(n, unit):
    return math.ceil(n / unit) * unit if unit > 1 else int(n)

def round_down(n, unit):
    return math.floor(n / unit) * unit if unit > 1 else int(n)

def is_blank(val):
    return str(val).strip() in BLANK_VALUES

#  loaders 

def load_attendance(file) -> pd.DataFrame:
    df = pd.read_excel(file, dtype=str)
    df["رقم الموظف"] = pd.to_numeric(df["رقم الموظف"], errors="coerce")
    df = df.dropna(subset=["رقم الموظف"])
    df["رقم الموظف"] = df["رقم الموظف"].astype(int)
    df["التاريخ"]    = df["التاريخ"].apply(normalize_date)
    return df.reset_index(drop=True)


def _join_date_columns(df, year_col, month_col, day_col) -> pd.Series:
    def _build(row):
        try:
            y = str(int(float(str(row[year_col])))).zfill(4)
            m = str(int(float(str(row[month_col])))).zfill(2)
            d = str(int(float(str(row[day_col])))).zfill(2)
            return f"{y}/{m}/{d}"
        except Exception:
            return ""
    return df.apply(_build, axis=1)


def _has_split_dates(df) -> bool:
    """Detects split date columns even if headers have extra text like (1-12)."""
    cols = " | ".join(df.columns)
    return "السنة" in cols and ("اليوم" in cols or "اليوم" in " ".join(df.columns))

def _find_col(df, keyword) -> str:
    """Find column name that contains the keyword."""
    for col in df.columns:
        if keyword in col:
            return col
    return None


def load_exceptions(file):
    xl     = pd.ExcelFile(file)
    sheets = xl.sheet_names
    df_ind = xl.parse(sheets[0], dtype=str) if sheets else pd.DataFrame()

    if not df_ind.empty:
        if _has_split_dates(df_ind):
            year_col  = _find_col(df_ind, "السنة")
            month_col = _find_col(df_ind, "الشهر")
            day_col   = _find_col(df_ind, "اليوم")
            df_ind["التاريخ"] = _join_date_columns(df_ind, year_col, month_col, day_col)
        else:
            if len(df_ind.columns) >= 2:
                df_ind["التاريخ"] = df_ind[df_ind.columns[1]].apply(normalize_date)

        # تأكد من وجود عمود النوع
        if "النوع" not in df_ind.columns and len(df_ind.columns) >= 3:
            df_ind = df_ind.rename(columns={df_ind.columns[2]: "النوع"})

    return df_ind

#  lookup builder 

def build_just_map(df_ind: pd.DataFrame) -> dict:
    result = {}
    if df_ind is None or df_ind.empty:
        return result
    for _, row in df_ind.iterrows():
        try:
            emp   = int(float(str(row[df_ind.columns[0]])))
            date  = str(row["التاريخ"]).strip()
            jtype = str(row["النوع"]).strip()
            if emp and date and jtype not in ("nan", ""):
                result.setdefault((emp, date), []).append(jtype)
        except Exception:
            continue
    return result

#  record-level processing 

def process_record(row: dict, just_map: dict, cfg: dict) -> dict:
    shift_start = hhmm_to_min(cfg["shift_start"])
    shift_end   = hhmm_to_min(cfg["shift_end"])
    grace       = cfg["grace_minutes"]

    try:
        emp = int(float(str(row.get("رقم الموظف", 0))))
    except Exception:
        emp = 0

    date   = str(row.get("التاريخ",   "")).strip()
    cin    = str(row.get("أول دخول",  "")).strip()
    cout   = str(row.get("آخر خروج", "")).strip()
    status = str(row.get("الحالة",    "")).strip()

    out = {
        "الحالة_الفعلية":  status,
        "غياب_مبرر":       False,
        "تأخير_خام":       0,
        "تأخير_مقرب":      0,
        "تأخير_مبرر":      False,
        "مبكر_خام":        0,
        "مبكر_مقرب":       0,
        "مبكر_مبرر":       False,
        "اضافي_خام":       0,
        "اضافي_مقرب":      0,
        "اضافي_صافي":      0,
        "مدة_العمل_دقيقة": None,
        "ملاحظة":          "",
    }

    justs = just_map.get((emp, date), [])

    #  غائب 
    if status == cfg["absent_status"]:
        if "غياب مبرر" in justs:
            out["غياب_مبرر"]      = True
            out["الحالة_الفعلية"] = "غياب مبرر"
        else:
            out["الحالة_الفعلية"] = "غياب غير مبرر"
        return out

    #  حاضر 
    cin_missing  = is_blank(cin)
    cout_missing = is_blank(cout)

    if cin_missing:
        if cfg["missing_checkin_action"] == "absent":
            out["الحالة_الفعلية"] = "غياب غير مبرر"
            out["ملاحظة"]         = "حاضر بدون تسجيل دخول — احتُسب غياباً"
        else:
            out["الحالة_الفعلية"] = "حاضر"
            out["ملاحظة"]         = "دخول غير مسجل"
        return out

    cin_m = hhmm_to_min(cin)

    if cout_missing:
        out["ملاحظة"] = "خروج غير مسجل"
        cout_m = hhmm_to_min(cfg["shift_end"]) if cfg["missing_checkout_action"] == "shift_end" else None
    else:
        cout_m = hhmm_to_min(cout)

    #  تأخير 
    late_threshold = shift_start + grace
    if cin_m > late_threshold:
        raw_tard = cin_m - (late_threshold if cfg["tardiness_base"] == "after_grace" else shift_start)
        out["تأخير_خام"] = raw_tard
        if "تأخير مبرر" in justs:
            out["تأخير_مبرر"] = True
        else:
            if cfg["tardiness_rounding"] == "daily":
                out["تأخير_مقرب"] = round_up(raw_tard, cfg["tardiness_round_up_to"])
            else:
                out["تأخير_مقرب"] = raw_tard

    #  مبكر وإضافي 
    if cout_m is not None:
        out["مدة_العمل_دقيقة"] = cout_m - cin_m

        tol = cfg["early_departure_tolerance_minutes"]
        if cout_m < shift_end - tol:
            raw_early = shift_end - cout_m
            out["مبكر_خام"] = raw_early
            if "خروج مبكر مبرر" in justs:
                out["مبكر_مبرر"] = True
            else:
                out["مبكر_مقرب"] = round_up(raw_early, cfg["early_departure_round_up_to"])

        ot_threshold = shift_end + cfg["overtime_threshold_minutes"]
        if cout_m > ot_threshold:
            raw_ot = cout_m - (ot_threshold if cfg["overtime_base"] == "after_threshold" else shift_end)
            out["اضافي_خام"]  = raw_ot
            out["اضافي_مقرب"] = round_down(raw_ot, cfg["overtime_round_down_to"])

    #  إضافي صافي 
    ot = out["اضافي_مقرب"]
    if cfg["deduct_tardiness_from_overtime"] and ot > 0:
        ot = max(0, ot - out["تأخير_مقرب"])
    out["اضافي_صافي"] = ot

    return out

#  pipeline 

def run(df_att, df_ind, cfg):
    just_map = build_just_map(df_ind)
    rows     = [process_record(r, just_map, cfg)
                for r in df_att.to_dict("records")]
    computed = pd.DataFrame(rows)
    detail   = pd.concat([df_att.reset_index(drop=True), computed], axis=1)
    detail["الشهر"]      = detail["التاريخ"].astype(str).str[:7]
    detail["عطلة_رسمية"] = False
    return detail, _summarise(detail, cfg)


def _summarise(detail, cfg):
    working = detail[~detail["عطلة_رسمية"]].copy()
    return (
        working
        .groupby(["رقم الموظف", "اسم الموظف", "المديرية"], as_index=False)
        .apply(_agg_employee, cfg=cfg)
        .reset_index(drop=True)
    )


def _agg_employee(g, cfg):
    present_statuses = {"حاضر", "خروج غير مسجل", "دخول غير مسجل"}
    total   = len(g)
    present = g["الحالة_الفعلية"].isin(present_statuses).sum()
    abs_unj = (g["الحالة_الفعلية"] == "غياب غير مبرر").sum()
    abs_jus = g["غياب_مبرر"].sum()

    raw_tard_total = int(g["تأخير_خام"].sum())
    if cfg["tardiness_rounding"] == "total":
        tard = round_up(raw_tard_total, cfg["tardiness_round_up_to"])
    elif cfg["tardiness_rounding"] == "none":
        tard = raw_tard_total
    else:
        tard = int(g["تأخير_مقرب"].sum())

    early  = int(g["مبكر_مقرب"].sum())
    ot_net = int(g["اضافي_صافي"].sum())
    rate   = round(present / total, 3) if total else 0

    return pd.Series({
        "أيام_عمل_متوقعة":  total,
        "أيام_حضور":        int(present),
        "غياب_غير_مبرر":    int(abs_unj),
        "غياب_مبرر":        int(abs_jus),
        "تأخير_دقيقة":      tard,
        "تأخير_hhmm":       min_to_hhmm(tard),
        "مبكر_دقيقة":       early,
        "مبكر_hhmm":        min_to_hhmm(early),
        "اضافي_صافي_دقيقة": ot_net,
        "اضافي_صافي_hhmm":  min_to_hhmm(ot_net),
        "معدل_الحضور":      rate,
    })
