# -*- coding: utf-8 -*-
"""
sa_ru.py — сезонная корректировка месячных рядов с учётом российского
производственного календаря. Python-версия sa_ru_functions_v2.R.

Считает ТОТ ЖЕ X-13ARIMA-SEATS (Census Bureau), что и R-пакет seasonal:
Python пишет spec-файл, запускает программу x13as, читает результат.
Поэтому результаты совпадают с R-версией.

Главные функции
---------------
sa_ru_batch(data_file, calendar_file, config_file, output_file)
    много рядов из одного Excel; настройки — на листе config отдельного Excel
sa_ru(df, calendar_file, ...)
    один ряд, как seas() в R
sa_ru_make_config(data_file, config_file)
    создать шаблон файла настроек под конкретный файл с данными

Что нужно: Python 3.9+, pandas, numpy, openpyxl; программа X-13 в папке bin/
рядом с этим файлом (x13ashtml / x13as для Mac, x13ashtml.exe / x13as.exe для Windows).

Типы входных рядов (input_type)
-------------------------------
mom_index  — месячный индекс, 103.7 = +3.7% м/м (как в inflcomponents_nonSA.xlsx)
log_level  — накопленный лог-уровень (как в inflserv_nonSA.xlsx и др.)
level      — обычный положительный уровень (ИПП, зарплаты, M2, ...)
Для mom_index и log_level ряд внутри переводится в лог-уровень и сезонится
аддитивно (= мультипликативно для цен, как в методике ЦБ); на выходе — тот же
формат, что на входе.

Тарифы ЖКУ (tariff)
-------------------
none       — ничего (по умолчанию)
schedule   — ряд = сами тарифы (ЖКУ): SA считается по формуле ЦБ — годовой рост
             ряда размазывается ровно по 12 месяцам календарного года. Для
             истории берётся факт из самого ряда, для незакрытого текущего года —
             план с листа tariff_plan файла календаря.
regressor  — ряд = агрегат, внутри которого сидят ЖКУ (услуги целиком, ИПЦ):
             в X-13 добавляется регрессор из графика индексаций (факт из колонки
             tariff_series + план), коэффициент оценивается.
"""

from __future__ import annotations

import glob
import json
import math
import os
import re
import shutil
import subprocess
import tempfile
import warnings
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple, Union

import numpy as np
import pandas as pd

__version__ = "1.0"

SPEC_VERSION = 1              # версия формата спецификаций
SPECS_DIR = "sa_specs"       # папка со спецификациями по умолчанию
_FS_FORBIDDEN = '<>:"/\\|?*'  # запрещённое в именах файлов Windows

# ============================================================================
# 0. Где лежит программа X-13
# ============================================================================

_X13_NAMES_WIN = ["x13ashtml.exe", "x13as_html.exe", "x13as.exe"]
_X13_NAMES_NIX = ["x13ashtml", "x13as_html", "x13as"]


def find_x13(path: Optional[Union[str, Path]] = None) -> Path:
    """Найти исполняемый файл X-13ARIMA-SEATS.

    Порядок поиска: явный `path` (файл или папка) -> папка bin/ рядом с sa_ru.py
    -> переменная окружения X13PATH -> папка R-пакета x13binary -> системный PATH.
    """
    names = _X13_NAMES_WIN if os.name == "nt" else _X13_NAMES_NIX

    def _check_dir(d: Path) -> Optional[Path]:
        for n in names:
            f = d / n
            if f.is_file():
                return f
        return None

    if path is not None:
        p = Path(path)
        if p.is_file():
            return p
        if p.is_dir():
            f = _check_dir(p)
            if f:
                return f
        raise FileNotFoundError(f"X-13 не найден по указанному пути: {path}")

    dirs: List[Path] = [Path(__file__).resolve().parent / "bin"]
    env = os.environ.get("X13PATH")
    if env:
        dirs.append(Path(env))
    home = Path.home()
    if os.name == "nt":
        patterns = [
            str(home / "AppData/Local/R/win-library/*/x13binary/bin"),
            str(home / "Documents/R/win-library/*/x13binary/bin"),
            "C:/Program Files/R/R-*/library/x13binary/bin",
        ]
    else:
        patterns = [
            "/Library/Frameworks/R.framework/Versions/*/Resources/library/x13binary/bin",
            str(home / "Library/R/*/library/x13binary/bin"),
            str(home / "R/*/*/x13binary/bin"),
            "/usr/lib/R/site-library/x13binary/bin",
            "/usr/local/lib/R/site-library/x13binary/bin",
        ]
    for pat in patterns:
        dirs += [Path(p) for p in sorted(glob.glob(pat), reverse=True)]

    for d in dirs:
        f = _check_dir(d)
        if f:
            if os.name != "nt" and not os.access(f, os.X_OK):
                try:
                    os.chmod(f, 0o755)
                except OSError:
                    pass
            return f
    for n in names:
        w = shutil.which(n)
        if w:
            return Path(w)

    raise FileNotFoundError(
        "Не найдена программа X-13ARIMA-SEATS. Положите файл "
        + (" или ".join(names))
        + f" в папку {Path(__file__).resolve().parent / 'bin'} "
        "(скачать: https://www.census.gov/data/software/x13as.html) "
        "или укажите путь аргументом x13_path / переменной окружения X13PATH."
    )


# ============================================================================
# 1. Разбор дат и чисел
# ============================================================================

def _parse_one_date(s) -> pd.Timestamp:
    """Одно значение даты -> первое число месяца (или NaT)."""
    if s is None or (isinstance(s, float) and math.isnan(s)):
        return pd.NaT
    if isinstance(s, pd.Period):
        return s.to_timestamp(how="start")
    if isinstance(s, (pd.Timestamp, datetime, date, np.datetime64)):
        ts = pd.Timestamp(s)
        return pd.NaT if pd.isna(ts) else ts.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    if isinstance(s, (int, float, np.integer, np.floating)):
        # числовой серийник Excel (origin 1899-12-30)
        return (pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(s))).replace(day=1)

    t = str(s).strip()
    if t == "" or t.lower() in ("nan", "nat", "none"):
        return pd.NaT
    m = re.fullmatch(r"(\d{4})[Mm](\d{1,2})", t)                # 2015M1
    if m:
        return pd.Timestamp(year=int(m.group(1)), month=int(m.group(2)), day=1)
    m = re.fullmatch(r"(\d{4})[-/.](\d{1,2})", t)               # 2015-01
    if m:
        return pd.Timestamp(year=int(m.group(1)), month=int(m.group(2)), day=1)
    m = re.fullmatch(r"(\d{1,2})[-/.](\d{4})", t)               # 01.2015
    if m:
        return pd.Timestamp(year=int(m.group(2)), month=int(m.group(1)), day=1)
    try:
        ts = pd.Timestamp(t)                                     # 2015-01-01 и т.п.
        return pd.NaT if pd.isna(ts) else ts.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    except Exception:
        return pd.NaT


def parse_to_month(x) -> pd.DatetimeIndex:
    """Даты в любом формате -> DatetimeIndex первых чисел месяцев.

    Понимает: Timestamp/datetime/date, Excel-серийник, "2015M1", "2015-01",
    "2015/01", "01.2015", "2015-01-01".
    """
    if isinstance(x, (str, int, float, date, datetime, pd.Timestamp, np.datetime64, pd.Period)):
        x = [x]
    s = pd.Series(x) if not isinstance(x, pd.Series) else x
    if pd.api.types.is_datetime64_any_dtype(s):
        return pd.DatetimeIndex(pd.to_datetime(s).dt.to_period("M").dt.to_timestamp())
    return pd.DatetimeIndex([_parse_one_date(v) for v in s])


def parse_numeric(x) -> np.ndarray:
    """Числа, устойчиво к запятой-разделителю ("1,5")."""
    s = pd.Series(x)
    out = pd.to_numeric(s, errors="coerce")
    bad = out.isna() & s.notna()
    if bad.any():
        out[bad] = pd.to_numeric(s[bad].astype(str).str.replace(",", ".", regex=False)
                                 .str.replace(" ", "", regex=False), errors="coerce")
    return out.to_numpy(dtype=float)


def _ym(ts: pd.Timestamp) -> str:
    return f"{ts.year:04d}-{ts.month:02d}"


# ============================================================================
# 2. Проверка и подготовка месячного ряда
# ============================================================================

def prepare_monthly_df(df: pd.DataFrame, date_col: str = "date", value_col: str = "value") -> pd.DataFrame:
    """Вернуть DataFrame(date, value): даты — первые числа месяцев, непрерывный ряд без NA."""
    if date_col not in df.columns:
        raise ValueError(f"Колонка '{date_col}' не найдена.")
    if value_col not in df.columns:
        raise ValueError(f"Колонка '{value_col}' не найдена.")
    out = pd.DataFrame({
        "date": parse_to_month(df[date_col]),
        "value": parse_numeric(df[value_col].values),
    }).sort_values("date").reset_index(drop=True)

    if out["date"].isna().any():
        raise ValueError("Не удалось разобрать часть дат. Проверьте колонку дат.")
    if np.isnan(out["value"]).any():
        bad = out.loc[np.isnan(out["value"]), "date"]
        raise ValueError("В значениях ряда есть пустые ячейки: "
                         + ", ".join(_ym(d) for d in bad[:5]) + (" ..." if len(bad) > 5 else ""))
    if out["date"].duplicated().any():
        raise ValueError("В ряду есть повторяющиеся месяцы.")
    expected = pd.date_range(out["date"].min(), out["date"].max(), freq="MS")
    if len(expected) != len(out):
        missing = expected.difference(pd.DatetimeIndex(out["date"]))
        raise ValueError("Пропущены месяцы: " + ", ".join(_ym(d) for d in missing[:5]))
    return out


# ============================================================================
# 3. Производственный календарь -> месячные центрированные регрессоры
# ============================================================================

def load_calendar(calendar_file: Union[str, Path], sheet: Union[int, str] = 0,
                  date_col: str = "date", workday_col: str = "is_workday",
                  holiday_col: str = "is_holiday", easter_col: str = "easter_effect") -> pd.DataFrame:
    """Прочитать дневной календарь из Excel -> DataFrame(date, is_workday, is_holiday, easter_effect)."""
    daily = pd.read_excel(calendar_file, sheet_name=sheet)
    daily.columns = [str(c).strip().lower() for c in daily.columns]
    date_col, workday_col, holiday_col, easter_col = (c.lower() for c in (date_col, workday_col, holiday_col, easter_col))

    missing = [c for c in (date_col, workday_col, holiday_col) if c not in daily.columns]
    if missing:
        raise ValueError(f"В календаре не найдены колонки: {', '.join(missing)}")
    has_easter = easter_col in daily.columns
    out = pd.DataFrame({
        "date": pd.to_datetime(daily[date_col]),
        "is_workday": daily[workday_col].astype(int),
        "is_holiday": daily[holiday_col].astype(int),
        "easter_effect": daily[easter_col].astype(int) if has_easter else 0,
    }).sort_values("date").reset_index(drop=True)
    out.attrs["has_easter"] = has_easter
    if out["date"].isna().any():
        raise ValueError("В календаре есть даты, которые не удалось прочитать.")
    for c in ("is_workday", "is_holiday", "easter_effect"):
        bad = set(out[c].unique()) - {0, 1}
        if bad:
            raise ValueError(f"Колонка календаря '{c}' должна содержать только 0 и 1, найдено: {bad}")
    return out


def make_ru_calendar(month_dates: pd.DatetimeIndex, daily: pd.DataFrame, include_easter: bool = True,
                     center_start=None, center_end=None) -> pd.DataFrame:
    """Месячные регрессоры из дневного календаря (та же логика центрирования, что в R).

    center_start / center_end — окно, по которому считаются средние для
    центрирования (None = весь календарь).
    """
    month_dates = parse_to_month(month_dates)
    if include_easter and not daily.attrs.get("has_easter", True):
        raise ValueError("include_easter=True, но в календаре нет колонки easter_effect.")

    d = daily.copy()
    d["month_date"] = d["date"].dt.to_period("M").dt.to_timestamp()
    monthly = d.groupby("month_date").agg(
        days_in_month=("date", "count"), workdays=("is_workday", "sum"),
        holidays=("is_holiday", "sum"), easter_days=("easter_effect", "sum")).reset_index()
    monthly["weekends"] = monthly["days_in_month"] - monthly["workdays"] - monthly["holidays"]
    monthly["moy"] = monthly["month_date"].dt.month

    center = monthly
    if center_start is not None:
        center = center[center["month_date"] >= parse_to_month(center_start)[0]]
    if center_end is not None:
        center = center[center["month_date"] <= parse_to_month(center_end)[0]]
    means = center.groupby("moy").agg(m_workdays=("workdays", "mean"), m_holidays=("holidays", "mean"),
                                      m_easter_days=("easter_days", "mean")).reset_index()
    if len(means) < 12:
        warnings.warn("Окно центрирования покрывает не все 12 месяцев; центрирую по всему календарю.")
        means = monthly.groupby("moy").agg(m_workdays=("workdays", "mean"), m_holidays=("holidays", "mean"),
                                           m_easter_days=("easter_days", "mean")).reset_index()

    out = pd.DataFrame({"date": month_dates})
    out["moy"] = out["date"].dt.month
    out = out.merge(monthly.drop(columns="moy"), left_on="date", right_on="month_date", how="left") \
             .drop(columns="month_date").merge(means, on="moy", how="left")
    if out["workdays"].isna().any():
        rng = (monthly["month_date"].min(), monthly["month_date"].max())
        raise ValueError(f"В производственном календаре не хватает месяцев для ряда + горизонта прогноза. "
                         f"Календарь покрывает {_ym(rng[0])} .. {_ym(rng[1])}. "
                         "Продлите календарь или уменьшите forecast_months.")
    out["workdays_c"] = out["workdays"] - out["m_workdays"]
    out["holidays_c"] = out["holidays"] - out["m_holidays"]
    out["easter_c"] = (out["easter_days"] - out["m_easter_days"]) if include_easter else 0.0
    return out[["date", "days_in_month", "workdays", "holidays", "weekends", "easter_days",
                "workdays_c", "holidays_c", "easter_c"]].reset_index(drop=True)


def make_ru_calendar_from_excel(month_dates, calendar_file, sheet=0, date_col="date", workday_col="is_workday",
                                holiday_col="is_holiday", easter_col="easter_effect", include_easter=True,
                                center_start=None, center_end=None) -> pd.DataFrame:
    """Как в R: календарь из Excel -> месячные центрированные регрессоры."""
    daily = load_calendar(calendar_file, sheet, date_col, workday_col, holiday_col, easter_col)
    return make_ru_calendar(month_dates, daily, include_easter, center_start, center_end)


XREG_TYPES = {"workdays": "td", "holidays": "td", "easter": "holiday", "tariff": "holiday"}


def build_xreg(cal_df: pd.DataFrame, mode: str, include_easter: bool) -> Optional[pd.DataFrame]:
    """Матрица календарных регрессоров: none / basic (рабочие дни) / extended (+ праздники), + Пасха."""
    if mode == "none":
        return None
    cols = {"workdays": cal_df["workdays_c"].to_numpy(float)}
    if mode == "extended":
        cols["holidays"] = cal_df["holidays_c"].to_numpy(float)
    if include_easter:
        cols["easter"] = cal_df["easter_c"].to_numpy(float)
    xr = pd.DataFrame(cols, index=pd.DatetimeIndex(cal_df["date"]))
    xr = xr.loc[:, xr.std(axis=0) > 1e-8]          # убрать константные
    return None if xr.shape[1] == 0 else xr


# ============================================================================
# 4. Тарифы ЖКУ: план индексаций, формула ЦБ, регрессор
# ============================================================================

def load_tariff_plan(calendar_file: Union[str, Path], sheet: str = "tariff_plan") -> Optional[pd.DataFrame]:
    """Лист tariff_plan (date, pct) из файла календаря -> DataFrame(date, pct). None, если листа нет."""
    try:
        xl = pd.ExcelFile(calendar_file)
    except Exception:
        return None
    if sheet not in xl.sheet_names:
        return None
    p = pd.read_excel(calendar_file, sheet_name=sheet)
    p.columns = [str(c).strip().lower() for c in p.columns]
    if "date" not in p.columns or "pct" not in p.columns:
        raise ValueError(f"Лист '{sheet}' должен содержать колонки date и pct.")
    p = p[p["date"].notna() & p["pct"].notna()]
    out = pd.DataFrame({"date": parse_to_month(p["date"]), "pct": parse_numeric(p["pct"].values)})
    if out["date"].isna().any():
        raise ValueError(f"Лист '{sheet}': не удалось разобрать часть дат.")
    return out.sort_values("date").reset_index(drop=True)


def tariff_schedule_adjust(dates: pd.DatetimeIndex, mom_index: np.ndarray,
                           plan: Optional[pd.DataFrame]) -> Tuple[np.ndarray, pd.DataFrame, List[str]]:
    """Формула ЦБ для регулируемых тарифов.

    SA-индекс каждого месяца года Y = (годовой рост ряда за Y)^(1/12).
    Для незакрытого последнего года к факту добавляется план (месяцы после
    последней фактической точки) с листа tariff_plan.

    Возвращает (sa_mom_index, таблица по годам, предупреждения).
    """
    warns: List[str] = []
    g = np.log(np.asarray(mom_index, float) / 100.0)
    years = dates.year.to_numpy()
    last = dates[-1]
    totals: Dict[int, float] = {}
    rows = []
    for Y in np.unique(years):
        tot = float(g[years == Y].sum())
        n_fact = int((years == Y).sum())
        plan_add, plan_months = 0.0, []
        if Y == last.year and last.month < 12:
            if plan is not None:
                sel = plan[(plan["date"] > last) & (plan["date"].dt.year == Y)]
                plan_add = float(np.log1p(sel["pct"].to_numpy(float) / 100.0).sum())
                plan_months = [_ym(d) for d in sel["date"]]
            if not plan_months:
                warns.append(f"tariff_plan: нет плана индексаций на {_ym(last)}..{Y}-12 — "
                             f"годовой итог {Y} посчитан только по факту (как будто индексаций больше не будет).")
        totals[Y] = tot + plan_add
        rows.append({"year": Y, "months_fact": n_fact, "growth_fact_pct": (math.exp(tot) - 1) * 100,
                     "plan_months": ", ".join(plan_months), "plan_pct": (math.exp(plan_add) - 1) * 100,
                     "growth_total_pct": (math.exp(tot + plan_add) - 1) * 100,
                     "sa_mom_pct": (math.exp((tot + plan_add) / 12) - 1) * 100})
    sa = np.array([100.0 * math.exp(totals[y] / 12.0) for y in years])
    return sa, pd.DataFrame(rows), warns


def build_tariff_regressor(xreg_dates: pd.DatetimeIndex, actual: Optional[pd.Series],
                           plan: Optional[pd.DataFrame]) -> Tuple[np.ndarray, List[str]]:
    """Регрессор «график индексаций» в накопленном (уровневом) виде.

    Импульс месяца = log(рост тарифов) — факт из `actual` (м/м индекс ЖКУ,
    Series с датами в индексе), а где факта нет — план; центрируется внутри
    календарного года (сумма за год = 0) и накапливается.
    """
    warns: List[str] = []
    n = len(xreg_dates)
    g = np.zeros(n)
    src = np.array(["none"] * n, dtype=object)
    if actual is not None:
        a = actual.copy()
        a.index = parse_to_month(a.index)
        for i, d in enumerate(xreg_dates):
            if d in a.index and not pd.isna(a.loc[d]):
                g[i] = math.log(float(a.loc[d]) / 100.0)
                src[i] = "fact"
    if plan is not None:
        pmap = dict(zip(plan["date"], plan["pct"]))
        for i, d in enumerate(xreg_dates):
            if src[i] == "none" and d in pmap:
                g[i] = math.log1p(float(pmap[d]) / 100.0)
                src[i] = "plan"
    if actual is None and plan is None:
        raise ValueError("tariff='regressor': нужен либо tariff_series (колонка ЖКУ в данных), "
                         "либо лист tariff_plan в файле календаря.")
    years = xreg_dates.year.to_numpy()
    centered = g.copy()
    for Y in np.unique(years):
        m = years == Y
        centered[m] = g[m] - g[m].sum() / m.sum()
    level = np.cumsum(centered)
    if actual is not None:
        last_fact = max((d for d, s in zip(xreg_dates, src) if s == "fact"), default=None)
        if last_fact is not None:
            future_plan = [d for d, s in zip(xreg_dates, src) if s == "plan" and d > last_fact]
            if not future_plan:
                warns.append("tariff_plan: на горизонт прогноза нет плановых индексаций — регрессор после "
                             f"{_ym(last_fact)} считается без индексаций.")
    return level, warns


# ============================================================================
# 5. Запуск X-13ARIMA-SEATS
# ============================================================================

_MONTH_ABBR = {m.lower(): i for i, m in enumerate(
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], start=1)}


def normalize_outlier_name(s: str) -> str:
    """'LS2015.Jul' / 'ls2015.7' -> 'LS2015.7'."""
    m = re.fullmatch(r"([A-Za-z]{2})(\d{4})\.(\w+)", s.strip())
    if not m:
        return s
    typ, yr, mo = m.group(1).upper(), m.group(2), m.group(3)
    moi = int(mo) if mo.isdigit() else _MONTH_ABBR.get(mo[:3].lower())
    return f"{typ}{yr}.{moi}" if moi else s


def fmt_coef(v: float, fix: bool = True) -> str:
    """Значение коэффициента для spec-файла. Суффикс 'f' = X-13 обязан взять его как есть."""
    return f"{float(v):.12f}" + ("f" if fix else "")


def _strip_html(text: str) -> str:
    text = re.sub(r"<[^>]+>", " ", text)
    text = text.replace("&nbsp;", " ").replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")
    return re.sub(r"[ \t]+", " ", text)


class X13Fit:
    """Результат одного запуска X-13."""

    def __init__(self):
        self.ok: bool = False
        self.error: str = ""
        self.method: str = ""                 # seats / x11
        self.udg: Dict[str, str] = {}
        self.reg: pd.DataFrame = pd.DataFrame()      # коэффициенты регрессоров
        self.arima_coef: pd.DataFrame = pd.DataFrame()
        self.tables: Dict[str, pd.Series] = {}       # s11/d11 и др., индекс — даты
        self.spec: str = ""

    # --- удобные извлекатели
    def udg_num(self, *keys) -> float:
        for k in keys:
            if k in self.udg:
                try:
                    return float(self.udg[k].split()[0])
                except ValueError:
                    pass
        return float("nan")

    def qs_pval(self, which: str) -> float:
        v = self.udg.get(which)
        if not v:
            return float("nan")
        parts = v.split()
        try:
            return float(parts[1]) if len(parts) > 1 else float("nan")
        except ValueError:
            return float("nan")

    @property
    def transform(self) -> str:
        t = self.udg.get("aictrans", "")
        if t:
            return "log" if "log" in t.lower() else "none"
        t = self.udg.get("transform", "").lower()
        if "log" in t:
            return "log"
        if "none" in t or "no" in t:
            return "none"
        return t or "none"

    @property
    def arima(self) -> str:
        return self.udg.get("arimamdl", self.udg.get("automdl", "")).strip()

    @property
    def outliers(self) -> List[str]:
        """Выбросы в читаемом виде: 'LS2022.3'."""
        return [normalize_outlier_name(v) for v in self.outliers_x13]

    @property
    def outliers_x13(self) -> List[str]:
        """Выбросы так, как их называет X-13 ('LS2022.Mar') — в этом виде их и пишем обратно."""
        if self.reg.empty:
            return []
        names = self.reg["variable"].astype(str)
        return [v for v in names if re.match(r"^(AO|LS|TC|SO|RP)\d{4}\.", v, re.I)]

    @property
    def has_constant(self) -> bool:
        """Включила ли модель константу (automdl с checkmu)."""
        if self.reg.empty:
            return False
        return bool(self.reg["variable"].astype(str).str.match(r"(?i)^const", na=False).any())

    @property
    def reg_map(self) -> Dict[str, float]:
        """Имя регрессора -> оценка."""
        return {} if self.reg.empty else dict(zip(self.reg["variable"].astype(str),
                                                  self.reg["estimate"].astype(float)))

    def arima_split(self) -> Tuple[List[float], List[float]]:
        """AR- и MA-коэффициенты в том порядке, в каком их выдал X-13."""
        if self.arima_coef.empty:
            return [], []
        ops = self.arima_coef["operator"].astype(str).str.upper()
        est = self.arima_coef["estimate"].astype(float)
        return list(est[ops.str.startswith("AR")]), list(est[ops.str.startswith("MA")])

    @property
    def seats_model(self) -> str:
        """Модель, которой SEATS реально считает фильтры (может отличаться от arima)."""
        return self.udg.get("seatsmdl", "").strip()

    def table(self, name: str) -> Optional[pd.Series]:
        return self.tables.get(name)

    @property
    def sa(self) -> Optional[pd.Series]:
        return self.table("s11" if self.method == "seats" else "d11")

    @property
    def trend(self) -> Optional[pd.Series]:
        return self.table("s12" if self.method == "seats" else "d12")


def _fmt_x13_date(ts: pd.Timestamp) -> str:
    return f"{ts.year}.{ts.month}"


def _wrap_paren_list(items: Sequence[str], indent: str = "    ", width: int = 90) -> str:
    """Список в скобках, разбитый на строки (в spec-файле X-13 строка не длиннее 132 символов)."""
    lines, cur = [], ""
    for it in items:
        if len(cur) + len(it) + 1 > width:
            lines.append(cur)
            cur = ""
        cur += (" " if cur else "") + it
    if cur:
        lines.append(cur)
    return "(" + ("\n" + indent).join(lines) + ")"


def write_spec(transform="none", user_vars: Optional[List[str]] = None, user_types: Optional[List[str]] = None,
               fixed_vars: Optional[List[str]] = None, auto_outlier=True, outlier_types="all",
               outlier_critical=None, outlier_span=None, arima_model=None, forecast=36,
               decomposition="seats", fixed_b: Optional[List[str]] = None,
               arima_ar: Optional[List[str]] = None, arima_ma: Optional[List[str]] = None) -> str:
    """Собрать текст spec-файла X-13 (по образцу того, что пишет R-пакет seasonal).

    fixed_b / arima_ar / arima_ma — уже отформатированные значения коэффициентов;
    суффикс "f" означает «зафиксировать» (см. fmt_coef).
    """
    L = ["series{", '  title = "series"', '  file = "iofile.dta"', '  format = "datevalue"', "  period = 12", "}", ""]
    L += ["transform{", f"  function = {transform}"]
    if transform == "auto":
        L.append("  print = aictransform")
    L += ["}", ""]

    reg_lines = []
    if fixed_vars:
        reg_lines.append("  variables = " + _wrap_paren_list(fixed_vars))
    if user_vars:
        reg_lines.append("  user = " + _wrap_paren_list(user_vars))
        types = user_types or ["td"] * len(user_vars)
        reg_lines.append("  usertype = (" + " ".join(types) + ")")
        reg_lines += ['  file = "iofile_xreg.dta"', '  format = "datevalue"']
    if fixed_b:
        reg_lines.append("  b = " + _wrap_paren_list(fixed_b))
    if reg_lines:
        L += ["regression{"] + reg_lines + ["}", ""]

    if auto_outlier:
        ot = outlier_types
        if isinstance(ot, (list, tuple)):
            ot = "(" + " ".join(ot) + ")"
        elif isinstance(ot, str) and "," in ot:
            ot = "(" + " ".join(p.strip() for p in ot.split(",")) + ")"
        L += ["outlier{", f"  types = {ot}"]
        if outlier_critical is not None:
            L.append(f"  critical = {float(outlier_critical)}")
        if outlier_span:
            L.append(f"  span = ({outlier_span})")
        L += ["}", ""]

    if arima_model:
        arima_lines = ["arima{", f"  model = {arima_model}"]
        if arima_ar:
            arima_lines.append("  ar = " + _wrap_paren_list(arima_ar))
        if arima_ma:
            arima_lines.append("  ma = " + _wrap_paren_list(arima_ma))
        L += arima_lines + ["}", ""]
    else:
        L += ["automdl{", "  print = bestfivemdl", "}", ""]

    L += ["forecast{", f"  maxlead = {int(forecast)}", "}", ""]
    L += ["estimate{", "  save = (model estimates residuals)", "}", ""]
    L += ["spectrum{", "  print = qs", "}", ""]
    if decomposition == "seats":
        L += ["seats{", "  noadmiss = yes", "  save = (s10 s11 s12 s13 s16)", "}", ""]
    else:
        L += ["x11{", "  save = (d10 d11 d12 d13 d16)", "}", ""]
    return "\n".join(L)


def _write_datevalue(path: Path, dates: pd.DatetimeIndex, values: np.ndarray) -> None:
    vals = np.atleast_2d(np.asarray(values, float))
    if vals.shape[0] != len(dates):
        vals = vals.T
    with open(path, "w") as f:
        for d, row in zip(dates, vals):
            f.write(f"{d.year} {d.month} " + " ".join(repr(float(v)) for v in row) + "\n")


def _read_table(path: Path) -> Optional[pd.Series]:
    if not path.exists():
        return None
    dates, vals = [], []
    with open(path) as f:
        for line in f.readlines()[2:]:
            parts = line.split()
            if len(parts) < 2 or not parts[0].isdigit():
                continue
            ym = parts[0]
            dates.append(pd.Timestamp(year=int(ym[:4]), month=int(ym[4:6]), day=1))
            vals.append(float(parts[1]))
    if not dates:
        return None
    s = pd.Series(vals, index=pd.DatetimeIndex(dates))
    s[s == -999] = np.nan
    return s


def _read_udg(path: Path) -> Dict[str, str]:
    out: Dict[str, str] = {}
    if not path.exists():
        return out
    with open(path, errors="replace") as f:
        for line in f:
            if ":" in line:
                k, v = line.split(":", 1)
                out[k.strip()] = v.strip()
    return out


def _read_est(path: Path) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Файл .est -> (регрессоры: group, variable, estimate, se, t, pval; ARIMA-коэффициенты)."""
    reg = pd.DataFrame(columns=["group", "variable", "estimate", "se", "t", "pval"])
    ar = pd.DataFrame(columns=["operator", "factor", "lag", "estimate", "se"])
    if not path.exists():
        return reg, ar
    lines = open(path, errors="replace").read().splitlines()

    def _block(header: str) -> List[List[str]]:
        try:
            i = next(k for k, l in enumerate(lines) if l.strip().startswith(header))
        except StopIteration:
            return []
        rows = []
        for l in lines[i + 3:]:            # заголовок, шапка, черта
            if l.startswith("$") or not l.strip():
                break
            rows.append(l.split("\t"))
        return rows

    rr = []
    for p in _block("$regression$estimates"):
        if len(p) >= 4:
            est, se = float(p[2]), float(p[3]) if p[3].strip() else float("nan")
            t = est / se if se and not math.isnan(se) and se != 0 else float("nan")
            pval = math.erfc(abs(t) / math.sqrt(2)) if not math.isnan(t) else float("nan")
            rr.append({"group": p[0].strip(), "variable": p[1].strip(), "estimate": est, "se": se, "t": t, "pval": pval})
    if rr:
        reg = pd.DataFrame(rr)
    aa = []
    for p in _block("$arima$estimates"):
        if len(p) >= 6:
            aa.append({"operator": p[0].strip(), "factor": p[1].strip(), "lag": p[3].strip(),
                       "estimate": float(p[4]), "se": float(p[5]) if p[5].strip() else float("nan")})
    if aa:
        ar = pd.DataFrame(aa)
    return reg, ar


def _read_error(workdir: Path) -> str:
    msgs = []
    for name in ("iofile.err", "iofile_err.html"):
        p = workdir / name
        if p.exists():
            txt = open(p, errors="replace").read()
            if name.endswith(".html"):
                txt = _strip_html(txt)
            for line in txt.splitlines():
                if "ERROR" in line.upper():
                    msgs.append(line.strip())
    return "; ".join(dict.fromkeys(msgs)) if msgs else ""


def x13_fit(y: np.ndarray, dates: pd.DatetimeIndex, xreg: Optional[pd.DataFrame] = None,
            xreg_types: Optional[Dict[str, str]] = None, transform: str = "none",
            auto_outlier: bool = True, outlier_types="all", outlier_critical=None, outlier_span=None,
            fixed_outliers: Optional[List[str]] = None, arima_model: Optional[str] = None,
            forecast: int = 36, decomposition: str = "seats",
            fixed_b: Optional[List[str]] = None, arima_ar: Optional[List[str]] = None,
            arima_ma: Optional[List[str]] = None,
            x13_path=None, keep_dir: Optional[Union[str, Path]] = None) -> X13Fit:
    """Один запуск X-13: записать файлы -> запустить -> прочитать результат."""
    exe = find_x13(x13_path)
    fit = X13Fit()
    fit.method = decomposition
    workdir = Path(tempfile.mkdtemp(prefix="sa_ru_"))
    try:
        _write_datevalue(workdir / "iofile.dta", dates, np.asarray(y, float))
        user_vars, user_types = None, None
        if xreg is not None and xreg.shape[1] > 0:
            user_vars = list(xreg.columns)
            user_types = [(xreg_types or {}).get(c, "td") for c in user_vars]
            _write_datevalue(workdir / "iofile_xreg.dta", pd.DatetimeIndex(xreg.index), xreg.to_numpy(float))
        fit.spec = write_spec(transform=transform, user_vars=user_vars, user_types=user_types,
                              fixed_vars=fixed_outliers, auto_outlier=auto_outlier, outlier_types=outlier_types,
                              outlier_critical=outlier_critical, outlier_span=outlier_span,
                              arima_model=arima_model, forecast=forecast, decomposition=decomposition,
                              fixed_b=fixed_b, arima_ar=arima_ar, arima_ma=arima_ma)
        (workdir / "iofile.spc").write_text(fit.spec)

        try:
            proc = subprocess.run([str(exe), "iofile", "-n", "-s"], cwd=str(workdir),
                                  capture_output=True, text=True, errors="replace", timeout=600)
            stdout = (proc.stdout or "") + (proc.stderr or "")
        except Exception as e:                       # noqa: BLE001
            fit.error = f"не удалось запустить X-13 ({exe}): {e}"
            return fit

        fit.udg = _read_udg(workdir / "iofile.udg")
        fit.reg, fit.arima_coef = _read_est(workdir / "iofile.est")
        prefix = "s" if decomposition == "seats" else "d"
        for code in ("10", "11", "12", "13", "16"):
            t = _read_table(workdir / f"iofile.{prefix}{code}")
            if t is not None:
                fit.tables[f"{prefix}{code}"] = t

        if fit.sa is None:
            err = _read_error(workdir)
            if not err:
                err = "; ".join(l.strip() for l in stdout.splitlines() if "ERROR" in l.upper()) \
                      or f"X-13 не выдал таблицу {prefix}11 ({decomposition})"
            fit.error = err
            fit.ok = False
        else:
            fit.ok = True
        if keep_dir:
            shutil.copytree(workdir, Path(keep_dir), dirs_exist_ok=True)
    finally:
        shutil.rmtree(workdir, ignore_errors=True)
    return fit


def fit_with_fallback(method: str = "prefer_seats", **kw) -> X13Fit:
    """prefer_seats: SEATS, при неудаче X11; seats: только SEATS; x11: только X11."""
    order = {"prefer_seats": ["seats", "x11"], "seats": ["seats"], "x11": ["x11"]}[method]
    last = None
    for decomp in order:
        fit = x13_fit(decomposition=decomp, **kw)
        if fit.ok:
            return fit
        last = fit
    return last


def calendar_significance(fit: X13Fit, xreg_names: Sequence[str], alpha: float = 0.05) -> Tuple[int, int]:
    if fit.reg.empty or not xreg_names:
        return 0, len(xreg_names or [])
    sub = fit.reg[fit.reg["variable"].isin(list(xreg_names))]
    return int((sub["pval"] < alpha).sum()), int(len(sub))


def summarise_fit(fit: X13Fit, model_name: str, xreg_names: Sequence[str]) -> Dict[str, Any]:
    n_sig, n_tot = calendar_significance(fit, xreg_names)
    return {"model": model_name, "method": fit.method, "transform": fit.transform, "arima": fit.arima,
            "n_outliers": len(fit.outliers), "aicc": fit.udg_num("aicc"), "aic": fit.udg_num("aic"),
            "bic": fit.udg_num("bic"), "cal_signif": n_sig, "cal_total": n_tot,
            "qs_resid_pval": fit.qs_pval("qssadj")}


# ============================================================================
# 6. Преобразование входа/выхода (mom_index / log_level / level)
# ============================================================================

INPUT_TYPES = ("mom_index", "log_level", "level")


def _to_internal(values: np.ndarray, input_type: str) -> np.ndarray:
    """Ряд для X-13: mom_index -> накопленный лог-уровень (с нуля), log_level/level -> как есть."""
    v = np.asarray(values, float)
    med = float(np.nanmedian(v))
    if input_type == "mom_index":
        if (v <= 0).any() or not (70 <= med <= 150):
            raise ValueError(
                f"input_type='mom_index' ожидает месячные индексы около 100 (103.7 = +3.7%), а в ряду "
                f"медиана {med:.4g}. Похоже, это накопленный лог-уровень или обычный уровень — "
                "поставьте input_type = log_level или level (лист config, колонка input_type)."
            )
        return np.cumsum(np.log(v / 100.0))
    if input_type == "log_level" and abs(med) > 20:
        warnings.warn(f"input_type='log_level', но медиана ряда {med:.4g} — для лог-уровня это необычно. "
                      "Проверьте, не индексы ли это (mom_index) или уровни (level).")
    return v.copy()


def _mom_from_loglevel(y: np.ndarray, prev0: float = 0.0) -> np.ndarray:
    """Лог-уровень -> м/м индекс.

    prev0 — условный уровень месяца ПЕРЕД выборкой. Для исходного ряда это 0
    (тогда первый м/м просто воспроизводится). Для сезонно скорректированного
    ряда там должен стоять -s(0), где s(0) — сезонный фактор того месяца:
    иначе первая точка окажется скорректирована на УРОВЕНЬ сезонного фактора,
    а все остальные — на его ИЗМЕНЕНИЕ (см. _seas0_estimate).
    """
    y = np.asarray(y, float)
    prev = np.concatenate([[float(prev0)], y[:-1]])
    return 100.0 * np.exp(y - prev)


def _seas0_estimate(fit, dates: pd.DatetimeIndex) -> float:
    """Сезонный фактор месяца, предшествующего выборке (его в данных нет).

    Берётся тот же календарный месяц следующего года: сезонные факторы меняются
    медленно, поэтому это хорошее приближение. Нужно только для input_type =
    'mom_index', где первую точку иначе не с чем сравнивать.
    """
    if fit is None or len(dates) < 13:
        return 0.0
    tbl = fit.table("s10" if fit.method == "seats" else "d10")
    if tbl is None:
        return 0.0
    val = tbl.reindex([dates[0] + pd.DateOffset(months=11)]).iloc[0]
    return 0.0 if pd.isna(val) else float(val)



def _to_output(input_type: str, v: np.ndarray, adj_y: np.ndarray, trend_y: np.ndarray,
               transform: str, seas0: float = 0.0) -> Dict[str, Any]:
    """Внутренний (лог-)ряд -> формат входа: original, adjusted, trend, seasonal_factor.

    seas0 — сезонный фактор месяца перед выборкой (см. _seas0_estimate); нужен,
    чтобы первая точка м/м-ряда считалась так же, как все остальные.
    """
    if input_type == "mom_index":
        adjusted = _mom_from_loglevel(adj_y, prev0=-float(seas0))
        if np.isnan(trend_y).all():
            trend_out = trend_y
        else:
            trend_out = _mom_from_loglevel(trend_y)
            trend_out[0] = np.nan      # уровня тренда перед выборкой нет -> м/м первой точки не считается
        return {"original": v, "adjusted": adjusted, "trend": trend_out,
                "seasonal_factor": v / adjusted, "factor_type": "multiplicative"}
    if input_type == "log_level":
        return {"original": v, "adjusted": adj_y, "trend": trend_y,
                "seasonal_factor": v - adj_y, "factor_type": "additive"}
    if transform == "log":
        return {"original": v, "adjusted": adj_y, "trend": trend_y,
                "seasonal_factor": np.where(adj_y != 0, v / adj_y, np.nan), "factor_type": "multiplicative"}
    return {"original": v, "adjusted": adj_y, "trend": trend_y,
            "seasonal_factor": v - adj_y, "factor_type": "additive"}


def _make_xreg(xreg_dates, daily, mode: str, easter: bool, center_start, center_end,
               tariff: str = "none", tariff_series=None, tariff_plan=None):
    """Матрица регрессоров по готовому описанию дизайна (используется при apply)."""
    cal = make_ru_calendar(xreg_dates, daily, easter, center_start, center_end)
    xr = build_xreg(cal, mode, easter)
    if tariff == "regressor":
        reg, _ = build_tariff_regressor(xreg_dates, tariff_series, tariff_plan)
        t = pd.DataFrame({"tariff": reg}, index=xreg_dates)
        xr = t if xr is None else xr.join(t)
    return xr


# ============================================================================
# 7. Главная функция — один ряд
# ============================================================================

def sa_ru(df: pd.DataFrame, calendar_file: Union[str, Path], *,
          date_col: str = "date", value_col: str = "value",
          input_type: str = "mom_index",
          calendar_sheet: Union[int, str] = 0, calendar_date_col: str = "date",
          calendar_workday_col: str = "is_workday", calendar_holiday_col: str = "is_holiday",
          calendar_easter_col: str = "easter_effect",
          calendar_mode: str = "basic", include_easter: Union[bool, str] = False,
          use_outliers: bool = True, outlier_types="all", outlier_critical=None,
          transform: str = "none", method: str = "prefer_seats",
          forecast_months: int = 36, seasonality_alpha: float = 0.05,
          center_start=None, center_end=None,
          tariff: str = "none", tariff_series: Optional[pd.Series] = None,
          tariff_plan: Optional[pd.DataFrame] = None, tariff_plan_sheet: str = "tariff_plan",
          series_name: str = "", verbose: bool = True,
          x13_path=None, _calendar_daily: Optional[pd.DataFrame] = None) -> Dict[str, Any]:
    """Сезонная корректировка одного месячного ряда (аналог sa_ru() в R).

    df            — таблица с колонками date_col и value_col
    calendar_file — russia_calendar.xlsx (лист с календарём + лист tariff_plan)
    input_type    — mom_index / log_level / level (см. шапку файла)
    calendar_mode — none / basic / extended / auto (auto = выбор по AICc)
    transform     — none / auto / log (для mom_index и log_level всегда none)
    method        — prefer_seats / seats / x11
    tariff        — none / schedule / regressor (см. шапку файла);
                    для regressor: tariff_series — м/м индекс ЖКУ (Series, индекс — даты)
    Возвращает dict; главное — res["data"] с колонками date, original, adjusted,
    trend, seasonal_factor (в формате входа).
    """
    say = (lambda *a: print(*a)) if verbose else (lambda *a: None)
    warns: List[str] = []

    for nm, val, allowed in (("input_type", input_type, INPUT_TYPES),
                             ("calendar_mode", calendar_mode, ("none", "basic", "extended", "auto")),
                             ("transform", transform, ("none", "auto", "log")),
                             ("method", method, ("prefer_seats", "seats", "x11")),
                             ("tariff", tariff, ("none", "schedule", "regressor"))):
        if val not in allowed:
            raise ValueError(f"{nm} должен быть одним из {allowed}, получено '{val}'.")

    # --- 1. данные
    dat = prepare_monthly_df(df, date_col, value_col)
    dates = pd.DatetimeIndex(dat["date"])
    v = dat["value"].to_numpy(float)
    n = len(dat)

    # --- 2. внутренний ряд и преобразование
    if input_type in ("mom_index", "log_level"):
        if transform != "none":
            warns.append(f"transform='{transform}' игнорируется для input_type='{input_type}' "
                         "(ряд уже в логах, используется 'none').")
        transform = "none"
    else:
        if transform == "log" and (v <= 0).any():
            raise ValueError("transform='log', но в ряду есть значения <= 0. Используйте 'none'.")
        if transform == "auto" and (v <= 0).any():
            say("В ряду есть значения <= 0 -> для auto принудительно ставим transform = 'none'.")
            transform = "none"
    y = _to_internal(v, input_type)

    # --- 3. тарифы: план
    if tariff != "none" and tariff_plan is None:
        tariff_plan = load_tariff_plan(calendar_file, tariff_plan_sheet)

    # ==========================================================================
    # ТАРИФНЫЙ РЕЖИМ schedule: формула ЦБ, без X-13
    # ==========================================================================
    if tariff == "schedule":
        if input_type == "level":
            mom = np.concatenate([[100.0], 100.0 * v[1:] / v[:-1]])
        else:
            mom = _mom_from_loglevel(y)
        sa_mom, by_year, w = tariff_schedule_adjust(dates, mom, tariff_plan)
        warns += w
        for msg in w:
            say("  ! " + msg)
        if input_type == "mom_index":
            original, adjusted = v, sa_mom
        elif input_type == "log_level":
            original, adjusted = v, np.cumsum(np.log(sa_mom / 100.0))
        else:
            original, adjusted = v, v[0] * np.cumprod(np.concatenate([[1.0], sa_mom[1:] / 100.0]))
        seas = original / adjusted if input_type != "log_level" else original - adjusted
        data = pd.DataFrame({"date": dates, "original": original, "adjusted": adjusted,
                             "trend": np.nan, "seasonal_factor": seas,
                             "factor_type": "multiplicative" if input_type != "log_level" else "additive"})
        say(f"Тарифный режим (формула ЦБ): годовой рост размазан по 12 месяцам; лет: {len(by_year)}")
        return {"data": data, "chosen_model": "tariff_schedule", "include_easter": False,
                "transform": "none",
                "decomposition_method": "schedule", "arima": "", "outliers": [], "outliers_x13": [],
                "xreg_names": [],
                "seasonality_detected": True, "qs_orig_pval": np.nan, "qs_orig_evadj_pval": np.nan,
                "comparison": pd.DataFrame(), "coefficients": pd.DataFrame(), "udg": {},
                "calendar_monthly": None, "tariff_by_year": by_year, "warnings": warns,
                "input_type": input_type, "tariff": tariff, "best_fit": None}

    # --- 4. календарь на горизонт ряд + прогноз
    xreg_dates = pd.date_range(dates.min(), dates.max() + pd.DateOffset(months=forecast_months), freq="MS")
    if center_start is None:
        center_start = dates.min()
    if center_end is None:
        center_end = dates.max()
    daily = _calendar_daily if _calendar_daily is not None else load_calendar(
        calendar_file, calendar_sheet, calendar_date_col, calendar_workday_col, calendar_holiday_col, calendar_easter_col)
    # Пасха: True / False / "auto" (перебираем оба варианта и выбираем по AICc)
    easter_auto = isinstance(include_easter, str) and include_easter.strip().lower() == "auto"
    if easter_auto and not daily.attrs.get("has_easter", True):
        warns.append("include_easter='auto', но в календаре нет колонки easter_effect -> Пасха не рассматривается.")
        easter_auto, include_easter = False, False
    cal_df = make_ru_calendar(xreg_dates, daily, easter_auto or include_easter is True, center_start, center_end)

    tariff_reg = None
    if tariff == "regressor":
        tariff_reg, w = build_tariff_regressor(xreg_dates, tariff_series, tariff_plan)
        warns += w
        for msg in w:
            say("  ! " + msg)

    def xreg_for(mode: str, easter: bool) -> Optional[pd.DataFrame]:
        xr = build_xreg(cal_df, mode, easter)
        if tariff_reg is not None:
            t = pd.DataFrame({"tariff": tariff_reg}, index=xreg_dates)
            xr = t if xr is None else xr.join(t)
        return xr

    def label_of(mode: str, easter: bool) -> str:
        return mode + ("+easter" if easter else "")

    # список кандидатов: (режим календаря, Пасха)
    modes = ("none", "basic", "extended") if calendar_mode == "auto" else (calendar_mode,)
    cand_list: List[Tuple[str, bool]] = []
    for m in modes:
        if m == "none":
            cand_list.append((m, False))            # none = вообще без регрессоров, включая Пасху
        elif easter_auto:
            cand_list += [(m, False), (m, True)]
        else:
            cand_list.append((m, bool(include_easter)))

    xreg_types = XREG_TYPES
    names_of = lambda xr: [] if xr is None else list(xr.columns)   # noqa: E731

    common = dict(y=y, dates=dates, xreg_types=xreg_types, forecast=forecast_months, x13_path=x13_path)

    # ==========================================================================
    # РЕЖИМ AUTO: фиксируем transform + выбросы на basic, сравниваем none/basic/extended
    # ==========================================================================
    attempts: List[Dict[str, Any]] = []
    if len(cand_list) > 1:
        say(f"AUTO: {len(cand_list)} кандидатов ({', '.join(label_of(*c) for c in cand_list)}); "
            "опорная модель фиксирует преобразование и выбросы...")
        ref = None
        for ref_mode, ref_eas in [("basic", False), ("none", False), ("extended", False)]:
            ref = fit_with_fallback("x11", xreg=xreg_for(ref_mode, ref_eas), transform=transform,
                                    auto_outlier=use_outliers, outlier_types=outlier_types,
                                    outlier_critical=outlier_critical, **common)
            if ref.ok:
                break
            say(f"  опорная на {ref_mode} не вышла, пробую дальше ...")
        if ref is None or not ref.ok:
            raise RuntimeError("Не удалось оценить опорную модель. Ошибка X-13: " + (ref.error if ref else ""))
        tf_fixed = ref.transform if transform == "auto" else transform
        outliers_fixed = ref.outliers_x13 if use_outliers else []
        say(f"  зафиксировано: преобразование = {tf_fixed} | выбросов = {len(outliers_fixed)}"
            + (f" ({', '.join(ref.outliers)})" if outliers_fixed else ""))

        candidates: Dict[str, X13Fit] = {}
        cand_design: Dict[str, Tuple[str, bool]] = {}
        for mode_c, eas_c in cand_list:
            nm = label_of(mode_c, eas_c)
            cand_design[nm] = (mode_c, eas_c)
            say(f"  оцениваю кандидата: {nm}")
            xr = xreg_for(mode_c, eas_c)
            fit = fit_with_fallback(method, xreg=xr, transform=tf_fixed, auto_outlier=False,
                                    fixed_outliers=outliers_fixed, **common)
            if fit.ok:
                candidates[nm] = fit
                row = summarise_fit(fit, nm, names_of(xr))
            else:
                row = {"model": nm, "method": None, "transform": tf_fixed, "arima": None,
                       "n_outliers": len(outliers_fixed), "aicc": np.nan, "aic": np.nan, "bic": np.nan,
                       "cal_signif": np.nan, "cal_total": len(names_of(xr)), "qs_resid_pval": np.nan,
                       "error": fit.error}
            row["calendar"], row["easter"] = mode_c, eas_c
            attempts.append(row)
        comparison = pd.DataFrame(attempts)
        if not candidates:
            hint = (" Похоже, SEATS не строит разложение для этого ряда — попробуйте method='x11' или 'prefer_seats'."
                    if method == "seats" else "")
            raise RuntimeError(f"Ни один кандидат не оценился (method='{method}')." + hint)
        valid = comparison[comparison["model"].isin(candidates) & comparison["aicc"].notna()]
        if valid.empty:
            raise RuntimeError("Ни у одного кандидата нет AICc для сравнения.")
        best_label = str(valid.loc[valid["aicc"].idxmin(), "model"])
        best_fit = candidates[best_label]
        best_name, best_easter = cand_design[best_label]
        comparison["chosen"] = comparison["model"] == best_label
        say(f"Выбранная спецификация (min AICc): {best_label}")
        if verbose:
            print(comparison.to_string(index=False))
    else:
        # ======================================================================
        # РУЧНОЙ РЕЖИМ: одна спецификация, авто-выбросы
        # ======================================================================
        best_name, best_easter = cand_list[0]
        say(f"Ручной режим, календарь = {label_of(best_name, best_easter)}")
        xr = xreg_for(best_name, best_easter)
        best_fit = fit_with_fallback(method, xreg=xr, transform=transform, auto_outlier=use_outliers,
                                     outlier_types=outlier_types, outlier_critical=outlier_critical, **common)
        if not best_fit.ok:
            raise RuntimeError(f"Не удалось оценить модель (method='{method}'). Ошибка X-13: {best_fit.error}")
        comparison = pd.DataFrame([summarise_fit(best_fit, label_of(best_name, best_easter), names_of(xr))])
        comparison["calendar"], comparison["easter"], comparison["chosen"] = best_name, best_easter, True
        if verbose:
            print(comparison.to_string(index=False))

    best_xreg = xreg_for(best_name, best_easter)

    # --- 5. есть ли сезонность вообще (QS на исходном ряду; главный — с поправкой на выбросы)
    qs_evadj_p = best_fit.qs_pval("qsorievadj")
    qs_ori_p = best_fit.qs_pval("qsori")
    cands = [p for p in (qs_evadj_p, qs_ori_p) if not math.isnan(p)]
    qs_guard = min(cands) if cands else float("nan")
    seasonality_detected = math.isnan(qs_guard) or qs_guard < seasonality_alpha

    tf_final = best_fit.transform
    if not seasonality_detected:
        say(f"Идентифицируемая сезонность не обнаружена (QS p = {qs_guard:.4f}) -> ряд возвращается КАК ЕСТЬ.")
        adj_y = y.copy()
        trend_y = np.full(n, np.nan)
    else:
        adj_y = best_fit.sa.reindex(dates).to_numpy(float)
        tr = best_fit.trend
        trend_y = tr.reindex(dates).to_numpy(float) if tr is not None else np.full(n, np.nan)

    # --- 6. обратно в формат входа
    seas0 = _seas0_estimate(best_fit, dates) if seasonality_detected else 0.0
    data = pd.DataFrame({"date": dates, **_to_output(input_type, v, adj_y, trend_y, tf_final, seas0)})

    return {"data": data, "chosen_model": best_name, "include_easter": bool(best_easter),
            "transform": tf_final,
            "decomposition_method": best_fit.method, "arima": best_fit.arima, "outliers": best_fit.outliers,
            "outliers_x13": best_fit.outliers_x13, "xreg_names": names_of(best_xreg),
            "seasonality_detected": seasonality_detected, "qs_orig_pval": qs_ori_p,
            "qs_orig_evadj_pval": qs_evadj_p, "comparison": comparison, "coefficients": best_fit.reg,
            "arima_coefficients": best_fit.arima_coef, "udg": best_fit.udg, "calendar_monthly": cal_df,
            "tariff_by_year": None, "warnings": warns, "input_type": input_type, "tariff": tariff,
            "center_start": center_start, "center_end": center_end,
            "xreg": best_xreg, "best_fit": best_fit}


# ============================================================================
# 8. Настройки для многих рядов (лист config)
# ============================================================================

DEFAULT_SETTINGS: Dict[str, Any] = {
    "skip": False,
    "input_type": "mom_index",
    "calendar_mode": "basic",
    "include_easter": False,
    "freeze": False,
    "transform": "none",
    "use_outliers": True,
    "outlier_types": "all",
    "outlier_critical": None,
    "method": "prefer_seats",
    "tariff": "none",
    "tariff_series": None,
    "forecast_months": 36,
    "seasonality_alpha": 0.05,
}

CONFIG_COLUMNS = ["series_id"] + list(DEFAULT_SETTINGS) + ["comment"]

_ALLOWED = {
    "input_type": INPUT_TYPES,
    "calendar_mode": ("none", "basic", "extended", "auto"),
    "transform": ("none", "auto", "log"),
    "method": ("prefer_seats", "seats", "x11"),
    "tariff": ("none", "schedule", "regressor"),
}
_BOOL_COLS = ("skip", "use_outliers", "freeze")
_BOOL_OR_AUTO_COLS = ("include_easter",)
_NUM_COLS = ("outlier_critical", "forecast_months", "seasonality_alpha")

CONFIG_HELP = {
    "series_id": "имя ряда = заголовок колонки в файле с данными; строка DEFAULT — настройки для всех",
    "skip": "TRUE — ряд не сезонить (в выходной файл попадёт как есть)",
    "input_type": "mom_index (103.7 = +3.7% м/м) / log_level (накопленный лог-уровень) / level (обычный уровень)",
    "calendar_mode": "none / basic (рабочие дни) / extended (рабочие дни + праздники) / auto (выбор по AICc)",
    "include_easter": "TRUE / FALSE / auto — регрессор Пасхи (auto = выбрать по AICc)",
    "freeze": "TRUE / FALSE — считать по замороженной спецификации из папки sa_specs (нужен прогон идентификации)",
    "transform": "none / auto / log — только для input_type = level (для индексов всегда none)",
    "use_outliers": "TRUE / FALSE — автопоиск выбросов AO/LS/TC",
    "outlier_types": "all или список через запятую: ao,ls,tc",
    "outlier_critical": "критическое t для выбросов; пусто = по умолчанию X-13",
    "method": "prefer_seats (SEATS, при неудаче X11) / seats / x11",
    "tariff": "none / schedule (ряд = тарифы ЖКУ, формула ЦБ) / regressor (ЖКУ внутри агрегата)",
    "tariff_series": "для tariff = regressor: имя колонки с м/м индексом ЖКУ в файле данных",
    "forecast_months": "горизонт прогноза X-13, месяцев",
    "seasonality_alpha": "уровень значимости QS-теста на сезонность",
    "comment": "любые пометки, кодом не читается",
}


def _to_bool(v, col: str, row: str) -> bool:
    """Ячейка -> TRUE/FALSE."""
    if isinstance(v, (bool, np.bool_)):
        return bool(v)
    if isinstance(v, (int, float, np.integer, np.floating)) and not pd.isna(v):
        return bool(int(v))
    s = str(v).strip().lower()
    if s in ("true", "1", "yes", "да", "y", "истина"):
        return True
    if s in ("false", "0", "no", "нет", "n", "ложь"):
        return False
    raise ValueError(f"config: строка '{row}', колонка '{col}': ожидается TRUE/FALSE, получено '{v}'.")


def _clean_value(col: str, v, row: str):
    """Одна ячейка листа config -> значение настройки (None = пусто)."""
    if v is None or (isinstance(v, float) and math.isnan(v)) or (isinstance(v, str) and v.strip() == ""):
        return None
    if col in _BOOL_OR_AUTO_COLS:
        if isinstance(v, str) and v.strip().lower() == "auto":
            return "auto"
        return _to_bool(v, col, row)
    if col in _BOOL_COLS:
        return _to_bool(v, col, row)
    if col in _NUM_COLS:
        try:
            x = float(str(v).replace(",", "."))
        except ValueError:
            raise ValueError(f"config: строка '{row}', колонка '{col}': ожидается число, получено '{v}'.")
        return int(x) if col == "forecast_months" else x
    s = str(v).strip()
    if col in _ALLOWED:
        s = s.lower()
        if s not in _ALLOWED[col]:
            raise ValueError(f"config: строка '{row}', колонка '{col}': допустимо "
                             f"{' / '.join(_ALLOWED[col])}, получено '{v}'.")
    return s


def read_config(config_file: Union[str, Path], sheet: str = "config") -> Tuple[Dict[str, Any], Dict[str, Dict[str, Any]]]:
    """Лист config -> (настройки DEFAULT, {ряд: перегрузки}). Пустая ячейка = как в DEFAULT."""
    cfg = pd.read_excel(config_file, sheet_name=sheet, dtype=object)
    cfg.columns = [str(c).strip() for c in cfg.columns]
    if "series_id" not in cfg.columns:
        raise ValueError(f"В листе '{sheet}' нет колонки series_id.")
    unknown = [c for c in cfg.columns if c not in CONFIG_COLUMNS]
    if unknown:
        warnings.warn(f"config: неизвестные колонки будут проигнорированы: {unknown}")

    default = dict(DEFAULT_SETTINGS)
    per_series: Dict[str, Dict[str, Any]] = {}
    for _, r in cfg.iterrows():
        sid = r["series_id"]
        if sid is None or (isinstance(sid, float) and math.isnan(sid)) or str(sid).strip() == "":
            continue
        sid = str(sid).strip()
        row = {c: _clean_value(c, r[c], sid) for c in DEFAULT_SETTINGS if c in cfg.columns}
        row = {k: v for k, v in row.items() if v is not None}
        if sid.upper() == "DEFAULT":
            default.update(row)
        else:
            per_series[sid] = row
    return default, per_series


def _data_sheet_name(xlsx: Union[str, Path], data_sheet: Union[int, str, None]) -> Union[int, str]:
    """Какой лист в файле — данные: заданный, иначе первый, не являющийся config/help."""
    if data_sheet is not None:
        return data_sheet
    names = pd.ExcelFile(xlsx).sheet_names
    for n in names:
        if n.strip().lower() not in ("config", "help"):
            return n
    return 0


def sa_ru_make_config(data_file: Union[str, Path], config_file: Optional[Union[str, Path]] = None,
                      data_sheet: Union[int, str, None] = None, overwrite: bool = False,
                      default: Optional[Dict[str, Any]] = None,
                      series: Optional[Dict[str, Dict[str, Any]]] = None) -> Path:
    """Создать лист config (строка DEFAULT + по строке на каждый ряд) и лист help.

    config_file=None (по умолчанию) — листы добавляются в САМ файл с данными;
    иначе создаётся отдельный Excel.
    """
    data_file = Path(data_file)
    into_data = config_file is None
    target = data_file if into_data else Path(config_file)
    if into_data:
        existing = pd.ExcelFile(data_file).sheet_names
        if "config" in existing and not overwrite:
            raise FileExistsError(f"В {data_file.name} уже есть лист config (overwrite=True, чтобы перезаписать).")
    elif target.exists() and not overwrite:
        raise FileExistsError(f"{target} уже существует (overwrite=True, чтобы перезаписать).")
    d = pd.read_excel(data_file, sheet_name=_data_sheet_name(data_file, data_sheet))
    names = [str(c).strip() for c in d.columns[1:]]
    dflt = dict(DEFAULT_SETTINGS)
    dflt.update(default or {})

    def _cell(v):
        if v is None:
            return None
        if isinstance(v, bool):
            return "TRUE" if v else "FALSE"
        return v

    rows = [{"series_id": "DEFAULT", **{k: _cell(v) for k, v in dflt.items()}, "comment": "настройки для всех рядов"}]
    for nm in names:
        r = {"series_id": nm}
        for k, v in (series or {}).get(nm, {}).items():
            r[k] = _cell(v)
        rows.append(r)
    cfg = pd.DataFrame(rows, columns=CONFIG_COLUMNS)
    help_df = pd.DataFrame({"колонка": list(CONFIG_HELP), "что это": list(CONFIG_HELP.values())})

    writer_kw = dict(engine="openpyxl")
    if into_data:
        writer_kw.update(mode="a", if_sheet_exists="replace")     # дописать листы, данные не трогать
    with pd.ExcelWriter(target, **writer_kw) as xw:
        cfg.to_excel(xw, sheet_name="config", index=False)
        help_df.to_excel(xw, sheet_name="help", index=False)
        ws = xw.sheets["config"]
        ws.freeze_panes = "B2"
        ws.column_dimensions["A"].width = 48
        for col_cells in ws.iter_cols(min_col=2, max_col=len(CONFIG_COLUMNS)):
            ws.column_dimensions[col_cells[0].column_letter].width = max(12, len(str(col_cells[0].value)) + 2)
        wh = xw.sheets["help"]
        wh.column_dimensions["A"].width = 20
        wh.column_dimensions["B"].width = 100
    return target


# ============================================================================
# 9. Пакетная обработка: один Excel с данными -> один Excel с результатом
# ============================================================================

def _read_data_wide(data_file: Union[str, Path], sheet: Union[int, str] = 0) -> pd.DataFrame:
    """Первая колонка = дата (заголовок любой), остальные = ряды."""
    d = pd.read_excel(data_file, sheet_name=sheet)
    d = d.rename(columns={d.columns[0]: "date"})
    d.columns = ["date"] + [str(c).strip() for c in d.columns[1:]]
    d["date"] = parse_to_month(d["date"])
    d = d[d["date"].notna()].sort_values("date").reset_index(drop=True)
    if d["date"].duplicated().any():
        raise ValueError("В файле данных есть повторяющиеся месяцы.")
    return d


def sa_ru_batch(data_file: Union[str, Path], calendar_file: Union[str, Path],
                config_file: Optional[Union[str, Path]] = None,
                output_file: Optional[Union[str, Path]] = None,
                data_sheet: Union[int, str, None] = None, config_sheet: str = "config",
                calendar_sheet: Union[int, str] = 0, specs_dir: Union[str, Path] = SPECS_DIR,
                fail_on_error: bool = False, verbose: bool = True, x13_path=None) -> Dict[str, Any]:
    """Сезонная корректировка всех рядов из одного Excel за один вызов.

    data_file   — Excel: лист с данными (первая колонка даты, дальше по колонке на ряд)
                  + лист config с настройками (см. sa_ru_make_config)
    calendar_file — russia_calendar.xlsx (календарь + лист tariff_plan)
    config_file — None (обычно) = лист config в самом data_file (обязателен, см. sa_ru_make_config);
                  или путь к отдельному Excel с листом config
    output_file — куда писать результат; None = <data_file>_SA.xlsx рядом с данными

    Возвращает dict: adjusted (DataFrame как входной файл), seasonal_factor, trend,
    summary (таблица по рядам), results ({ряд: полный вывод sa_ru}), errors.
    """
    say = (lambda *a: print(*a)) if verbose else (lambda *a: None)
    data_file = Path(data_file)

    data = _read_data_wide(data_file, _data_sheet_name(data_file, data_sheet))
    series_names = list(data.columns[1:])
    say(f"Данные: {len(series_names)} рядов, {_ym(data['date'].min())} .. {_ym(data['date'].max())}")

    if config_file is None:
        if config_sheet not in pd.ExcelFile(data_file).sheet_names:
            raise ValueError(
                f"В файле {data_file.name} нет листа '{config_sheet}' с настройками. Создайте его один раз:\n"
                f"    sa_ru_make_config(\"{data_file.name}\")\n"
                "(файл должен быть закрыт в Excel), затем откройте лист config, проверьте строку DEFAULT "
                "(особенно input_type) и при необходимости настройки отдельных рядов."
            )
        config_file = data_file
        say(f"Настройки: лист '{config_sheet}' в {data_file.name}")
    default, per_series = read_config(config_file, config_sheet)
    unknown = sorted(set(per_series) - set(series_names))
    if unknown:
        warnings.warn(f"config: ряды, которых нет в данных: {unknown}")

    daily = load_calendar(calendar_file, calendar_sheet)
    plan = load_tariff_plan(calendar_file)
    find_x13(x13_path)   # проверить сразу, до цикла

    results: Dict[str, Any] = {}
    errors: Dict[str, str] = {}
    rows: List[Dict[str, Any]] = []
    adj = pd.DataFrame({"date": data["date"]})
    fac = pd.DataFrame({"date": data["date"]})
    trd = pd.DataFrame({"date": data["date"]})

    for i, nm in enumerate(series_names, 1):
        cfg = dict(default)
        cfg.update(per_series.get(nm, {}))
        say(f"[{i}/{len(series_names)}] {nm} ...")

        col = data[nm]
        nonna = col.notna()
        if not nonna.any():
            errors[nm] = "ряд пустой"
            rows.append({"series_id": nm, "status": "ERROR", "error": "ряд пустой"})
            adj[nm] = np.nan; fac[nm] = np.nan; trd[nm] = np.nan
            say("    ОШИБКА: ряд пустой")
            continue
        first, last = nonna.idxmax(), nonna[::-1].idxmax()
        sub = data.loc[first:last, ["date", nm]].rename(columns={nm: "value"})

        if cfg.get("skip"):
            adj[nm] = col; fac[nm] = np.nan; trd[nm] = np.nan
            rows.append({"series_id": nm, "status": "SKIPPED", "input_type": cfg["input_type"],
                         "n_obs": int(nonna.sum()), "start": _ym(sub["date"].iloc[0]), "end": _ym(sub["date"].iloc[-1])})
            say("    пропущен (skip = TRUE), в выходной файл как есть")
            continue

        tariff_series = None
        if cfg.get("tariff") == "regressor":
            ts_name = cfg.get("tariff_series")
            if ts_name:
                if ts_name not in data.columns:
                    errors[nm] = f"tariff_series '{ts_name}' нет в данных"
                    rows.append({"series_id": nm, "status": "ERROR", "error": errors[nm]})
                    adj[nm] = np.nan; fac[nm] = np.nan; trd[nm] = np.nan
                    say("    ОШИБКА: " + errors[nm])
                    if fail_on_error:
                        raise ValueError(errors[nm])
                    continue
                ts_cfg = dict(default); ts_cfg.update(per_series.get(ts_name, {}))
                tcol = data[ts_name]
                tvals = tcol.to_numpy(float)
                if ts_cfg.get("input_type") == "log_level":
                    tvals = _mom_from_loglevel(np.where(np.isnan(tvals), 0, tvals))
                    tvals[np.isnan(tcol.to_numpy(float))] = np.nan
                tariff_series = pd.Series(tvals, index=pd.DatetimeIndex(data["date"])).dropna()

        try:
            if cfg.get("freeze"):
                res = sa_ru_apply(sub, calendar_file, series_id=nm, dir=specs_dir,
                                  date_col="date", value_col="value",
                                  tariff_series=tariff_series, tariff_plan=plan,
                                  verbose=False, x13_path=x13_path, _calendar_daily=daily)
                if res["spec"]["input_type"] != cfg["input_type"]:
                    raise ValueError(f"input_type в config ('{cfg['input_type']}') не совпадает со "
                                     f"спецификацией ('{res['spec']['input_type']}') — переидентифицируйте ряд.")
            else:
                res = sa_ru(sub, calendar_file, date_col="date", value_col="value",
                            input_type=cfg["input_type"], calendar_sheet=calendar_sheet,
                            calendar_mode=cfg["calendar_mode"], include_easter=cfg["include_easter"],
                            use_outliers=cfg["use_outliers"], outlier_types=cfg["outlier_types"],
                            outlier_critical=cfg["outlier_critical"], transform=cfg["transform"],
                            method=cfg["method"], forecast_months=cfg["forecast_months"],
                            seasonality_alpha=cfg["seasonality_alpha"], tariff=cfg["tariff"],
                            tariff_series=tariff_series, tariff_plan=plan, series_name=nm,
                            verbose=False, x13_path=x13_path, _calendar_daily=daily)
        except Exception as e:                       # noqa: BLE001
            errors[nm] = str(e)
            rows.append({"series_id": nm, "status": "ERROR", "input_type": cfg["input_type"], "error": str(e)})
            adj[nm] = np.nan; fac[nm] = np.nan; trd[nm] = np.nan
            say(f"    ОШИБКА: {e}")
            if fail_on_error:
                raise
            continue

        results[nm] = res
        d = res["data"].set_index("date")
        adj[nm] = d["adjusted"].reindex(data["date"]).to_numpy()
        fac[nm] = d["seasonal_factor"].reindex(data["date"]).to_numpy()
        trd[nm] = d["trend"].reindex(data["date"]).to_numpy()

        status = "OK" if res["seasonality_detected"] else "NO_SEASONALITY"
        cmp_tbl = res.get("comparison", pd.DataFrame())
        cmp_row = cmp_tbl[cmp_tbl.get("chosen", True) == True] if len(cmp_tbl) else None  # noqa: E712
        qs = res.get("qs_orig_evadj_pval", np.nan)
        if qs is None or (isinstance(qs, float) and math.isnan(qs)):
            qs = res.get("qs_orig_pval", np.nan)
        frozen = bool(res.get("coef_frozen", False))
        rows.append({
            "series_id": nm, "status": status, "input_type": cfg["input_type"], "tariff": cfg["tariff"],
            "frozen": bool(cfg.get("freeze")), "coef_frozen": frozen,
            "spec_through": res["spec"]["data_through"] if cfg.get("freeze") else "",
            "new_outliers": ", ".join(res.get("new_outliers", [])),
            "seasonality_detected": res["seasonality_detected"], "qs_pval": qs,
            "calendar_mode": res["chosen_model"], "easter": bool(res.get("include_easter", False)),
            "transform": res["transform"],
            "method": res["decomposition_method"], "arima": res["arima"],
            "n_outliers": len(res["outliers"]), "outliers": ", ".join(res["outliers"]),
            "aicc": float(cmp_row["aicc"].iloc[0]) if cmp_row is not None and len(cmp_row) else np.nan,
            "cal_signif": int(cmp_row["cal_signif"].iloc[0]) if cmp_row is not None and len(cmp_row) else np.nan,
            "cal_total": int(cmp_row["cal_total"].iloc[0]) if cmp_row is not None and len(cmp_row) else np.nan,
            "n_obs": len(d), "start": _ym(d.index[0]), "end": _ym(d.index[-1]),
            "warnings": " | ".join(res.get("warnings", [])), "error": "",
        })
        tag = "заморожено | " if cfg.get("freeze") else ""
        if res["tariff"] == "schedule":
            say("    OK | тарифы ЖКУ по формуле ЦБ"
                + (" | " + " | ".join(res.get("warnings", [])) if res.get("warnings") else ""))
        else:
            say(f"    {status} | {tag}календарь={res['chosen_model']}"
                + ("+Пасха" if res.get("include_easter") else "")
                + f" | {res['decomposition_method']} | ARIMA {res['arima']}"
                f" | выбросов {len(res['outliers'])}"
                + (f" | QS p={qs:.3f}" if isinstance(qs, float) and not math.isnan(qs) else ""))

    summary = pd.DataFrame(rows)
    col_order = ["series_id", "status", "input_type", "tariff", "frozen", "coef_frozen", "spec_through",
                 "seasonality_detected", "qs_pval", "calendar_mode", "easter", "transform", "method", "arima",
                 "n_outliers", "outliers", "new_outliers", "aicc", "cal_signif", "cal_total",
                 "n_obs", "start", "end", "warnings", "error"]
    summary = summary.reindex(columns=[c for c in col_order if c in summary.columns])
    ok_n = int((summary["status"] == "OK").sum()) if not summary.empty else 0
    say(f"Готово: OK={ok_n}, без сезонности={int((summary['status'] == 'NO_SEASONALITY').sum())}, "
        f"пропущено={int((summary['status'] == 'SKIPPED').sum())}, ошибок={len(errors)}")

    if output_file is None:
        output_file = data_file.with_name(data_file.stem + "_SA.xlsx")
    output_file = Path(output_file)
    _write_output(output_file, adj, fac, trd, summary)
    say(f"Результат сохранён: {output_file}")

    return {"adjusted": adj, "seasonal_factor": fac, "trend": trd, "summary": summary,
            "results": results, "errors": errors, "output_file": output_file}


def _write_output(path: Path, adj: pd.DataFrame, fac: pd.DataFrame, trd: pd.DataFrame, summary: pd.DataFrame) -> None:
    with pd.ExcelWriter(path, engine="openpyxl", datetime_format="yyyy-mm-dd") as xw:
        adj.to_excel(xw, sheet_name="adjusted", index=False)
        fac.to_excel(xw, sheet_name="seasonal_factor", index=False)
        trd.to_excel(xw, sheet_name="trend", index=False)
        summary.to_excel(xw, sheet_name="summary", index=False)
        for name in ("adjusted", "seasonal_factor", "trend"):
            ws = xw.sheets[name]
            ws.freeze_panes = "B2"
            ws.column_dimensions["A"].width = 12
        ws = xw.sheets["summary"]
        ws.freeze_panes = "B2"
        ws.column_dimensions["A"].width = 48
        for col_cells in ws.iter_cols(min_col=2, max_col=ws.max_column):
            width = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells[:200])
            ws.column_dimensions[col_cells[0].column_letter].width = min(max(10, width + 2), 60)


# ============================================================================
# 10. ЗАМОРОЗКА СПЕЦИФИКАЦИИ И КОЭФФИЦИЕНТОВ
#     sa_ru_identify() — раз в год: подбирает модель и записывает её в sa_specs/<ряд>.json
#     sa_ru_apply()    — каждый месяц: применяет замороженную модель к новым данным
# ============================================================================

def spec_path(series_id: str, dir: Union[str, Path] = SPECS_DIR) -> Path:
    """Файл спецификации ряда: имя файла = имя ряда (кириллица, пробелы, запятые — можно)."""
    name = "".join("_" if (ch in _FS_FORBIDDEN or ord(ch) < 32) else ch for ch in str(series_id)).strip(" .")
    return Path(dir) / ((name[:120] or "series") + ".json")


def _spec_defaults(spec: Dict[str, Any]) -> Dict[str, Any]:
    spec.setdefault("has_constant", False)
    spec.setdefault("coef_names", [])
    spec.setdefault("coef_values", [])
    spec.setdefault("arima_ar", [])
    spec.setdefault("arima_ma", [])
    return spec


def sa_ru_identify(df: pd.DataFrame, calendar_file: Union[str, Path], series_id: Optional[str] = None, *,
                   cutoff_date=None, dir: Union[str, Path] = SPECS_DIR, freeze_coef: bool = True,
                   date_col: str = "date", value_col: str = "value", tariff_series_name: Optional[str] = None,
                   verbose: bool = True, _calendar_daily: Optional[pd.DataFrame] = None,
                   **sa_kwargs) -> Dict[str, Any]:
    """Подобрать модель на данных ПО cutoff_date включительно и заморозить её.

    Записывает в sa_specs/<series_id>.json: преобразование, режим календаря, Пасху,
    порядок ARIMA, список выбросов, окно центрирования И САМИ КОЭФФИЦИЕНТЫ
    (календарные беты, величины выбросов, AR/MA). Дальше sa_ru_apply() каждый месяц
    считает ряд по этой замороженной модели, не переоценивая историю.

    series_id=None -> спецификация не сохраняется, только возвращается.
    Остальные аргументы (input_type, calendar_mode, include_easter, ...) — как у sa_ru().
    """
    say = (lambda *a: print(*a)) if verbose else (lambda *a: None)

    dat_full = prepare_monthly_df(df, date_col, value_col)
    cutoff = parse_to_month(cutoff_date)[0] if cutoff_date is not None else dat_full["date"].max()
    dat = dat_full[dat_full["date"] <= cutoff].reset_index(drop=True)
    if len(dat) < 36:
        raise ValueError(f"Слишком короткая выборка для идентификации по {_ym(cutoff)}: "
                         f"{len(dat)} месяцев (нужно >= 36).")
    if dat_full["date"].max() < cutoff:
        warnings.warn(f"Данные заканчиваются {_ym(dat_full['date'].max())}, раньше границы {_ym(cutoff)}.")
    sample_start = dat["date"].min()
    sa_kwargs.setdefault("center_start", sample_start)
    sa_kwargs.setdefault("center_end", cutoff)

    res = sa_ru(dat, calendar_file, date_col="date", value_col="value", verbose=verbose,
                _calendar_daily=_calendar_daily, **sa_kwargs)

    af = cutoff + pd.DateOffset(months=1)
    spec: Dict[str, Any] = {
        "spec_version": SPEC_VERSION,
        "series_id": series_id,
        "identified_on": date.today().isoformat(),
        "input_type": res["input_type"],
        "sample_start": _ym(sample_start),
        "data_through": _ym(cutoff),
        "transform": res["transform"],
        "calendar_mode": res["chosen_model"],
        "include_easter": bool(res["include_easter"]),
        "method": sa_kwargs.get("method", "prefer_seats"),
        "decomposition_used": res["decomposition_method"],
        "arima": res["arima"],
        "outliers": list(res["outliers_x13"]),
        "outliers_readable": list(res["outliers"]),
        "outlier_auto_from": f"{af.year}.{af.month}",
        "outlier_types": sa_kwargs.get("outlier_types", "all"),
        "outlier_critical": sa_kwargs.get("outlier_critical"),
        "seasonality_detected": bool(res["seasonality_detected"]),
        "forecast_months": int(sa_kwargs.get("forecast_months", 36)),
        "center_start": _ym(sample_start),
        "center_end": _ym(cutoff),
        "xreg_names": list(res["xreg_names"]),
        "tariff": res["tariff"],
        "tariff_series": tariff_series_name,
        "calendar_sheet": sa_kwargs.get("calendar_sheet", 0),
        "freeze_coef": False,
        "has_constant": False,
        "coef_names": [], "coef_values": [], "arima_ar": [], "arima_ma": [],
        "seats_model": "",
        "spec_reproduces_max_gap": None,
        "data_check": {"n": int(len(dat)), "s1": float(dat["value"].sum()),
                       "s2": float((dat["value"] ** 2).sum())},
    }

    # --- коэффициенты: повторяем модель с ЯВНЫМИ выбросами и фиксированной ARIMA
    if freeze_coef and res["seasonality_detected"] and res["tariff"] != "schedule":
        y = _to_internal(dat["value"].to_numpy(float), res["input_type"])
        dates = pd.DatetimeIndex(dat["date"])
        fit0 = res["best_fit"]
        vars_written = (["const"] if fit0.has_constant else []) + list(res["outliers_x13"])
        refit = fit_with_fallback(spec["method"], y=y, dates=dates, xreg=res["xreg"], xreg_types=XREG_TYPES,
                                  transform=res["transform"], auto_outlier=False,
                                  fixed_outliers=vars_written or None, arima_model=res["arima"],
                                  forecast=spec["forecast_months"], x13_path=sa_kwargs.get("x13_path"))
        if not refit.ok or refit.sa is None:
            warnings.warn(f"Не удалось повторить модель с явными выбросами -> коэффициенты не заморожены "
                          f"({refit.error}). Спецификация сохранена без них.")
        else:
            gap = float(np.nanmax(np.abs(refit.sa.reindex(dates).to_numpy(float)
                                         - fit0.sa.reindex(dates).to_numpy(float))))
            reg_map = refit.reg_map
            ar, ma = refit.arima_split()
            spec.update({
                "freeze_coef": True,
                "has_constant": bool(refit.has_constant),
                "coef_names": list(reg_map.keys()),
                "coef_values": [float(x) for x in reg_map.values()],
                "arima_ar": [float(x) for x in ar],
                "arima_ma": [float(x) for x in ma],
                "seats_model": refit.seats_model,
                "spec_reproduces_max_gap": gap,
            })
            say(f"  коэффициентов заморожено: регрессионных {len(reg_map)}, AR {len(ar)}, MA {len(ma)}"
                f" | спецификация воспроизводит ряд с точностью {gap:.2e}")
            if gap > 1e-6:
                warnings.warn(f"Повтор модели с явными выбросами расходится с автоподбором на {gap:.2e} — "
                              "проверьте ряд перед тем, как полагаться на заморозку.")

    if series_id is not None:
        path = spec_path(series_id, dir)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(spec, ensure_ascii=False, indent=2), encoding="utf-8")
        say(f"  спецификация сохранена: {path}")
    else:
        say("  series_id не задан -> спецификация не сохранена (только возвращена).")
    return spec


def sa_ru_apply(df: pd.DataFrame, calendar_file: Union[str, Path], series_id: str, *,
                dir: Union[str, Path] = SPECS_DIR, date_col: str = "date", value_col: str = "value",
                tariff_series: Optional[pd.Series] = None, tariff_plan: Optional[pd.DataFrame] = None,
                freeze_coef: Optional[bool] = None, coef_tol: float = 1e-5,
                verbose: bool = True, x13_path=None,
                _calendar_daily: Optional[pd.DataFrame] = None) -> Dict[str, Any]:
    """Посчитать ряд по замороженной спецификации из sa_specs/<series_id>.json.

    Модель, выбросы и коэффициенты берутся из спецификации; автопоиск новых выбросов
    разрешён только ПОСЛЕ границы заморозки (правило Банка России). История не
    переоценивается. freeze_coef=None -> как записано в спецификации.
    """
    say = (lambda *a: print(*a)) if verbose else (lambda *a: None)
    warns: List[str] = []

    path = spec_path(series_id, dir)
    if not path.exists():
        raise FileNotFoundError(f"Нет спецификации {path}. Сначала запустите идентификацию "
                                f"(sa_ru_identify_batch / sa_ru_identify) для ряда '{series_id}'.")
    spec = _spec_defaults(json.loads(path.read_text(encoding="utf-8")))
    if spec.get("series_id") and str(spec["series_id"]) != str(series_id):
        raise ValueError(f"В файле {path.name} лежит спецификация ряда '{spec['series_id']}', "
                         f"а применяется к '{series_id}'. Переименуйте файл или переидентифицируйте ряд.")

    dat = prepare_monthly_df(df, date_col, value_col)
    dates = pd.DatetimeIndex(dat["date"])
    v = dat["value"].to_numpy(float)
    n = len(dat)
    input_type = spec["input_type"]
    border = parse_to_month(spec["data_through"])[0]

    if _ym(dates.min()) != spec["sample_start"]:
        raise ValueError(f"Спецификация оценена на выборке с {spec['sample_start']}, а данные начинаются "
                         f"с {_ym(dates.min())}. Замороженная модель к другой стартовой точке неприменима — "
                         "переидентифицируйте ряд.")
    if dates.max() < border:
        warns.append(f"Данные заканчиваются {_ym(dates.max())} — раньше границы заморозки {spec['data_through']}.")

    # история периода идентификации не должна была измениться
    dc = spec.get("data_check") or {}
    if dc:
        idw = dat[dat["date"] <= border]["value"]
        same = (len(idw) == dc.get("n") and np.isclose(idw.sum(), dc.get("s1", np.nan), rtol=1e-9, atol=1e-9)
                and np.isclose((idw ** 2).sum(), dc.get("s2", np.nan), rtol=1e-9, atol=1e-9))
        if not same:
            warns.append(f"Данные периода {spec['sample_start']}..{spec['data_through']} изменились с момента "
                         "идентификации (пересмотр Росстата?) — замороженные коэффициенты оценены на другой "
                         "выборке, ряд стоит переидентифицировать.")

    base = {"spec": spec, "warnings": warns, "input_type": input_type, "tariff": spec["tariff"],
            "chosen_model": spec["calendar_mode"], "include_easter": spec["include_easter"],
            "transform": spec["transform"], "arima": spec["arima"],
            "outliers": [normalize_outlier_name(o) for o in spec["outliers"]],
            "seasonality_detected": spec["seasonality_detected"], "frozen": True}

    # --- тарифный режим: формула ЦБ, модели нет
    if spec["tariff"] == "schedule":
        mom = (np.concatenate([[100.0], 100.0 * v[1:] / v[:-1]]) if input_type == "level"
               else _mom_from_loglevel(_to_internal(v, input_type)))
        sa_mom, by_year, w = tariff_schedule_adjust(dates, mom, tariff_plan)
        warns += w
        if input_type == "mom_index":
            adjusted = sa_mom
        elif input_type == "log_level":
            adjusted = np.cumsum(np.log(sa_mom / 100.0))
        else:
            adjusted = v[0] * np.cumprod(np.concatenate([[1.0], sa_mom[1:] / 100.0]))
        seas = v - adjusted if input_type == "log_level" else v / adjusted
        data = pd.DataFrame({"date": dates, "original": v, "adjusted": adjusted, "trend": np.nan,
                             "seasonal_factor": seas,
                             "factor_type": "additive" if input_type == "log_level" else "multiplicative"})
        say("  тарифный режим: заморозка не нужна, расчёт детерминированный")
        return {**base, "data": data, "new_points": data[data["date"] > border].copy(),
                "decomposition_method": "schedule", "coef_frozen": False, "coef_check": None,
                "new_outliers": [], "tariff_by_year": by_year, "best_fit": None}

    # --- сезонности нет: ряд как есть
    if not spec["seasonality_detected"]:
        data = pd.DataFrame({"date": dates, "original": v, "adjusted": v, "trend": np.nan,
                             "seasonal_factor": 1.0 if input_type != "log_level" else 0.0,
                             "factor_type": "additive" if input_type == "log_level" else "multiplicative"})
        say("  в спецификации отмечено отсутствие сезонности -> ряд возвращается как есть")
        return {**base, "data": data, "new_points": data[data["date"] > border].copy(),
                "decomposition_method": "", "coef_frozen": False, "coef_check": None,
                "new_outliers": [], "tariff_by_year": None, "best_fit": None}

    # --- тот же дизайн регрессоров, что при идентификации
    daily = _calendar_daily if _calendar_daily is not None else load_calendar(
        calendar_file, spec.get("calendar_sheet", 0))
    xreg_dates = pd.date_range(dates.min(), dates.max() + pd.DateOffset(months=spec["forecast_months"]), freq="MS")
    xr = _make_xreg(xreg_dates, daily, spec["calendar_mode"], spec["include_easter"],
                    spec["center_start"], spec["center_end"], spec["tariff"], tariff_series, tariff_plan)
    names = [] if xr is None else list(xr.columns)
    if names != list(spec["xreg_names"]):
        raise ValueError("Состав календарных регрессоров не совпал со спецификацией.\n"
                         f"  при идентификации: {spec['xreg_names']}\n  сейчас собралось: {names}\n"
                         "Проверьте производственный календарь и настройки, либо переидентифицируйте ряд.")

    y = _to_internal(v, input_type)
    fixed_out = list(spec["outliers"])
    af_y, af_m = (int(x) for x in str(spec["outlier_auto_from"]).split("."))
    do_auto = pd.Timestamp(year=af_y, month=af_m, day=1) <= dates.max()
    const_vars = ["const"] if spec.get("has_constant") else []

    common = dict(y=y, dates=dates, xreg=xr, xreg_types=XREG_TYPES, transform=spec["transform"],
                  arima_model=spec["arima"], forecast=spec["forecast_months"], x13_path=x13_path)

    # ЭТАП A: коэффициенты свободны, новые выбросы ищем только после границы
    fitA = fit_with_fallback(spec["method"], auto_outlier=do_auto,
                             outlier_span=(f"{spec['outlier_auto_from']}," if do_auto else None),
                             outlier_types=spec.get("outlier_types", "all"),
                             outlier_critical=spec.get("outlier_critical"),
                             fixed_outliers=(const_vars + fixed_out) or None, **common)
    if not fitA.ok:
        raise RuntimeError(f"Не удалось применить спецификацию '{series_id}'. Ошибка X-13: {fitA.error}")
    known = {normalize_outlier_name(o) for o in fixed_out}
    new_outliers = [o for o in fitA.outliers_x13 if normalize_outlier_name(o) not in known]

    want_freeze = spec.get("freeze_coef", False) if freeze_coef is None else bool(freeze_coef)
    if want_freeze and not (spec["coef_values"] or spec["arima_ar"] or spec["arima_ma"]):
        warns.append("В спецификации нет коэффициентов -> заморожена только спецификация, "
                     "коэффициенты переоцениваются.")
        want_freeze = False

    coef_check = None
    coef_frozen = False
    fit = fitA

    if want_freeze:
        all_out = fixed_out + new_outliers
        vars_written = const_vars + all_out
        # ЭТАП B: детекция выключена, все выбросы явные — отсюда стартовые значения для новых
        fitB = fit_with_fallback(spec["method"], auto_outlier=False,
                                 fixed_outliers=vars_written or None, **common)
        if not fitB.ok:
            raise RuntimeError(f"Не удалось собрать модель с явными выбросами для '{series_id}': {fitB.error}")

        frozen = dict(zip(spec["coef_names"], [float(x) for x in spec["coef_values"]]))
        mapB = fitB.reg_map
        const_key = ["Constant"] if const_vars else []
        ar_str = [fmt_coef(x, True) for x in spec["arima_ar"]]
        ma_str = [fmt_coef(x, True) for x in spec["arima_ma"]]

        # порядок b: как X-13 строит матрицу — сначала variables, потом user.
        # На всякий случай проверяем и обратный порядок, правильный определяем сверкой.
        orders = [const_key + all_out + names, names + const_key + all_out]
        for order in orders:
            bvec = [fmt_coef(frozen[nm], True) if nm in frozen else fmt_coef(mapB.get(nm, 0.0), False)
                    for nm in order]
            cand = fit_with_fallback(spec["method"], auto_outlier=False, fixed_outliers=vars_written or None,
                                     fixed_b=bvec, arima_ar=ar_str or None, arima_ma=ma_str or None, **common)
            if not cand.ok:
                continue
            got = cand.reg_map
            ar_c, ma_c = cand.arima_split()
            chk = pd.DataFrame(
                [{"coefficient": k, "frozen": val, "realized": got.get(k, np.nan)} for k, val in frozen.items()]
                + [{"coefficient": f"AR{i+1}", "frozen": a, "realized": (ar_c[i] if i < len(ar_c) else np.nan)}
                   for i, a in enumerate(spec["arima_ar"])]
                + [{"coefficient": f"MA{i+1}", "frozen": m, "realized": (ma_c[i] if i < len(ma_c) else np.nan)}
                   for i, m in enumerate(spec["arima_ma"])])
            chk["abs_diff"] = (chk["frozen"] - chk["realized"]).abs()
            worst = float(chk["abs_diff"].max()) if len(chk) else 0.0
            if np.isfinite(worst) and worst < coef_tol:
                fit, coef_check, coef_frozen = cand, chk, True
                break
            if coef_check is None:
                coef_check = chk
        if not coef_frozen:
            raise RuntimeError(
                f"Не удалось зафиксировать коэффициенты ряда '{series_id}': X-13 вернул другие значения "
                f"(макс. расхождение {float(coef_check['abs_diff'].max()):.3g}). "
                "Временный обход — freeze=FALSE для этого ряда.")
        if spec.get("seats_model") and fit.seats_model and spec["seats_model"] != fit.seats_model:
            warns.append(f"SEATS использует другую модель, чем при идентификации: было {spec['seats_model']}, "
                         f"стало {fit.seats_model} — фильтры разложения изменились.")

    adj_y = fit.sa.reindex(dates).to_numpy(float)
    tr = fit.trend
    trend_y = tr.reindex(dates).to_numpy(float) if tr is not None else np.full(n, np.nan)
    data = pd.DataFrame({"date": dates,
                         **_to_output(input_type, v, adj_y, trend_y, spec["transform"],
                                      _seas0_estimate(fit, dates))})

    say(f"  спецификация от {spec['identified_on']}, заморожена по {spec['data_through']} | "
        f"{fit.method} | коэффициенты заморожены: {coef_frozen} | "
        f"новых выбросов после границы: {len(new_outliers)} | "
        f"новых точек: {int((dates > border).sum())}")
    for w in warns:
        say("  ! " + w)

    return {**base, "data": data, "new_points": data[data["date"] > border].copy(),
            "decomposition_method": fit.method, "coef_frozen": coef_frozen, "coef_check": coef_check,
            "new_outliers": [normalize_outlier_name(o) for o in new_outliers],
            "qs_orig_pval": fit.qs_pval("qsori"), "qs_orig_evadj_pval": fit.qs_pval("qsorievadj"),
            "tariff_by_year": None, "coefficients": fit.reg, "udg": fit.udg, "best_fit": fit}


def sa_ru_list_specs(dir: Union[str, Path] = SPECS_DIR) -> pd.DataFrame:
    """Все сохранённые спецификации из папки: что заморожено и когда."""
    d = Path(dir)
    if not d.exists():
        print(f"Папки '{d}' нет — спецификаций пока не создавали.")
        return pd.DataFrame()
    rows = []
    for f in sorted(d.glob("*.json")):
        try:
            s = _spec_defaults(json.loads(f.read_text(encoding="utf-8")))
        except Exception as e:                       # noqa: BLE001
            rows.append({"файл": f.name, "ошибка": str(e)})
            continue
        rows.append({"series_id": s.get("series_id"), "identified_on": s.get("identified_on"),
                     "data_through": s.get("data_through"), "input_type": s.get("input_type"),
                     "calendar_mode": s.get("calendar_mode"), "easter": s.get("include_easter"),
                     "transform": s.get("transform"), "arima": s.get("arima"),
                     "n_outliers": len(s.get("outliers", [])), "tariff": s.get("tariff"),
                     "seasonality": s.get("seasonality_detected"), "coef_frozen": s.get("freeze_coef"),
                     "n_coef": len(s.get("coef_values", [])) + len(s.get("arima_ar", [])) + len(s.get("arima_ma", []))})
    return pd.DataFrame(rows)


def sa_ru_identify_batch(data_file: Union[str, Path], calendar_file: Union[str, Path], cutoff_date: str, *,
                         config_file: Optional[Union[str, Path]] = None, dir: Union[str, Path] = SPECS_DIR,
                         data_sheet: Union[int, str, None] = None, config_sheet: str = "config",
                         calendar_sheet: Union[int, str] = 0, only_freeze: bool = True,
                         verbose: bool = True, x13_path=None) -> pd.DataFrame:
    """Идентификация (раз в год): подобрать и заморозить модели для рядов из файла.

    Обрабатываются ряды, у которых на листе config стоит freeze = TRUE
    (only_freeze=False -> все ряды). Для каждого создаётся файл
    <dir>/<имя ряда>.json, который потом используется при обычном прогоне.

    cutoff_date — по какой месяц включительно фиксируем модель, например "2025-12".
    """
    say = (lambda *a: print(*a)) if verbose else (lambda *a: None)
    data_file = Path(data_file)
    data = _read_data_wide(data_file, _data_sheet_name(data_file, data_sheet))
    series_names = list(data.columns[1:])

    if config_file is None:
        if config_sheet not in pd.ExcelFile(data_file).sheet_names:
            raise ValueError(f"В файле {data_file.name} нет листа '{config_sheet}'. "
                             f"Создайте его: sa_ru_make_config(\"{data_file.name}\").")
        config_file = data_file
    default, per_series = read_config(config_file, config_sheet)

    daily = load_calendar(calendar_file, calendar_sheet)
    plan = load_tariff_plan(calendar_file)
    find_x13(x13_path)

    todo = []
    for nm in series_names:
        cfg = dict(default); cfg.update(per_series.get(nm, {}))
        if cfg.get("skip"):
            continue
        if only_freeze and not cfg.get("freeze"):
            continue
        todo.append((nm, cfg))
    if not todo:
        say("Нет рядов с freeze = TRUE на листе config — идентифицировать нечего.")
        return pd.DataFrame()

    say(f"Идентификация по {cutoff_date} включительно: {len(todo)} рядов -> {Path(dir).resolve()}")
    rows = []
    for i, (nm, cfg) in enumerate(todo, 1):
        say(f"[{i}/{len(todo)}] {nm} ...")
        col = data[nm]
        nonna = col.notna()
        if not nonna.any():
            rows.append({"series_id": nm, "status": "ERROR", "error": "ряд пустой"})
            say("    ОШИБКА: ряд пустой")
            continue
        first, last = nonna.idxmax(), nonna[::-1].idxmax()
        sub = data.loc[first:last, ["date", nm]].rename(columns={nm: "value"})

        tariff_series = None
        ts_name = cfg.get("tariff_series")
        if cfg.get("tariff") == "regressor" and ts_name and ts_name in data.columns:
            tariff_series = pd.Series(data[ts_name].to_numpy(float),
                                      index=pd.DatetimeIndex(data["date"])).dropna()
        try:
            spec = sa_ru_identify(
                sub, calendar_file, series_id=nm, cutoff_date=cutoff_date, dir=dir,
                freeze_coef=True, tariff_series_name=ts_name, verbose=False, _calendar_daily=daily,
                input_type=cfg["input_type"], calendar_sheet=calendar_sheet,
                calendar_mode=cfg["calendar_mode"], include_easter=cfg["include_easter"],
                use_outliers=cfg["use_outliers"], outlier_types=cfg["outlier_types"],
                outlier_critical=cfg["outlier_critical"], transform=cfg["transform"],
                method=cfg["method"], forecast_months=cfg["forecast_months"],
                seasonality_alpha=cfg["seasonality_alpha"], tariff=cfg["tariff"],
                tariff_series=tariff_series, tariff_plan=plan, x13_path=x13_path)
        except Exception as e:                       # noqa: BLE001
            rows.append({"series_id": nm, "status": "ERROR", "error": str(e)})
            say(f"    ОШИБКА: {e}")
            continue
        rows.append({"series_id": nm, "status": "OK", "data_through": spec["data_through"],
                     "calendar_mode": spec["calendar_mode"], "easter": spec["include_easter"],
                     "transform": spec["transform"], "arima": spec["arima"],
                     "n_outliers": len(spec["outliers"]), "outliers": ", ".join(spec["outliers_readable"]),
                     "seasonality": spec["seasonality_detected"], "coef_frozen": spec["freeze_coef"],
                     "n_coef": len(spec["coef_values"]) + len(spec["arima_ar"]) + len(spec["arima_ma"]),
                     "reproduces": spec["spec_reproduces_max_gap"], "error": ""})
        say(f"    OK | календарь={spec['calendar_mode']}" + ("+Пасха" if spec["include_easter"] else "")
            + f" | ARIMA {spec['arima']} | выбросов {len(spec['outliers'])}"
            + f" | коэффициентов заморожено {rows[-1]['n_coef']}")

    out = pd.DataFrame(rows)
    say(f"Готово: успешно {int((out['status'] == 'OK').sum())}, ошибок {int((out['status'] == 'ERROR').sum())}."
        f" Спецификации: {Path(dir).resolve()}")
    say("Дальше обычный прогон (sa_ru_batch) будет считать эти ряды по замороженным моделям.")
    return out
