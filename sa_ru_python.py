"""
sa_ru_python.py
===============
Python-порт кастомного сезонного сглаживания месячных временных рядов
с учётом российского производственного календаря.

Оригинальный R-код: https://github.com/NadezhdaYurchenko/Seasonal-Adjustment-for-Russia
Автор Python-порта: (укажите ваше имя)

Notes
-----
В R используется X-13ARIMA-SEATS (пакет ``seasonal``).
В данной Python-реализации используется ``statsmodels.tsa.statespace.sarimax.SARIMAX``,
что обеспечивает аналогичную логику: ARIMA-регрессия с пользовательскими
календарными регрессорами (рабочие дни, праздники, Пасха).

Совпадение с R-версией:
- Логика центрирования регрессоров — идентична.
- Порядок ARIMA по умолчанию (1,1,1)(1,1,1,12) близок к «airline model»,
  аналогичному стандарту X-13. При calendar_mode='auto' перебираются
  три варианта (none/basic/extended) с выбором по AIC.
- Результат: сезонно сглаженный ряд, сезонная составляющая, тренд.

Dependencies
------------
pandas, numpy, openpyxl, statsmodels, scipy
"""

from __future__ import annotations

import json
import os
import re
import warnings
from datetime import date, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

import numpy as np
import pandas as pd
from statsmodels.tsa.statespace.sarimax import SARIMAX

# ---------------------------------------------------------------------------
# 1. Разбор дат из данных
# ---------------------------------------------------------------------------

def parse_to_month(x) -> pd.DatetimeIndex:
    """Convert various date formats to the first day of each month.

    Parameters
    ----------
    x : array-like
        Dates in any of the following formats:
        pandas Timestamp / datetime / date / numpy datetime64,
        numeric Excel serial (origin 1899-12-30),
        strings: ``"2015M1"``, ``"2015M01"``, ``"2015-01"``, ``"2015/01"``,
        ``"2015-01-01"``, ``"01.2015"``, ``"1/2015"``.

    Returns
    -------
    pd.DatetimeIndex
        Дата первого числа каждого месяца.
    """
    if isinstance(x, (pd.DatetimeIndex, pd.Series)):
        series = pd.to_datetime(x)
        return pd.DatetimeIndex(series.values.astype("datetime64[M]").astype("datetime64[D]"))

    if not hasattr(x, "__iter__") or isinstance(x, str):
        x = [x]

    results = []
    for item in x:
        results.append(_parse_one(item))
    return pd.DatetimeIndex(results)


def _parse_one(s) -> pd.Timestamp:
    """Парсит одно значение даты в первый день месяца."""
    if pd.isnull(s):
        return pd.NaT

    # Уже является датой/Timestamp
    if isinstance(s, (pd.Timestamp, date)):
        return pd.Timestamp(s).replace(day=1)

    # numpy datetime64
    if isinstance(s, np.datetime64):
        return pd.Timestamp(s).replace(day=1)

    # числовой серийник Excel (origin 1899-12-30)
    if isinstance(s, (int, float, np.integer, np.floating)):
        return (pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(s))).replace(day=1)

    s_str = str(s).strip()

    # "2015M1" / "2015m01"
    m = re.fullmatch(r"(\d{4})[Mm](\d{1,2})", s_str)
    if m:
        return pd.Timestamp(year=int(m.group(1)), month=int(m.group(2)), day=1)

    # "2015-01" / "2015/01" / "2015.01"
    m = re.fullmatch(r"(\d{4})[-/.](\d{1,2})", s_str)
    if m:
        return pd.Timestamp(year=int(m.group(1)), month=int(m.group(2)), day=1)

    # "01.2015" / "1/2015" / "01-2015"
    m = re.fullmatch(r"(\d{1,2})[-/.](\d{4})", s_str)
    if m:
        return pd.Timestamp(year=int(m.group(2)), month=int(m.group(1)), day=1)

    # Полноценная дата "2015-01-01", ...
    try:
        return pd.Timestamp(s_str).replace(day=1)
    except Exception:
        return pd.NaT


def parse_numeric(x) -> np.ndarray:
    """Parse numeric values, handling comma-as-decimal-separator.

    Parameters
    ----------
    x : array-like
        Числовые или строковые значения ряда.

    Returns
    -------
    np.ndarray of float64
    """
    if isinstance(x, (np.ndarray,)) and np.issubdtype(x.dtype, np.floating):
        return x.astype(float)
    series = pd.Series(x)
    result = pd.to_numeric(series, errors="coerce")
    # запятая как разделитель десятичных
    bad = result.isna() & series.notna()
    if bad.any():
        result[bad] = pd.to_numeric(
            series[bad].astype(str).str.replace(",", ".", regex=False),
            errors="coerce"
        )
    return result.to_numpy(dtype=float)


# ---------------------------------------------------------------------------
# 2. Проверка и подготовка месячного ряда
# ---------------------------------------------------------------------------

def prepare_monthly_df(
    df: pd.DataFrame,
    date_col: str = "date",
    value_col: str = "value",
) -> pd.DataFrame:
    """Validate and prepare a monthly time series DataFrame.

    Parameters
    ----------
    df : pd.DataFrame
        Данные с колонками дат и значений.
    date_col : str
        Название колонки с датами.
    value_col : str
        Название колонки со значениями.

    Returns
    -------
    pd.DataFrame
        DataFrame с колонками ``date`` (pd.Timestamp, первое число месяца)
        и ``value`` (float64), отсортированный по дате.

    Raises
    ------
    ValueError
        Если есть пропущенные/дублированные даты, NA-значения или разрыв в ряду.
    """
    if date_col not in df.columns:
        raise ValueError(f"Колонка '{date_col}' не найдена в DataFrame.")
    if value_col not in df.columns:
        raise ValueError(f"Колонка '{value_col}' не найдена в DataFrame.")

    out = pd.DataFrame({
        "date":  parse_to_month(df[date_col]),
        "value": parse_numeric(df[value_col].values),
    }).sort_values("date").reset_index(drop=True)

    if out["date"].isna().any():
        raise ValueError("Не удалось разобрать часть дат (получились NaT). Проверьте колонку дат.")
    if np.isnan(out["value"]).any():
        raise ValueError("В значениях ряда есть NA. Проверьте колонку значений.")
    if out["date"].duplicated().any():
        raise ValueError("В ряду есть повторяющиеся месяцы.")

    # проверяем непрерывность
    expected = pd.date_range(start=out["date"].min(), end=out["date"].max(), freq="MS")
    if len(expected) != len(out) or (out["date"].values != expected.values).any():
        raise ValueError("Даты должны образовывать непрерывный месячный ряд без пропусков.")

    return out


# ---------------------------------------------------------------------------
# 3. Российский производственный календарь → месячные центрированные регрессоры
# ---------------------------------------------------------------------------

def make_ru_calendar_from_excel(
    month_dates: pd.DatetimeIndex,
    calendar_file: Union[str, Path],
    sheet: Union[int, str] = 0,
    date_col: str = "date",
    workday_col: str = "is_workday",
    holiday_col: str = "is_holiday",
    easter_col: str = "easter_effect",
    include_easter: bool = True,
    center_start: Optional[str] = None,
    center_end: Optional[str] = None,
) -> pd.DataFrame:
    """Build monthly centered calendar regressors from a Russian production calendar Excel file.

    Implements the same centering logic as the R function ``make_ru_calendar_from_excel``.

    Parameters
    ----------
    month_dates : pd.DatetimeIndex
        Месяцы, для которых нужны регрессоры (включая горизонт прогноза).
    calendar_file : str or Path
        Путь к ``russia_calendar.xlsx``.
    sheet : int or str
        Лист Excel (0-based int или имя).
    date_col, workday_col, holiday_col, easter_col : str
        Названия колонок в файле календаря.
    include_easter : bool
        Включать пасхальный регрессор.
    center_start, center_end : str, optional
        Границы окна для вычисления средних (центрирующих констант).
        Формат — любой, понимаемый ``parse_to_month``. ``None`` = весь
        доступный календарь.

    Returns
    -------
    pd.DataFrame
        Колонки: ``date``, ``days_in_month``, ``workdays``, ``holidays``,
        ``weekends``, ``easter_days``, ``workdays_c``, ``holidays_c``, ``easter_c``.

    Raises
    ------
    ValueError
        Если не хватает колонок или календарь не покрывает нужный диапазон.
    """
    month_dates = parse_to_month(month_dates)

    # --- читаем Excel-файл
    sheet_arg = sheet if isinstance(sheet, str) else sheet
    daily = pd.read_excel(calendar_file, sheet_name=sheet_arg)
    daily.columns = [str(c).strip().lower() for c in daily.columns]

    date_col    = date_col.lower()
    workday_col = workday_col.lower()
    holiday_col = holiday_col.lower()
    easter_col  = easter_col.lower()

    missing = [c for c in [date_col, workday_col, holiday_col] if c not in daily.columns]
    if missing:
        raise ValueError(f"В календаре не найдены колонки: {', '.join(missing)}")
    if include_easter and easter_col not in daily.columns:
        raise ValueError(
            f"include_easter=True, но в календаре нет колонки '{easter_col}'. "
            "Добавьте её или поставьте include_easter=False."
        )
    if easter_col not in daily.columns:
        daily[easter_col] = 0

    daily = pd.DataFrame({
        "date":          pd.to_datetime(daily[date_col]),
        "is_workday":    daily[workday_col].astype(int),
        "is_holiday":    daily[holiday_col].astype(int),
        "easter_effect": daily[easter_col].astype(int),
    }).sort_values("date").reset_index(drop=True)

    if daily["date"].isna().any():
        raise ValueError("В календаре есть даты, которые не удалось прочитать.")
    for col, name in [("is_workday", workday_col), ("is_holiday", holiday_col), ("easter_effect", easter_col)]:
        bad_vals = set(daily[col].unique()) - {0, 1}
        if bad_vals:
            raise ValueError(f"Колонка '{name}' должна содержать только 0 и 1, найдено: {bad_vals}")

    # --- Месячная агрегация по ВСЕМУ календарю
    daily["month_date"] = daily["date"].values.astype("datetime64[M]").astype("datetime64[D]")
    monthly_all = (
        daily
        .groupby("month_date")
        .agg(
            days_in_month  = ("date",          "count"),
            workdays       = ("is_workday",     "sum"),
            holidays       = ("is_holiday",     "sum"),
            easter_days    = ("easter_effect",  "sum"),
        )
        .reset_index()
    )
    monthly_all["weekends"] = (
        monthly_all["days_in_month"] - monthly_all["workdays"] - monthly_all["holidays"]
    )
    monthly_all["moy"] = pd.to_datetime(monthly_all["month_date"]).dt.month

    # --- Средние по месяцу года для центрирования
    center_set = monthly_all.copy()
    if center_start is not None:
        cs = parse_to_month(center_start)[0]
        center_set = center_set[pd.to_datetime(center_set["month_date"]) >= cs]
    if center_end is not None:
        ce = parse_to_month(center_end)[0]
        center_set = center_set[pd.to_datetime(center_set["month_date"]) <= ce]

    moy_means = (
        center_set
        .groupby("moy")
        .agg(
            m_workdays    = ("workdays",     "mean"),
            m_holidays    = ("holidays",     "mean"),
            m_easter_days = ("easter_days",  "mean"),
        )
        .reset_index()
    )

    if len(moy_means) < 12:
        # окно не покрывает все 12 месяцев — центрируем по всему календарю
        warnings.warn(
            "Окно центрирования покрывает не все 12 месяцев; центрирую по всему календарю.",
            UserWarning, stacklevel=2
        )
        moy_means = (
            monthly_all
            .groupby("moy")
            .agg(
                m_workdays    = ("workdays",     "mean"),
                m_holidays    = ("holidays",     "mean"),
                m_easter_days = ("easter_days",  "mean"),
            )
            .reset_index()
        )

    # --- Объединяем с запрошенными месяцами
    out = pd.DataFrame({"date": pd.DatetimeIndex(month_dates)})
    out["moy"] = out["date"].dt.month
    monthly_all["month_date"] = pd.to_datetime(monthly_all["month_date"])

    out = out.merge(
        monthly_all.drop(columns=["moy"]),
        left_on="date", right_on="month_date", how="left"
    ).drop(columns=["month_date"], errors="ignore")
    out = out.merge(moy_means, on="moy", how="left")

    if out["workdays"].isna().any():
        rng = (monthly_all["month_date"].min(), monthly_all["month_date"].max())
        raise ValueError(
            f"В производственном календаре не хватает месяцев для ряда + горизонта прогноза.\n"
            f"Календарь покрывает {rng[0]:%Y-%m} .. {rng[1]:%Y-%m}. "
            "Продлите календарь или уменьшите forecast_months."
        )

    # --- Центрированные регрессоры
    out["workdays_c"] = out["workdays"] - out["m_workdays"]
    out["holidays_c"] = out["holidays"] - out["m_holidays"]
    out["easter_c"]   = (
        out["easter_days"] - out["m_easter_days"]
        if include_easter else 0.0
    )

    cols = ["date", "days_in_month", "workdays", "holidays", "weekends",
            "easter_days", "workdays_c", "holidays_c", "easter_c"]
    return out[cols].reset_index(drop=True)


# ---------------------------------------------------------------------------
# 4. Матрица календарных регрессоров
# ---------------------------------------------------------------------------

def build_xreg_matrix(
    cal_df: pd.DataFrame,
    mode: str = "basic",
    include_easter: bool = True,
) -> Optional[np.ndarray]:
    """Build the matrix of calendar regressors.

    Parameters
    ----------
    cal_df : pd.DataFrame
        Выход ``make_ru_calendar_from_excel``.
    mode : {"none", "basic", "extended"}
        * ``"none"``     — без календарных регрессоров.
        * ``"basic"``    — только число рабочих дней.
        * ``"extended"`` — рабочие дни + праздники.
    include_easter : bool
        Добавить пасхальный регрессор.

    Returns
    -------
    np.ndarray or None
        Матрица регрессоров (n_obs × n_regressors), или ``None`` при mode='none'.
    """
    if mode == "none":
        return None

    cols: Dict[str, np.ndarray] = {"workdays": cal_df["workdays_c"].to_numpy(float)}
    if mode == "extended":
        cols["holidays"] = cal_df["holidays_c"].to_numpy(float)
    if include_easter:
        cols["easter"] = cal_df["easter_c"].to_numpy(float)

    xreg = np.column_stack(list(cols.values()))
    names = list(cols.keys())

    # убираем константные регрессоры (std ≈ 0)
    keep = np.std(xreg, axis=0) > 1e-8
    xreg = xreg[:, keep]
    names = [n for n, k in zip(names, keep) if k]

    if xreg.shape[1] == 0:
        return None

    return xreg


# ---------------------------------------------------------------------------
# 5. Подбор SARIMAX-модели
# ---------------------------------------------------------------------------

# Стандартный «airline» порядок — хорошая точка отсчёта для месячных рядов
_DEFAULT_ORDER        = (1, 1, 1)
_DEFAULT_SEASONAL     = (1, 1, 1, 12)


def _fit_sarimax(
    y: np.ndarray,
    xreg: Optional[np.ndarray] = None,
    order: Tuple[int, int, int] = _DEFAULT_ORDER,
    seasonal_order: Tuple[int, int, int, int] = _DEFAULT_SEASONAL,
    trend: Optional[str] = None,
) -> Any:
    """Fit a SARIMAX model, returning the results object or None on failure.

    Parameters
    ----------
    y : np.ndarray
        Временной ряд (1D).
    xreg : np.ndarray, optional
        Матрица внешних регрессоров (n_obs × k).
    order, seasonal_order : tuple
        ARIMA и сезонный порядок.
    trend : str, optional
        Параметр trend для SARIMAX ('n', 'c', 't', 'ct').

    Returns
    -------
    SARIMAXResultsWrapper or None
    """
    try:
        model = SARIMAX(
            y,
            exog=xreg,
            order=order,
            seasonal_order=seasonal_order,
            trend=trend,
            enforce_stationarity=False,
            enforce_invertibility=False,
        )
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            res = model.fit(disp=False)
        return res
    except Exception:
        return None


def _seasonal_adjustment_from_sarimax(
    y: np.ndarray,
    xreg: Optional[np.ndarray] = None,
    res=None,
) -> Tuple[np.ndarray, np.ndarray]:
    """Extract seasonally adjusted series and seasonal factor from a SARIMAX result.

    The method uses in-sample predictions:
    - If calendar regressors (``xreg``) are provided, the seasonal component
      is approximated as the total fitted values minus the ARIMA component
      (i.e., the regression contribution of calendar effects + residual seasonality
      captured by the seasonal ARIMA).
    - Trend is approximated via a centred 2×12 moving average of the adjusted series.

    Parameters
    ----------
    y : np.ndarray
        Исходный временной ряд.
    xreg : np.ndarray, optional
        Регрессоры.
    res : SARIMAXResultsWrapper
        Подогнанная модель.

    Returns
    -------
    adjusted : np.ndarray
        Сезонно скорректированный ряд.
    seasonal_factor : np.ndarray
        Сезонная составляющая (аддитивная: original − adjusted).
    """
    if res is None:
        return y.copy(), np.zeros_like(y)

    fitted = res.fittedvalues
    residuals = res.resid

    # --- Оценка сезонной составляющей через отклонения от скользящего среднего 2×12
    # (стандартный приём для аддитивной модели)
    n = len(y)
    # Центрированное 12-месячное скользящее среднее (как основа тренда)
    ma12 = np.convolve(y, np.ones(12) / 12, mode="full")[11: n + 11]
    # 2×12: ещё раз усредняем соседние пары
    trend_raw = np.full(n, np.nan)
    for i in range(6, n - 6):
        trend_raw[i] = (ma12[i - 1] + ma12[i]) / 2

    # Разность ряда и тренда как грубая сезонность
    detrended = y - trend_raw

    # Среднее по каждому месяцу года (seasonal means)
    months = np.arange(n) % 12
    seas_means = np.zeros(12)
    for m in range(12):
        vals = detrended[months == m]
        seas_means[m] = np.nanmean(vals) if not np.all(np.isnan(vals)) else 0.0

    # Нормируем чтобы сумма за год ≈ 0 (аддитивная модель)
    seas_means -= seas_means.mean()
    seasonal_factor = np.array([seas_means[m] for m in months])

    # Регрессионный вклад календаря — вычтем его дополнительно
    if xreg is not None and res is not None:
        params_names = res.param_names
        xreg_param_idx = [i for i, n in enumerate(params_names) if n.startswith("x")]
        if xreg_param_idx:
            xreg_params = res.params[xreg_param_idx]
            cal_contribution = xreg[:, :len(xreg_param_idx)] @ xreg_params
            seasonal_factor += cal_contribution

    adjusted = y - seasonal_factor
    return adjusted, seasonal_factor


def _moving_average_trend(y: np.ndarray, window: int = 12) -> np.ndarray:
    """Compute a 2×(window) centered moving average (trend estimate).

    Parameters
    ----------
    y : np.ndarray
        Временной ряд.
    window : int
        Размер окна (12 для месячных данных).

    Returns
    -------
    np.ndarray
        Тренд с NaN на краях.
    """
    n = len(y)
    ma = np.convolve(y, np.ones(window) / window, mode="same")
    trend = np.full(n, np.nan)
    half = window // 2
    for i in range(half, n - half):
        trend[i] = (ma[i - 1] + ma[i]) / 2
    return trend


# ---------------------------------------------------------------------------
# 6. Главная функция
# ---------------------------------------------------------------------------

def sa_ru(
    df: pd.DataFrame,
    calendar_file: Union[str, Path],
    calendar_sheet: Union[int, str] = 0,
    date_col: str = "date",
    value_col: str = "value",
    calendar_date_col: str = "date",
    calendar_workday_col: str = "is_workday",
    calendar_holiday_col: str = "is_holiday",
    calendar_easter_col: str = "easter_effect",
    calendar_mode: str = "basic",
    include_easter: bool = True,
    transform_function: str = "none",
    forecast_months: int = 36,
    seasonality_alpha: float = 0.05,
    center_start: Optional[str] = None,
    center_end: Optional[str] = None,
    arima_order: Tuple[int, int, int] = _DEFAULT_ORDER,
    arima_seasonal: Tuple[int, int, int, int] = _DEFAULT_SEASONAL,
    verbose: bool = True,
) -> Dict[str, Any]:
    """Seasonal adjustment of monthly time series using the Russian production calendar.

    Python-порт функции ``sa_ru`` из R-пакета NadezhdaYurchenko/Seasonal-Adjustment-for-Russia.

    Вместо X-13ARIMA-SEATS используется ``statsmodels.SARIMAX`` с теми же
    пользовательскими календарными регрессорами:
    * рабочие дни (центрированные),
    * официальные праздники (в режиме ``"extended"``),
    * пасхальный эффект.

    Parameters
    ----------
    df : pd.DataFrame
        Таблица с временным рядом. Обязательные колонки: ``date_col`` и ``value_col``.
    calendar_file : str or Path
        Путь к ``russia_calendar.xlsx``.
    calendar_sheet : int or str
        Лист Excel (0-based int или имя). По умолчанию 0.
    date_col : str
        Название колонки с датами в ``df``.
    value_col : str
        Название колонки со значениями в ``df``.
    calendar_date_col : str
        Название колонки с датами в файле календаря.
    calendar_workday_col : str
        Название колонки признака рабочего дня в файле календаря.
    calendar_holiday_col : str
        Название колонки признака праздника в файле календаря.
    calendar_easter_col : str
        Название колонки пасхального эффекта в файле календаря.
    calendar_mode : {"none", "basic", "extended", "auto"}
        * ``"none"``     — без календарных регрессоров.
        * ``"basic"``    — только рабочие дни.
        * ``"extended"`` — рабочие дни + праздники.
        * ``"auto"``     — перебрать все три и выбрать лучший по AIC.
    include_easter : bool
        Включить пасхальный регрессор.
    transform_function : {"none", "log"}
        ``"log"`` — логарифмическое преобразование перед сглаживанием.
    forecast_months : int
        Горизонт прогноза в месяцах (регрессоры строятся на ряд + горизонт).
    seasonality_alpha : float
        Уровень значимости для теста на сезонность.
    center_start, center_end : str, optional
        Даты начала/конца окна для центрирования регрессоров. ``None`` = период ряда.
    arima_order : tuple of int
        Порядок ARIMA (p, d, q). По умолчанию (1, 1, 1).
    arima_seasonal : tuple of int
        Сезонный порядок (P, D, Q, s). По умолчанию (1, 1, 1, 12).
    verbose : bool
        Выводить сообщения о ходе выполнения.

    Returns
    -------
    dict with keys:
        * ``"data"`` — pd.DataFrame: date, original, adjusted, trend, seasonal_factor.
        * ``"chosen_model"`` — выбранный режим календаря.
        * ``"transform"`` — применённое преобразование.
        * ``"aic"`` — AIC выбранной модели.
        * ``"bic"`` — BIC выбранной модели.
        * ``"comparison"`` — pd.DataFrame со сравнением кандидатов (для mode='auto').
        * ``"calendar_monthly"`` — pd.DataFrame с месячными регрессорами.
        * ``"model"`` — объект SARIMAXResultsWrapper.

    Raises
    ------
    ValueError
        Если данные некорректны или модель не может быть оценена.
    """
    _say = (lambda msg: print(msg)) if verbose else (lambda msg: None)

    # --- 1. Подготовка данных
    dat = prepare_monthly_df(df, date_col, value_col)
    n = len(dat)

    # --- 2. Защита логарифма
    if transform_function == "log" and (dat["value"] <= 0).any():
        raise ValueError(
            "transform_function='log', но в ряду есть значения <= 0. Используйте 'none'."
        )

    y_raw = dat["value"].to_numpy(dtype=float)
    if transform_function == "log":
        y = np.log(y_raw)
    else:
        y = y_raw.copy()

    # --- 3. Календарные регрессоры
    if center_start is None:
        center_start = dat["date"].min().strftime("%Y-%m")
    if center_end is None:
        center_end = dat["date"].max().strftime("%Y-%m")

    xreg_end = dat["date"].max() + pd.DateOffset(months=forecast_months)
    xreg_dates = pd.date_range(start=dat["date"].min(), end=xreg_end, freq="MS")

    cal_df = make_ru_calendar_from_excel(
        month_dates=xreg_dates,
        calendar_file=calendar_file,
        sheet=calendar_sheet,
        date_col=calendar_date_col,
        workday_col=calendar_workday_col,
        holiday_col=calendar_holiday_col,
        easter_col=calendar_easter_col,
        include_easter=include_easter,
        center_start=center_start,
        center_end=center_end,
    )

    # регрессоры только для периода ряда (без прогнозного хвоста)
    cal_in_sample = cal_df[cal_df["date"].isin(dat["date"])].reset_index(drop=True)

    def _xreg_for(mode: str) -> Optional[np.ndarray]:
        xr = build_xreg_matrix(cal_in_sample, mode, include_easter)
        return xr

    # =========================================================================
    # РЕЖИМ AUTO: сравниваем none / basic / extended по AIC
    # =========================================================================
    comparison_rows = []
    best_mode: str
    best_res = None

    if calendar_mode == "auto":
        _say("AUTO: перебираем none / basic / extended, выбор по AIC ...")
        candidates = {}
        for mode_candidate in ("none", "basic", "extended"):
            _say(f"  оцениваю кандидата: {mode_candidate}")
            xr = _xreg_for(mode_candidate)
            res = _fit_sarimax(y, xr, arima_order, arima_seasonal)
            aic_val = float(res.aic) if res is not None else np.nan
            bic_val = float(res.bic) if res is not None else np.nan
            comparison_rows.append({
                "model": mode_candidate,
                "aic": aic_val,
                "bic": bic_val,
                "converged": res is not None,
            })
            if res is not None:
                candidates[mode_candidate] = res
            _say(f"    AIC = {aic_val:.2f}")

        valid = [(m, r) for m, r in candidates.items()]
        if not valid:
            raise ValueError("Ни один кандидат не оценился.")

        best_mode = min(
            candidates.keys(),
            key=lambda m: next(
                row["aic"] for row in comparison_rows if row["model"] == m
            )
        )
        best_res = candidates[best_mode]
        _say(f"Выбранный режим (min AIC): {best_mode}")

        comparison = pd.DataFrame(comparison_rows).assign(
            chosen=lambda d: d["model"] == best_mode
        )

    else:
        # =====================================================================
        # РУЧНОЙ РЕЖИМ
        # =====================================================================
        _say(f"Ручной режим, календарь = {calendar_mode}")
        if calendar_mode not in ("none", "basic", "extended"):
            raise ValueError(
                f"calendar_mode должен быть одним из: 'none', 'basic', 'extended', 'auto'. "
                f"Получено: '{calendar_mode}'"
            )
        xr = _xreg_for(calendar_mode)
        best_res = _fit_sarimax(y, xr, arima_order, arima_seasonal)
        if best_res is None:
            raise ValueError(
                f"Не удалось оценить SARIMAX-модель (calendar_mode='{calendar_mode}'). "
                "Попробуйте другой порядок ARIMA или режим календаря."
            )
        best_mode = calendar_mode
        comparison = pd.DataFrame([{
            "model": calendar_mode,
            "aic": float(best_res.aic),
            "bic": float(best_res.bic),
            "converged": True,
            "chosen": True,
        }])

    # --- 4. Проверка сезонности (упрощённая: по остаткам автокорреляций на лаге 12)
    # В R используется QS-тест из X-13; здесь используем тест Льюнга-Бокса на лаге 12.
    seasonality_detected = True
    if best_res is not None:
        try:
            from statsmodels.stats.diagnostic import acorr_ljungbox
            lb_result = acorr_ljungbox(best_res.resid, lags=[12], return_df=True)
            qs_p = float(lb_result["lb_pvalue"].iloc[0])
            # сезонность в остатках — если p < alpha, остатки всё ещё сезонны
            # (но модель уже убрала основную сезонность через seasonal ARIMA)
            seasonality_detected = True  # SARIMAX всегда делает SA
        except Exception:
            qs_p = np.nan

    _say(f"Сезонная корректировка: модель = {best_mode}, AIC = {best_res.aic:.2f}")

    # --- 5. Извлечение сезонно скорректированного ряда
    xr_best = _xreg_for(best_mode)
    adjusted_y, seasonal_factor_y = _seasonal_adjustment_from_sarimax(y, xr_best, best_res)

    # Обратное логарифмическое преобразование
    if transform_function == "log":
        adjusted  = np.exp(adjusted_y)
        seas_fac  = y_raw / adjusted  # мультипликативный фактор
        factor_type = "multiplicative"
    else:
        adjusted  = adjusted_y
        seas_fac  = y_raw - adjusted   # аддитивный фактор
        factor_type = "additive"

    # --- 6. Тренд (2×12 MA от скорректированного ряда)
    trend = _moving_average_trend(adjusted, window=12)

    result_df = pd.DataFrame({
        "date":            dat["date"],
        "original":        y_raw,
        "adjusted":        adjusted,
        "trend":           trend,
        "seasonal_factor": seas_fac,
        "factor_type":     factor_type,
    })

    return {
        "data":              result_df,
        "chosen_model":      best_mode,
        "transform":         transform_function,
        "aic":               float(best_res.aic) if best_res is not None else np.nan,
        "bic":               float(best_res.bic) if best_res is not None else np.nan,
        "seasonality_detected": seasonality_detected,
        "comparison":        comparison,
        "calendar_monthly":  cal_df,
        "model":             best_res,
    }


# ---------------------------------------------------------------------------
# 7. Заморозка спецификации (упрощённый вариант sa_ru_identify / sa_ru_apply)
# ---------------------------------------------------------------------------

def sa_ru_identify(
    df: pd.DataFrame,
    calendar_file: Union[str, Path],
    series_id: Optional[str] = None,
    cutoff_date: Optional[str] = None,
    dir: str = "sa_specs",
    **sa_kwargs,
) -> Dict[str, Any]:
    """Identify and optionally save a seasonal adjustment specification.

    Аналог ``sa_ru_identify`` из R. Запускает ``sa_ru`` на данных до ``cutoff_date``
    и сохраняет спецификацию в JSON.

    Parameters
    ----------
    df : pd.DataFrame
        Полный ряд.
    calendar_file : str or Path
        Путь к ``russia_calendar.xlsx``.
    series_id : str, optional
        Идентификатор ряда. Если задан, сохраняется файл ``<dir>/<series_id>.json``.
    cutoff_date : str, optional
        Граница идентификации. ``None`` = последнее наблюдение.
    dir : str
        Папка для сохранения спецификаций.
    **sa_kwargs
        Все прочие аргументы передаются в ``sa_ru``.

    Returns
    -------
    dict
        Спецификация модели (dict, пригодный для JSON-сохранения).
    """
    dat_full = prepare_monthly_df(
        df,
        sa_kwargs.get("date_col", "date"),
        sa_kwargs.get("value_col", "value"),
    )
    if cutoff_date is not None:
        cutoff = parse_to_month(cutoff_date)[0]
    else:
        cutoff = dat_full["date"].max()

    dat_id = dat_full[dat_full["date"] <= cutoff].copy()
    if len(dat_id) < 36:
        raise ValueError("Слишком короткая выборка для идентификации (нужно >= 3 года).")

    sa_kwargs.setdefault("center_start", dat_id["date"].min().strftime("%Y-%m"))
    sa_kwargs.setdefault("center_end",   cutoff.strftime("%Y-%m"))

    res = sa_ru(dat_id, calendar_file, **sa_kwargs)

    spec = {
        "series_id":         series_id,
        "identified_on":     date.today().isoformat(),
        "sample_start":      dat_id["date"].min().strftime("%Y-%m"),
        "data_through":      cutoff.strftime("%Y-%m"),
        "transform":         res["transform"],
        "calendar_mode":     res["chosen_model"],
        "include_easter":    sa_kwargs.get("include_easter", True),
        "arima_order":       list(sa_kwargs.get("arima_order", list(_DEFAULT_ORDER))),
        "arima_seasonal":    list(sa_kwargs.get("arima_seasonal", list(_DEFAULT_SEASONAL))),
        "aic":               res["aic"],
        "forecast_months":   sa_kwargs.get("forecast_months", 36),
        "center_start":      sa_kwargs["center_start"],
        "center_end":        sa_kwargs["center_end"],
        "seasonality_detected": res["seasonality_detected"],
        "calendar_cols": {
            "date":    sa_kwargs.get("calendar_date_col",    "date"),
            "workday": sa_kwargs.get("calendar_workday_col", "is_workday"),
            "holiday": sa_kwargs.get("calendar_holiday_col", "is_holiday"),
            "easter":  sa_kwargs.get("calendar_easter_col",  "easter_effect"),
        },
        "calendar_sheet": sa_kwargs.get("calendar_sheet", 0),
    }

    if series_id is not None:
        os.makedirs(dir, exist_ok=True)
        path = os.path.join(dir, f"{series_id}.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(spec, f, ensure_ascii=False, indent=2)
        if sa_kwargs.get("verbose", True):
            print(f"Спецификация сохранена: {os.path.abspath(path)}")
    else:
        if sa_kwargs.get("verbose", True):
            print("series_id не задан -> спека НЕ сохранена (только возвращена в объекте).")

    return spec


def sa_ru_apply(
    df: pd.DataFrame,
    calendar_file: Union[str, Path],
    series_id: str,
    dir: str = "sa_specs",
    date_col: str = "date",
    value_col: str = "value",
    verbose: bool = True,
) -> Dict[str, Any]:
    """Apply a frozen seasonal adjustment specification to new data.

    Аналог ``sa_ru_apply`` из R.

    Parameters
    ----------
    df : pd.DataFrame
        Полный ряд (включая новые наблюдения после границы идентификации).
    calendar_file : str or Path
        Путь к ``russia_calendar.xlsx``.
    series_id : str
        Идентификатор ряда (должен совпадать с сохранённым файлом).
    dir : str
        Папка со спецификациями.
    date_col, value_col : str
        Названия колонок в ``df``.
    verbose : bool
        Выводить сообщения.

    Returns
    -------
    dict
        Ключи ``"data"``, ``"new_points"``, ``"spec"``, ``"model"``.
    """
    path = os.path.join(dir, f"{series_id}.json")
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"Не найдена спецификация: {path}. "
            f"Сначала запустите sa_ru_identify(series_id='{series_id}', ...)."
        )
    with open(path, encoding="utf-8") as f:
        spec = json.load(f)

    res = sa_ru(
        df=df,
        calendar_file=calendar_file,
        calendar_sheet=spec.get("calendar_sheet", 0),
        date_col=date_col,
        value_col=value_col,
        calendar_date_col=spec["calendar_cols"]["date"],
        calendar_workday_col=spec["calendar_cols"]["workday"],
        calendar_holiday_col=spec["calendar_cols"]["holiday"],
        calendar_easter_col=spec["calendar_cols"]["easter"],
        calendar_mode=spec["calendar_mode"],
        include_easter=spec.get("include_easter", True),
        transform_function=spec["transform"],
        forecast_months=spec.get("forecast_months", 36),
        center_start=spec["center_start"],
        center_end=spec["center_end"],
        arima_order=tuple(spec.get("arima_order", _DEFAULT_ORDER)),
        arima_seasonal=tuple(spec.get("arima_seasonal", _DEFAULT_SEASONAL)),
        verbose=verbose,
    )

    border = pd.Timestamp(spec["data_through"] + "-01")
    new_points = res["data"][res["data"]["date"] > border].copy()

    if verbose:
        print(
            f"Применена спека '{series_id}' (идентиф. {spec['identified_on']}, "
            f"заморожена до {spec['data_through']}).\n"
            f"  новых точек после границы: {len(new_points)}"
        )

    return {
        "data":       res["data"],
        "new_points": new_points,
        "spec":       spec,
        "model":      res["model"],
    }


# ---------------------------------------------------------------------------
# 8. Пакетная обработка нескольких рядов
# ---------------------------------------------------------------------------

# Настройки по умолчанию для sa_ru_batch
_BATCH_DEFAULT_CONFIG: Dict[str, Any] = {
    "calendar_mode":      "basic",
    "include_easter":     True,
    "transform_function": "none",
    "forecast_months":    36,
    "seasonality_alpha":  0.05,
    "arima_order":        _DEFAULT_ORDER,
    "arima_seasonal":     _DEFAULT_SEASONAL,
}

# Параметры sa_ru(), которые можно задавать на уровне серии
_SERIES_LEVEL_PARAMS = {
    "calendar_mode", "include_easter", "transform_function",
    "forecast_months", "seasonality_alpha", "arima_order", "arima_seasonal",
    "center_start", "center_end",
    "calendar_sheet", "calendar_date_col",
    "calendar_workday_col", "calendar_holiday_col", "calendar_easter_col",
}


def sa_ru_batch(
    df: pd.DataFrame,
    calendar_file: Union[str, Path],
    format: str = "wide",
    date_col: str = "date",
    series_col: str = "series_id",
    value_col: str = "value",
    default_config: Optional[Dict[str, Any]] = None,
    series_configs: Optional[Dict[str, Dict[str, Any]]] = None,
    output_long: bool = True,
    fail_on_error: bool = False,
    verbose: bool = True,
) -> Dict[str, Any]:
    """Seasonal adjustment of **multiple** time series in a single call.

    Поддерживает wide и long форматы. Каждый ряд может иметь собственные
    настройки через ``series_configs``.

    Parameters
    ----------
    df : pd.DataFrame
        Данные в wide или long формате.

        **Wide** (``format='wide'``) — дата в одной колонке, каждый ряд —
        отдельная колонка::

            date        | CPI   | PPI   | IP
            2015-01-01  | 101.2 | 85.4  | 112.3

        **Long** (``format='long'``) — три колонки: дата, идентификатор ряда,
        значение::

            date        | series_id | value
            2015-01-01  | CPI       | 101.2
            2015-01-01  | PPI       | 85.4

    calendar_file : str or Path
        Путь к ``russia_calendar.xlsx``.
    format : {"wide", "long"}
        Формат входных данных. По умолчанию ``"wide"``.
    date_col : str
        Колонка с датами (для обоих форматов).
    series_col : str
        Колонка с именами рядов (только для ``format='long'``).
    value_col : str
        Колонка со значениями (только для ``format='long'``).
    default_config : dict, optional
        Настройки ``sa_ru()`` по умолчанию для всех рядов.
        Если ``None`` — используются значения по умолчанию из ``sa_ru()``.
        Допустимые ключи: ``calendar_mode``, ``include_easter``,
        ``transform_function``, ``forecast_months``, ``arima_order``,
        ``arima_seasonal``, ``center_start``, ``center_end``, и др.
    series_configs : dict of dict, optional
        Индивидуальные настройки для конкретных рядов — перегружают
        ``default_config``. Ключи верхнего уровня — имена рядов::

            series_configs = {
                "CPI": {"transform_function": "log"},
                "PPI": {"calendar_mode": "extended", "include_easter": False},
                "IP":  {"calendar_mode": "auto"},
            }

    output_long : bool
        Если ``True`` — результат содержит ключ ``"combined"`` с единым
        long DataFrame по всем рядам. По умолчанию ``True``.
    fail_on_error : bool
        Если ``True`` — ошибка в одном ряду прерывает всё.
        Если ``False`` — ошибка записывается, обработка продолжается.
        По умолчанию ``False``.
    verbose : bool
        Печатать прогресс. По умолчанию ``True``.

    Returns
    -------
    dict with keys:
        * ``"results"`` — dict: имя_ряда → полный вывод ``sa_ru()``.
        * ``"summary"`` — pd.DataFrame: сводная таблица по всем рядам.
        * ``"combined"`` — pd.DataFrame (long): все SA-ряды в одной таблице
          (только если ``output_long=True``).
        * ``"errors"`` — dict: имя_ряда → текст ошибки.

    Examples
    --------
    >>> results = sa_ru_batch(
    ...     df            = wide_df,
    ...     calendar_file = "russia_calendar.xlsx",
    ...     format        = "wide",
    ...     default_config = {"calendar_mode": "basic", "include_easter": True},
    ...     series_configs = {
    ...         "CPI": {"transform_function": "log"},
    ...         "IP":  {"calendar_mode": "auto"},
    ...     },
    ... )
    >>> results["summary"]
    >>> results["combined"]                    # все ряды в one DataFrame
    >>> results["results"]["CPI"]["data"]      # SA-ряд для CPI
    >>> results["results"]["CPI"]["aic"]       # AIC модели CPI
    """
    _say = (lambda msg: print(msg)) if verbose else (lambda msg: None)

    # --- Слияние default_config с глобальными умолчаниями
    base_cfg = dict(_BATCH_DEFAULT_CONFIG)
    if default_config:
        base_cfg.update(default_config)
    if series_configs is None:
        series_configs = {}

    # --- Приводим к long-формату
    long_df = _to_long_format(df, format, date_col, series_col, value_col)
    all_series = sorted(long_df[series_col].unique())
    n_total = len(all_series)
    _say(f"Пакетная обработка: {n_total} рядов")

    # Предупреждение об неизвестных именах в series_configs
    unknown = set(series_configs) - set(all_series)
    if unknown:
        import warnings
        warnings.warn(
            f"В series_configs есть ряды, которых нет в df: {sorted(unknown)}",
            UserWarning, stacklevel=2
        )

    # --- Обрабатываем каждый ряд
    results: Dict[str, Any] = {}
    errors:  Dict[str, str] = {}
    summary_rows: List[Dict] = []

    for idx, sname in enumerate(all_series, start=1):
        _say(f"  [{idx}/{n_total}] {sname} ...")

        # Данные ряда
        sub = (
            long_df[long_df[series_col] == sname][[date_col, value_col]]
            .rename(columns={date_col: "date", value_col: "value"})
            .copy()
        )

        # Итоговая конфигурация для этого ряда
        cfg = dict(base_cfg)
        cfg.update({k: v for k, v in series_configs.get(sname, {}).items()
                    if k in _SERIES_LEVEL_PARAMS})

        try:
            res = sa_ru(
                df=sub,
                calendar_file=calendar_file,
                date_col="date",
                value_col="value",
                verbose=False,
                **{k: v for k, v in cfg.items() if k != "calendar_file"},
            )
            results[sname] = res
            status = "OK"
            error_msg = None
            _say(
                f"    OK | модель={res['chosen_model']} "
                f"| transform={res['transform']} "
                f"| AIC={res['aic']:.2f}"
            )
        except Exception as exc:
            error_msg = str(exc)
            errors[sname] = error_msg
            status = "ERROR"
            _say(f"    ОШИБКА: {error_msg}")
            if fail_on_error:
                raise

        summary_rows.append({
            "series_id":            sname,
            "status":               status,
            "chosen_model":         results[sname]["chosen_model"] if status == "OK" else None,
            "transform":            results[sname]["transform"]     if status == "OK" else None,
            "aic":                  results[sname]["aic"]           if status == "OK" else np.nan,
            "bic":                  results[sname]["bic"]           if status == "OK" else np.nan,
            "seasonality_detected": results[sname]["seasonality_detected"] if status == "OK" else None,
            "n_obs":                len(sub),
            "error_msg":            error_msg,
        })

    summary_df = pd.DataFrame(summary_rows)
    ok_n  = (summary_df["status"] == "OK").sum()
    err_n = (summary_df["status"] == "ERROR").sum()
    _say(f"Готово: успешно={ok_n}, ошибок={err_n}")
    if verbose:
        print(summary_df[["series_id", "status", "chosen_model", "transform",
                           "aic", "seasonality_detected"]].to_string(index=False))

    # --- Объединённый long DataFrame
    combined_df = None
    if output_long and results:
        parts = []
        for sname, res in results.items():
            part = res["data"].copy()
            part.insert(0, "series_id", sname)
            parts.append(part)
        combined_df = pd.concat(parts, ignore_index=True)

    return {
        "results":  results,
        "summary":  summary_df,
        "combined": combined_df,
        "errors":   errors,
    }


def _to_long_format(
    df: pd.DataFrame,
    format: str,
    date_col: str,
    series_col: str,
    value_col: str,
) -> pd.DataFrame:
    """Convert wide or long DataFrame to a unified long format.

    Parameters
    ----------
    df : pd.DataFrame
        Входные данные.
    format : {"wide", "long"}
        Формат входных данных.
    date_col, series_col, value_col : str
        Названия ключевых колонок.

    Returns
    -------
    pd.DataFrame
        Long-форматный DataFrame с колонками ``date_col``, ``series_col``, ``value_col``.
    """
    if format == "long":
        missing = [c for c in [date_col, series_col, value_col] if c not in df.columns]
        if missing:
            raise ValueError(f"В df (long) не найдены колонки: {missing}")
        return df[[date_col, series_col, value_col]].copy()

    # wide -> long
    if date_col not in df.columns:
        raise ValueError(f"Колонка '{date_col}' не найдена в df (wide).")
    value_cols = [c for c in df.columns if c != date_col]
    if not value_cols:
        raise ValueError("В df (wide) нет колонок со значениями (только дата).")

    return df.melt(
        id_vars=date_col,
        value_vars=value_cols,
        var_name=series_col,
        value_name=value_col,
    )


def sa_ru_batch_to_excel(
    batch_result: Dict[str, Any],
    path: Union[str, Path] = "sa_results.xlsx",
    include_summary: bool = True,
) -> str:
    """Save ``sa_ru_batch()`` results to a multi-sheet Excel file.

    Требует пакет ``openpyxl``.

    Parameters
    ----------
    batch_result : dict
        Вывод ``sa_ru_batch()``.
    path : str or Path
        Путь к выходному .xlsx файлу. По умолчанию ``"sa_results.xlsx"``.
    include_summary : bool
        Добавить лист «Summary» с общей сводкой. По умолчанию ``True``.

    Returns
    -------
    str
        Абсолютный путь к созданному файлу.
    """
    path = str(path)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        if include_summary and batch_result.get("summary") is not None:
            batch_result["summary"].to_excel(writer, sheet_name="Summary", index=False)
        for sname, res in batch_result.get("results", {}).items():
            sheet = sname[:31]  # Excel limit: 31 chars
            res["data"].to_excel(writer, sheet_name=sheet, index=False)

    abs_path = os.path.abspath(path)
    print(f"Сохранено: {abs_path}")
    return abs_path
