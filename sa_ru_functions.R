# ============================================================================
# КОД ФУНКЦИИ КАСТОМНОГО СЕЗОННОГО СГЛАЖИВАНИЯ
# ============================================================================

library(seasonal)
library(readxl)
library(lubridate)
library(dplyr)
library(tibble)
library(stringr)

# ----------------------------------------------------------------------------
# 1. Разбор дат из данных
# ----------------------------------------------------------------------------

parse_to_month <- function(x) {
  # Приводит вход (Date / POSIXct / Excel-serial / текст) к первому числу месяца.
  # Понимает: "2015M1", "2015M01", "2015-01", "2015/01", "2015-01-01", "01.2015", "1/2015"

  to_first <- function(d) as.Date(format(d, "%Y-%m-01"))

  if (inherits(x, "Date"))    return(to_first(x))
  if (inherits(x, "POSIXct")) return(to_first(as.Date(x)))

  if (is.numeric(x)) {
    # числовой серийник Excel (origin 1899-12-30)
    return(to_first(as.Date(x, origin = "1899-12-30")))
  }

  x_chr <- stringr::str_trim(as.character(x))

  parse_one <- function(s) {
    if (is.na(s) || s == "") return(as.Date(NA))

    # "2015M1" / "2015m01"
    m <- stringr::str_match(s, "^(\\d{4})[Mm](\\d{1,2})$")
    if (!is.na(m[1, 1])) {
      return(as.Date(sprintf("%s-%02d-01", m[1, 2], as.integer(m[1, 3]))))
    }

    # "2015-01" / "2015/01" / "2015.01"
    m <- stringr::str_match(s, "^(\\d{4})[-/.](\\d{1,2})$")
    if (!is.na(m[1, 1])) {
      return(as.Date(sprintf("%s-%02d-01", m[1, 2], as.integer(m[1, 3]))))
    }

    # "01.2015" / "1/2015" / "01-2015"
    m <- stringr::str_match(s, "^(\\d{1,2})[-/.](\\d{4})$")
    if (!is.na(m[1, 1])) {
      return(as.Date(sprintf("%s-%02d-01", m[1, 3], as.integer(m[1, 2]))))
    }

    # полноценная дата "2015-01-01", "2015/01/15", ...
    d <- suppressWarnings(as.Date(s))
    if (!is.na(d)) return(as.Date(format(d, "%Y-%m-01")))

    as.Date(NA)
  }

  out <- as.Date(vapply(x_chr, function(s) as.character(parse_one(s)), character(1)))
  out
}

# Числовой парсер, устойчивый к запятой-разделителю
parse_numeric <- function(x) {
  if (is.numeric(x)) return(as.numeric(x))
  v <- suppressWarnings(as.numeric(as.character(x)))
  bad <- is.na(v) & !is.na(x) & stringr::str_detect(as.character(x), ",")
  if (any(bad)) {
    v[bad] <- suppressWarnings(as.numeric(
      stringr::str_replace_all(as.character(x)[bad], ",", ".")
    ))
  }
  v
}

# ----------------------------------------------------------------------------
# 2. Проверка и подготовка месячного ряда
# ----------------------------------------------------------------------------

prepare_monthly_df <- function(df, date_col = "date", value_col = "value") {

  stopifnot(date_col %in% names(df))
  stopifnot(value_col %in% names(df))

  out <- tibble(
    date  = parse_to_month(df[[date_col]]),
    value = parse_numeric(df[[value_col]])
  ) %>%
    arrange(date)

  if (any(is.na(out$date)))  stop("Не удалось разобрать часть дат (получились NA). Проверьте колонку дат.")
  if (any(is.na(out$value))) stop("В значениях ряда есть NA (или их не удалось привести к числу).")
  if (any(duplicated(out$date))) stop("В ряду есть повторяющиеся месяцы.")

  expected <- seq(min(out$date), max(out$date), by = "month")
  if (length(expected) != nrow(out) || any(expected != out$date)) {
    stop("Даты должны образовывать непрерывный месячный ряд без пропусков.")
  }

  out
}

to_ts <- function(values, dates) {
  ts(values,
     start = c(lubridate::year(min(dates)), lubridate::month(min(dates))),
     frequency = 12)
}

# ----------------------------------------------------------------------------
# 3. Преобразование производственного календаря из Excel в месячные (центрированные) регрессоры
# ----------------------------------------------------------------------------

make_ru_calendar_from_excel <- function(month_dates,
                                        calendar_file,
                                        sheet = 1,
                                        date_col = "date",
                                        workday_col = "is_workday",
                                        holiday_col = "is_holiday",
                                        easter_col = "easter_effect",
                                        include_easter = TRUE,
                                        center_start = NULL,
                                        center_end = NULL) {

  # center_start / center_end задают ОКНО, по которому считаются средние
  # для центрирования. Если зафиксировать это окно (как в "заморозке"),
  # центрирующие константы перестают плыть от продления календаря и роста
  # выборки. NULL = центрирование по всему календарю.
  month_dates <- parse_to_month(month_dates)

  daily <- readxl::read_xlsx(calendar_file, sheet = sheet)
  names(daily) <- stringr::str_to_lower(stringr::str_trim(names(daily)))

  date_col    <- stringr::str_to_lower(date_col)
  workday_col <- stringr::str_to_lower(workday_col)
  holiday_col <- stringr::str_to_lower(holiday_col)
  easter_col  <- stringr::str_to_lower(easter_col)

  missing_cols <- setdiff(c(date_col, workday_col, holiday_col), names(daily))
  if (length(missing_cols) > 0) {
    stop("В календаре не найдены колонки: ", paste(missing_cols, collapse = ", "))
  }
  if (include_easter && !(easter_col %in% names(daily))) {
    stop("include_easter = TRUE, но в календаре нет колонки '", easter_col,
         "'. Добавьте её или поставьте include_easter = FALSE.")
  }
  if (!(easter_col %in% names(daily))) daily[[easter_col]] <- 0L

  daily <- daily %>%
    transmute(
      date          = as.Date(.data[[date_col]]),
      is_workday    = as.integer(.data[[workday_col]]),
      is_holiday    = as.integer(.data[[holiday_col]]),
      easter_effect = as.integer(.data[[easter_col]])
    ) %>%
    arrange(date)

  if (any(is.na(daily$date)))                     stop("В календаре есть даты, которые R не смог прочитать.")
  if (!all(daily$is_workday    %in% c(0, 1)))     stop("is_workday должен содержать только 0 и 1.")
  if (!all(daily$is_holiday    %in% c(0, 1)))     stop("is_holiday должен содержать только 0 и 1.")
  if (!all(daily$easter_effect %in% c(0, 1)))     stop("easter_effect должен содержать только 0 и 1.")

  # Месячная агрегация ПО ВСЕМУ календарю (нужно для устойчивого центрирования)
  monthly_all <- daily %>%
    mutate(month_date = as.Date(format(date, "%Y-%m-01"))) %>%
    group_by(month_date) %>%
    summarise(
      days_in_month = dplyr::n(),
      workdays      = sum(is_workday == 1),
      holidays      = sum(is_holiday == 1),
      weekends      = sum(is_workday == 0 & is_holiday == 0),
      easter_days   = sum(easter_effect == 1),
      .groups = "drop"
    ) %>%
    mutate(moy = lubridate::month(month_date))

  # Средние по месяцу года (январь, февраль, ...) для центрирования —
  # считаются по окну [center_start, center_end], если оно задано.
  center_set <- monthly_all
  if (!is.null(center_start)) center_set <- center_set %>% filter(month_date >= parse_to_month(center_start))
  if (!is.null(center_end))   center_set <- center_set %>% filter(month_date <= parse_to_month(center_end))

  moy_means <- center_set %>%
    group_by(moy) %>%
    summarise(
      m_workdays    = mean(workdays),
      m_holidays    = mean(holidays),
      m_easter_days = mean(easter_days),
      .groups = "drop"
    )

  if (nrow(moy_means) < 12) {
    # окно центрирования не покрывает все 12 месяцев — откатываемся на весь календарь
    warning("Окно центрирования покрывает не все 12 месяцев; центрирую по всему календарю.")
    moy_means <- monthly_all %>%
      group_by(moy) %>%
      summarise(m_workdays = mean(workdays), m_holidays = mean(holidays),
                m_easter_days = mean(easter_days), .groups = "drop")
  }

  out <- tibble(date = month_dates) %>%
    mutate(moy = lubridate::month(date)) %>%
    left_join(monthly_all %>% select(-moy), by = c("date" = "month_date")) %>%
    left_join(moy_means, by = "moy")

  if (any(is.na(out$workdays))) {
    rng <- range(monthly_all$month_date)
    stop("В производственном календаре не хватает месяцев для ряда + горизонта прогноза.\n",
         "Календарь покрывает ", format(rng[1]), " .. ", format(rng[2]),
         ". Продлите календарь или уменьшите forecast_months.")
  }

  out %>%
    transmute(
      date, days_in_month, workdays, holidays, weekends, easter_days,
      # центрированные регрессоры (отклонение месяца от типичного такого месяца)
      workdays_c    = workdays    - m_workdays,
      holidays_c    = holidays    - m_holidays,
      easter_c      = if (include_easter) easter_days - m_easter_days else 0
    )
}

# ----------------------------------------------------------------------------
# 4. Матрица календарных регрессоров (на основе центрированных колонок)
# ----------------------------------------------------------------------------

build_xreg_matrix <- function(cal_df,
                              mode = c("none", "basic", "extended"),
                              include_easter = TRUE) {
  mode <- match.arg(mode)
  if (mode == "none") return(NULL)

  cols <- list(workdays = cal_df$workdays_c)
  if (mode == "extended") cols$holidays <- cal_df$holidays_c
  if (include_easter)     cols$easter   <- cal_df$easter_c

  xreg <- as.matrix(as.data.frame(cols))

  keep <- apply(xreg, 2, function(z) stats::sd(z, na.rm = TRUE) > 1e-8)
  xreg <- xreg[, keep, drop = FALSE]
  if (ncol(xreg) == 0) return(NULL)
  xreg
}

# ----------------------------------------------------------------------------
# 5. Extractor'ы
# ----------------------------------------------------------------------------

safe_final <- function(m) tryCatch(seasonal::final(m), error = function(e) NULL)
safe_udg   <- function(m) tryCatch(seasonal::udg(m),   error = function(e) NULL)

get_udg_value <- function(u, names_try) {
  if (is.null(u) || is.null(names(u))) return(NA_real_)
  hit <- intersect(names_try, names(u))
  if (length(hit) == 0) return(NA_real_)
  suppressWarnings(as.numeric(u[[hit[1]]]))
}

# Преобразование, фактически выбранное X-13 ("log" / "none" / ...)
get_transform <- function(m) {
  tf <- tryCatch(seasonal::transformfunction(m), error = function(e) NA_character_)
  if (!is.na(tf)) return(tf)
  u <- safe_udg(m)
  if (!is.null(u) && "transform" %in% names(u)) return(tolower(as.character(u[["transform"]])))
  NA_character_
}

# Строка ARIMA-модели, например "(0 1 1)(0 1 1)"
get_arima <- function(m) {
  a <- tryCatch(m$model$arima$model, error = function(e) NULL)
  if (!is.null(a) && length(a) == 1 && !is.na(a)) return(as.character(a))
  NA_character_
}

# Имена аутлаеров из оценённой модели, нормализованные к виду "AO2020.3"
extract_outliers <- function(m) {
  cf <- tryCatch(stats::coef(m), error = function(e) NULL)
  if (is.null(cf) || is.null(names(cf))) return(character(0))
  raw <- names(cf)[grepl("^(AO|LS|TC|SO|RP)\\d{4}\\.", names(cf), ignore.case = TRUE)]
  if (length(raw) == 0) return(character(0))

  norm <- vapply(raw, function(s) {
    mm <- stringr::str_match(s, "^([A-Za-z]{2})(\\d{4})\\.(.+)$")
    if (is.na(mm[1, 1])) return(s)
    type <- toupper(mm[1, 2]); yr <- mm[1, 3]; mo_raw <- mm[1, 4]
    moi <- suppressWarnings(as.integer(mo_raw))
    if (is.na(moi)) moi <- match(tolower(substr(mo_raw, 1, 3)), tolower(month.abb))
    if (is.na(moi)) return(s)
    sprintf("%s%s.%d", type, yr, moi)
  }, character(1))
  unname(norm)
}

# p-value QS-теста на сезонность для конкретного ряда (qsori, qsorievadj, qssadj, ...)
# Колонка p-value в qs() называется "p-val", поэтому ищем её через grep.
get_qs_pval <- function(m, which) {
  q <- tryCatch(seasonal::qs(m), error = function(e) NULL)
  if (is.null(q) || !which %in% rownames(q)) return(NA_real_)
  pcol <- grep("^p", colnames(q), ignore.case = TRUE)
  if (length(pcol) == 0) pcol <- ncol(q)
  suppressWarnings(as.numeric(q[which, pcol[1]]))
}

# Сколько календарных регрессоров значимо
calendar_significance <- function(m, alpha = 0.05) {
  s <- tryCatch(summary(m)$coefficients, error = function(e) NULL)
  if (is.null(s)) return(c(n_sig = NA_integer_, n_total = NA_integer_))
  rows <- grep("^xreg[0-9]+$", rownames(s))
  if (length(rows) == 0) return(c(n_sig = 0L, n_total = 0L))
  pcol <- grep("Pr\\(|p.val|p\\.value", colnames(s), ignore.case = TRUE)
  if (length(pcol) == 0) pcol <- ncol(s)
  pv <- suppressWarnings(as.numeric(s[rows, pcol[1]]))
  c(n_sig = sum(pv < alpha, na.rm = TRUE), n_total = length(rows))
}

# Тренд из модели: s12 для SEATS, d12 для X11
get_trend <- function(m, n) {
  for (code in c("s12", "d12")) {
    tr <- tryCatch(suppressMessages(as.numeric(seasonal::series(m, code))),
                   error = function(e) NULL)
    if (!is.null(tr) && length(tr) >= n) return(tr[seq_len(n)])
  }
  rep(NA_real_, n)
}

# ----------------------------------------------------------------------------
# 6. Одна оценка seas() (с переходом, если упадет, SEATS в X11)
# ----------------------------------------------------------------------------

fit_one <- function(y_ts,
                    xreg = NULL,
                    transform_function = "none",
                    auto_outlier = TRUE,
                    outlier_types = "all",
                    outlier_critical = NULL,
                    outlier_span = NULL,
                    fixed_outliers = NULL,
                    arima_model = NULL,
                    forecast_maxlead = 36,
                    decomposition = c("seats", "x11")) {

  decomposition <- match.arg(decomposition)

  args <- list(
    x                  = y_ts,
    regression.aictest = NULL,                 # свои календарные регрессоры, авто-тест X-13 не нужен
    transform.function = transform_function,
    forecast.maxlead   = forecast_maxlead
  )

  if (!is.null(xreg)) {
    args$xreg <- ts(xreg, start = stats::start(y_ts), frequency = stats::frequency(y_ts))
    args$regression.usertype <- "td"           # всё это календарные эффекты -> убираются из SA-ряда
  }

  if (!is.null(fixed_outliers) && length(fixed_outliers) > 0) {
    args$regression.variables <- fixed_outliers
  }

  if (isTRUE(auto_outlier)) {
    args$outlier       <- ""
    args$outlier.types <- outlier_types
    if (!is.null(outlier_critical)) args$outlier.critical <- outlier_critical
    if (!is.null(outlier_span))     args$outlier.span     <- outlier_span  # искать выбросы только в этом окне
  } else {
    # ВАЖНО: args$outlier <- NULL УДАЛИЛ бы элемент -> seas() включил бы детекцию
    # по умолчанию. Чтобы реально отключить авто-поиск, нужен именованный NULL:
    args["outlier"] <- list(NULL)
  }

  if (!is.null(arima_model)) {
    args$arima.model    <- arima_model
    args["automdl"]     <- list(NULL)           # именованный NULL -> отключить automdl
  }

  if (decomposition == "seats") {
    args$seats          <- ""
    args$seats.noadmiss <- "yes"
  } else {
    args$x11 <- ""
  }

  do.call(seasonal::seas, args)
}

fit_with_fallback <- function(..., method = c("prefer_seats", "seats", "x11")) {
  method <- match.arg(method)
  dots <- list(...)
  run <- function(decomp) {
    a <- dots; a$decomposition <- decomp
    tryCatch(do.call(fit_one, a), error = function(e) e)
  }

  try_order <- switch(method,
    prefer_seats = c("seats", "x11"),   # SEATS, при неудаче -> X11
    seats        = "seats",             # только SEATS (если не строится — честно падаем)
    x11          = "x11"                # только X11
  )

  last_err <- NULL
  for (decomp in try_order) {
    m <- run(decomp)
    if (!inherits(m, "error") && !is.null(safe_final(m))) {
      return(list(ok = TRUE, model = m, method = decomp, error = NA_character_))
    }
    last_err <- if (inherits(m, "error")) conditionMessage(m) else
      paste0("final() не отработал для ", toupper(decomp))
  }

  if (is.null(last_err)) last_err <- "не удалось оценить модель"
  list(ok = FALSE, model = NULL, method = NA_character_, error = last_err)
}

# Сводка по одной оценённой модели (для таблицы сравнения)
summarise_fit <- function(fit, model_name) {
  m <- fit$model
  u <- safe_udg(m)
  cs <- calendar_significance(m)
  tibble(
    model            = model_name,
    method           = fit$method,
    transform        = get_transform(m),
    arima            = get_arima(m),
    n_outliers       = length(extract_outliers(m)),
    aicc             = get_udg_value(u, c("aicc", "AICC")),
    aic              = get_udg_value(u, c("aic", "AIC")),
    bic              = get_udg_value(u, c("bic", "BIC")),
    cal_signif       = unname(cs["n_sig"]),
    cal_total        = unname(cs["n_total"]),
    qs_resid_pval    = get_qs_pval(m, "qssadj")   # сезонность, ОСТАВШАЯСЯ в SA-ряду
  )
}

# ----------------------------------------------------------------------------
# 7. Главная функция
# ----------------------------------------------------------------------------

sa_ru <- function(df,
                  calendar_file,
                  calendar_sheet = 1,
                  date_col = "date",
                  value_col = "value",
                  calendar_date_col = "date",
                  calendar_workday_col = "is_workday",
                  calendar_holiday_col = "is_holiday",
                  calendar_easter_col = "easter_effect",
                  calendar_mode = c("auto", "none", "basic", "extended"),
                  include_easter = TRUE,
                  use_outliers = TRUE,
                  outlier_types = "all",
                  outlier_critical = NULL,
                  transform_function = c("none", "auto", "log"),
                  method = c("prefer_seats", "seats", "x11"),
                  forecast_months = 36,
                  seasonality_alpha = 0.05,
                  center_start = NULL,
                  center_end = NULL,
                  verbose = TRUE) {

  calendar_mode      <- match.arg(calendar_mode)
  transform_function <- match.arg(transform_function)
  method             <- match.arg(method)

  say <- function(...) if (isTRUE(verbose)) message(...)

  # --- 1. данные
  dat  <- prepare_monthly_df(df, date_col, value_col)
  y_ts <- to_ts(dat$value, dat$date)
  n    <- nrow(dat)

  # --- 2. защита логарифма
  if (transform_function == "log" && any(dat$value <= 0)) {
    stop("transform_function = 'log', но в ряду есть значения <= 0. Используйте 'none'.")
  }
  if (transform_function == "auto" && any(dat$value <= 0)) {
    say("В ряду есть значения <= 0 -> для auto принудительно ставим transform = 'none'.")
    transform_function <- "none"
  }

  # --- 3. календарь на горизонт ряд + прогноз
  xreg_dates <- seq(min(dat$date),
                    max(dat$date) %m+% months(forecast_months),
                    by = "month")

  # центрирование по умолчанию — по периоду выборки ряда (а не по всему
  # календарю), чтобы продление календаря в будущее не двигало историю.
  if (is.null(center_start)) center_start <- min(dat$date)
  if (is.null(center_end))   center_end   <- max(dat$date)

  cal_df <- make_ru_calendar_from_excel(
    month_dates  = xreg_dates,
    calendar_file = calendar_file,
    sheet        = calendar_sheet,
    date_col     = calendar_date_col,
    workday_col  = calendar_workday_col,
    holiday_col  = calendar_holiday_col,
    easter_col   = calendar_easter_col,
    include_easter = include_easter,
    center_start = center_start,
    center_end   = center_end
  )

  xreg_for <- function(mode) build_xreg_matrix(cal_df, mode, include_easter)
  cal_names_of <- function(x) if (is.null(x)) character(0) else colnames(x)

  comparison <- NULL
  attempts   <- list()

  # ==========================================================================
  # РЕЖИМ AUTO: фиксируем transform + аутлаеры на basic, сравниваем none/basic/extended
  # ==========================================================================
  if (calendar_mode == "auto") {

    say("AUTO: опорная модель на basic-календаре (фиксируем преобразование и аутлаеры)...")

    # ВАЖНО: опорную модель всегда считаем через X11. Аутлаеры и преобразование
    # определяются в regARIMA (до разложения) и НЕ зависят от seats/x11, а X11
    # практически всегда строится. Так фиксированный набор аутлаеров одинаков при любом выбранном method 
    ref <- fit_with_fallback(
      y_ts = y_ts, xreg = xreg_for("basic"),
      transform_function = transform_function,
      auto_outlier = use_outliers, outlier_types = outlier_types,
      outlier_critical = outlier_critical,
      fixed_outliers = NULL, arima_model = NULL,
      forecast_maxlead = forecast_months, method = "x11"
    )

    # если basic не оценился — пробуем none, потом extended как опорную
    if (!ref$ok) {
      for (fallback_mode in c("none", "extended")) {
        say("  опорная на basic не вышла, пробую ", fallback_mode, " ...")
        ref <- fit_with_fallback(
          y_ts = y_ts, xreg = xreg_for(fallback_mode),
          transform_function = transform_function,
          auto_outlier = use_outliers, outlier_types = outlier_types,
          outlier_critical = outlier_critical,
          fixed_outliers = NULL, arima_model = NULL,
          forecast_maxlead = forecast_months, method = "x11"
        )
        if (ref$ok) break
      }
    }
    if (!ref$ok) stop("Не удалось оценить опорную модель. Ошибка X-13: ", ref$error)

    tf_fixed       <- get_transform(ref$model)
    if (is.na(tf_fixed)) tf_fixed <- if (transform_function == "auto") "none" else transform_function
    outliers_fixed <- if (use_outliers) extract_outliers(ref$model) else character(0)

    say("  зафиксировано: преобразование = ", tf_fixed,
        " | аутлаеров = ", length(outliers_fixed),
        if (length(outliers_fixed)) paste0(" (", paste(outliers_fixed, collapse = ", "), ")") else "")

    candidates <- list()
    for (nm in c("none", "basic", "extended")) {
      say("  оцениваю кандидата: ", nm)
      xr  <- xreg_for(nm)
      fit <- fit_with_fallback(
        y_ts = y_ts, xreg = xr,
        transform_function = tf_fixed,
        auto_outlier = FALSE,
        fixed_outliers = outliers_fixed,
        arima_model = NULL,                 # ARIMA подбирается заново под каждый календарь
        forecast_maxlead = forecast_months, method = method
      )
      if (fit$ok) {
        candidates[[nm]] <- fit
        attempts[[nm]] <- summarise_fit(fit, nm)
      } else {
        attempts[[nm]] <- tibble(model = nm, method = NA_character_, transform = tf_fixed,
                                 arima = NA_character_, n_outliers = length(outliers_fixed),
                                 aicc = NA_real_, aic = NA_real_, bic = NA_real_,
                                 cal_signif = NA_integer_, cal_total = length(cal_names_of(xr)),
                                 qs_resid_pval = NA_real_)
      }
    }

    comparison <- bind_rows(attempts)
    if (length(candidates) == 0) {
      print(comparison)
      hint <- if (method == "seats")
        " Похоже, SEATS не строит разложение для этого ряда — попробуйте method = 'x11' или 'prefer_seats'." else ""
      stop("Ни один кандидат не оценился (method = '", method, "').", hint)
    }

    valid <- comparison %>% filter(model %in% names(candidates), !is.na(aicc))
    if (nrow(valid) == 0) stop("Ни у одного кандидата нет AICc для сравнения.")
    best_name <- valid$model[which.min(valid$aicc)]
    best_fit  <- candidates[[best_name]]

    comparison <- comparison %>% mutate(chosen = (model == best_name))

    say("Выбранная спецификация (min AICc): ", best_name)
    if (isTRUE(verbose)) print(comparison)

  } else {
    # ========================================================================
    # РУЧНОЙ РЕЖИМ: одна спецификация, авто-аутлаеры, всё детерминировано
    # ========================================================================
    say("Ручной режим, календарь = ", calendar_mode)
    xr <- xreg_for(calendar_mode)

    best_fit <- fit_with_fallback(
      y_ts = y_ts, xreg = xr,
      transform_function = transform_function,
      auto_outlier = use_outliers, outlier_types = outlier_types,
      outlier_critical = outlier_critical,
      fixed_outliers = NULL, arima_model = NULL,
      forecast_maxlead = forecast_months, method = method
    )
    if (!best_fit$ok) stop("Не удалось оценить модель (method = '", method, "'). Ошибка X-13: ", best_fit$error)

    best_name  <- calendar_mode
    comparison <- summarise_fit(best_fit, calendar_mode) %>%
      mutate(chosen = TRUE)
    if (isTRUE(verbose)) print(comparison)
  }

  best_model <- best_fit$model

  # --- 4. ЕСТЬ ЛИ В РЯДУ СЕЗОННОСТЬ ВООБЩЕ
  # QS-тест на исходном ряду. Основной — qsorievadj (с поправкой на выбросы):
  # российские ряды полны шоков (COVID, 2022), и без поправки QS их не "видит".
  # Сезонность считаем найденной, если значим хотя бы один из qsorievadj / qsori.
  qs_evadj_p <- get_qs_pval(best_model, "qsorievadj")
  qs_ori_p   <- get_qs_pval(best_model, "qsori")
  qs_guard   <- suppressWarnings(min(c(qs_evadj_p, qs_ori_p), na.rm = TRUE))
  if (is.infinite(qs_guard)) qs_guard <- NA_real_
  # если тест не посчитался (NA) — консервативно считаем, что сезонность есть
  seasonality_detected <- is.na(qs_guard) || qs_guard < seasonality_alpha

  tf_final <- get_transform(best_model)
  multiplicative <- identical(tf_final, "log")

  if (!seasonality_detected) {
    say("Идентифицируемая сезонность не обнаружена (QS p = ",
        round(qs_guard, 4), ") -> ряд возвращается КАК ЕСТЬ.")
    adjusted <- dat$value
  } else {
    adjusted <- as.numeric(safe_final(best_model))
    if (length(adjusted) != n) adjusted <- adjusted[seq_len(n)]
  }

  # фактор сезонности в правильной метрике
  if (multiplicative) {
    factor_type <- "multiplicative"
    seas_factor <- ifelse(!is.na(adjusted) & adjusted != 0, dat$value / adjusted, NA_real_)
  } else {
    factor_type <- "additive"
    seas_factor <- dat$value - adjusted
  }

  # тренд (по возможности): s12 (SEATS) или d12 (X11)
  trend <- if (seasonality_detected) get_trend(best_model, n) else rep(NA_real_, n)

  result_df <- tibble(
    date            = dat$date,
    original        = dat$value,
    adjusted        = adjusted,
    trend           = trend,
    seasonal_factor = seas_factor,
    factor_type     = factor_type
  )

  # статическая (воспроизводимая) спецификация — пригодится для будущей "заморозки".
  # static() печатает спеку в поток сообщений (stderr) -> глушим type="message".
  # coef = TRUE фиксирует и сами коэффициенты; evaluate = FALSE -> просто текст спеки.
  spec_static <- tryCatch({
    st <- NULL
    invisible(utils::capture.output(
      st <- seasonal::static(best_model, coef = TRUE, evaluate = FALSE),
      type = "message"
    ))
    paste(deparse(st), collapse = "\n")
  }, error = function(e) NA_character_)

  list(
    data                 = result_df,
    chosen_model         = best_name,
    transform            = tf_final,
    decomposition_method = best_fit$method,
    arima                = get_arima(best_model),
    outliers             = extract_outliers(best_model),
    seasonality_detected = seasonality_detected,
    qs_orig_pval         = qs_ori_p,
    qs_orig_evadj_pval   = qs_evadj_p,
    comparison           = comparison,
    best_model_object    = best_model,
    udg                  = safe_udg(best_model),
    calendar_monthly     = cal_df,
    spec_static          = spec_static
  )
}

# ----------------------------------------------------------------------------
# 8. ЗАМОРОЗКА СПЕЦИФИКАЦИИ
#    sa_ru_identify() — раз в год: подбирает и сохраняет спеку в sa_specs/<id>.json
#    sa_ru_apply()    — каждый месяц: применяет замороженную спеку к новым данным
#    sa_ru_list_specs() — показывает все сохранённые спеки
# ----------------------------------------------------------------------------

sa_ru_identify <- function(df,
                           calendar_file,
                           series_id = NULL,                  # имя ряда -> файл sa_specs/<id>.json; NULL = НЕ сохранять, только вернуть спеку
                           cutoff_date = NULL,                # дата, ДО которой фиксируем (вкл.); NULL = последнее наблюдение
                           dir = "sa_specs",
                           calendar_sheet = 1,
                           date_col = "date",
                           value_col = "value",
                           calendar_date_col = "date",
                           calendar_workday_col = "is_workday",
                           calendar_holiday_col = "is_holiday",
                           calendar_easter_col = "easter_effect",
                           calendar_mode = c("auto", "none", "basic", "extended"),
                           include_easter = TRUE,
                           use_outliers = TRUE,
                           outlier_types = "all",
                           outlier_critical = NULL,
                           transform_function = c("none", "auto", "log"),
                           method = c("prefer_seats", "seats", "x11"),
                           forecast_months = 36,
                           seasonality_alpha = 0.05,
                           verbose = TRUE) {

  if (!requireNamespace("jsonlite", quietly = TRUE))
    stop("Нужен пакет jsonlite: install.packages('jsonlite').")

  calendar_mode      <- match.arg(calendar_mode)
  transform_function <- match.arg(transform_function)
  method             <- match.arg(method)

  dat_full <- prepare_monthly_df(df, date_col, value_col)
  cutoff   <- if (is.null(cutoff_date)) max(dat_full$date) else parse_to_month(cutoff_date)
  dat      <- dat_full %>% filter(date <= cutoff)
  if (nrow(dat) < 36) stop("Слишком короткая выборка для идентификации (нужно >= 3 года).")
  sample_start <- min(dat$date)

  # обычный прогон на данных до границы — он выбирает transform / календарь / ARIMA / выбросы
  res <- sa_ru(
    df = dat, calendar_file = calendar_file, calendar_sheet = calendar_sheet,
    date_col = "date", value_col = "value",
    calendar_date_col = calendar_date_col, calendar_workday_col = calendar_workday_col,
    calendar_holiday_col = calendar_holiday_col, calendar_easter_col = calendar_easter_col,
    calendar_mode = calendar_mode, include_easter = include_easter,
    use_outliers = use_outliers, outlier_types = outlier_types, outlier_critical = outlier_critical,
    transform_function = transform_function, method = method,
    forecast_months = forecast_months, seasonality_alpha = seasonality_alpha,
    center_start = sample_start, center_end = cutoff,    # окно центрирования = период идентификации
    verbose = verbose
  )

  af <- cutoff %m+% months(1)
  spec <- list(
    series_id            = if (is.null(series_id)) NA_character_ else series_id,
    identified_on        = as.character(Sys.Date()),
    sample_start         = format(sample_start, "%Y-%m"),
    data_through         = format(cutoff, "%Y-%m"),
    transform            = res$transform,
    calendar_mode        = res$chosen_model,
    include_easter       = isTRUE(include_easter),
    method               = method,
    arima                = res$arima,
    outliers             = as.character(res$outliers),
    outlier_auto_from    = paste0(lubridate::year(af), ".", lubridate::month(af)),  # выбросы после границы ищем сами
    seasonality_detected = isTRUE(res$seasonality_detected),
    forecast_months      = forecast_months,
    center_start         = format(sample_start, "%Y-%m"),
    center_end           = format(cutoff, "%Y-%m"),
    calendar_cols        = list(date = calendar_date_col, workday = calendar_workday_col,
                                holiday = calendar_holiday_col, easter = calendar_easter_col),
    calendar_sheet       = calendar_sheet
  )

  # сохраняем ТОЛЬКО если задан series_id; иначе просто возвращаем спеку
  if (!is.null(series_id)) {
    if (!dir.exists(dir)) dir.create(dir, recursive = TRUE)
    path <- file.path(dir, paste0(series_id, ".json"))
    jsonlite::write_json(spec, path, auto_unbox = TRUE, pretty = TRUE)
    if (isTRUE(verbose)) message("Спецификация сохранена: ", normalizePath(path))
  } else if (isTRUE(verbose)) {
    message("series_id не задан -> спека НЕ сохранена (только возвращена в объекте).")
  }

  if (isTRUE(verbose)) {
    message("  до ", spec$data_through, " | transform=", spec$transform,
            " | календарь=", spec$calendar_mode, " | ARIMA=", spec$arima,
            " | выбросов=", length(spec$outliers), " | сезонность=", spec$seasonality_detected)
  }
  invisible(spec)
}

sa_ru_apply <- function(df,
                        calendar_file,
                        series_id,
                        dir = "sa_specs",
                        date_col = "date",
                        value_col = "value",
                        verbose = TRUE) {

  if (!requireNamespace("jsonlite", quietly = TRUE))
    stop("Нужен пакет jsonlite: install.packages('jsonlite').")

  path <- file.path(dir, paste0(series_id, ".json"))
  if (!file.exists(path))
    stop("Не найдена спецификация: ", path, ". Сначала запустите sa_ru_identify(series_id = '", series_id, "', ...).")
  spec <- jsonlite::fromJSON(path)
  say  <- function(...) if (isTRUE(verbose)) message(...)

  dat      <- prepare_monthly_df(df, date_col, value_col)
  y_ts     <- to_ts(dat$value, dat$date)
  n        <- nrow(dat)
  data_end <- max(dat$date)

  # календарь с ФИКСИРОВАННЫМ окном центрирования из спеки (история не едет)
  cc <- spec$calendar_cols
  xreg_dates <- seq(min(dat$date), data_end %m+% months(spec$forecast_months), by = "month")
  cal_df <- make_ru_calendar_from_excel(
    month_dates = xreg_dates, calendar_file = calendar_file, sheet = spec$calendar_sheet,
    date_col = cc$date, workday_col = cc$workday, holiday_col = cc$holiday, easter_col = cc$easter,
    include_easter = isTRUE(spec$include_easter),
    center_start = spec$center_start, center_end = spec$center_end
  )
  xr <- build_xreg_matrix(cal_df, spec$calendar_mode, isTRUE(spec$include_easter))

  fixed_out <- as.character(spec$outliers)
  fixed_out <- fixed_out[!is.na(fixed_out) & nzchar(fixed_out)]

  # авто-поиск выбросов разрешён только ПОСЛЕ границы заморозки (правило Банка России)
  af_parts <- as.integer(strsplit(spec$outlier_auto_from, "\\.")[[1]])
  af_date  <- as.Date(sprintf("%04d-%02d-01", af_parts[1], af_parts[2]))
  do_auto  <- af_date <= data_end

  fit <- list(method = NA_character_, model = NULL)
  if (!isTRUE(spec$seasonality_detected)) {
    say("В спеке отмечено отсутствие сезонности -> ряд возвращается как есть.")
    adjusted <- dat$value
  } else {
    fit <- fit_with_fallback(
      y_ts = y_ts, xreg = xr,
      transform_function = spec$transform,
      auto_outlier = do_auto,
      outlier_span = if (do_auto) paste0(spec$outlier_auto_from, ",") else NULL,  # "2026.1," = от границы до конца
      fixed_outliers = fixed_out,
      arima_model = spec$arima,
      forecast_maxlead = spec$forecast_months,
      method = spec$method
    )
    if (!isTRUE(fit$ok)) stop("Не удалось применить спеку '", series_id, "'. Ошибка X-13: ", fit$error)
    adjusted <- as.numeric(safe_final(fit$model))
    if (length(adjusted) != n) adjusted <- adjusted[seq_len(n)]
  }

  multiplicative <- identical(spec$transform, "log")
  if (multiplicative) {
    seas_factor <- ifelse(!is.na(adjusted) & adjusted != 0, dat$value / adjusted, NA_real_)
    factor_type <- "multiplicative"
  } else {
    seas_factor <- dat$value - adjusted
    factor_type <- "additive"
  }

  result_df <- tibble(date = dat$date, original = dat$value, adjusted = adjusted,
                      seasonal_factor = seas_factor, factor_type = factor_type)

  border    <- parse_to_month(spec$data_through)
  new_points <- result_df %>% filter(date > border)

  if (isTRUE(verbose)) {
    say("Применена спека '", series_id, "' (идентиф. ", spec$identified_on, ", заморожена до ", spec$data_through, ").")
    say("  метод=", fit$method, " | авто-выбросы в новом периоде: ", do_auto,
        " | новых точек после границы: ", nrow(new_points))
  }

  list(
    data                 = result_df,
    new_points           = new_points,    # точки после границы заморозки — их и дописывают в базу
    spec                 = spec,
    decomposition_method = fit$method,
    model                = fit$model
  )
}

sa_ru_list_specs <- function(dir = "sa_specs") {
  if (!requireNamespace("jsonlite", quietly = TRUE)) stop("Нужен пакет jsonlite.")
  if (!dir.exists(dir)) { message("Папка '", dir, "' не найдена."); return(invisible(tibble())) }
  files <- list.files(dir, pattern = "\\.json$", full.names = TRUE)
  if (length(files) == 0) { message("В '", dir, "' нет сохранённых спек."); return(invisible(tibble())) }
  rows <- lapply(files, function(f) {
    s <- tryCatch(jsonlite::fromJSON(f), error = function(e) NULL)
    if (is.null(s)) return(NULL)
    tibble(
      series_id     = s$series_id,
      identified_on = s$identified_on,
      data_through  = s$data_through,
      transform     = s$transform,
      calendar_mode = s$calendar_mode,
      include_easter = isTRUE(s$include_easter),
      method        = s$method,
      arima         = s$arima,
      n_outliers    = length(as.character(s$outliers)),
      seasonality   = isTRUE(s$seasonality_detected)
    )
  })
  dplyr::bind_rows(rows)
}
