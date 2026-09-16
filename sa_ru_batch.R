# ============================================================================
# ПАКЕТНАЯ ОБРАБОТКА НЕСКОЛЬКИХ РЯДОВ — R-РЕАЛИЗАЦИЯ
# sa_ru_batch.R
#
# Источник: NadezhdaYurchenko/Seasonal-Adjustment-for-Russia
# Дополнение: ChernenkoRuslan
# ============================================================================
#
# Функция sa_ru_batch() позволяет сезонно скорректировать НЕСКОЛЬКО рядов из
# одного файла за один вызов. Каждый ряд может иметь собственные настройки.
#
# Поддерживаемые форматы входных данных:
#   "wide" — дата в одной колонке, каждый ряд — отдельная колонка
#            date        | series_A | series_B | series_C
#            2015-01-01  | 101.2    | 85.4     | 12.3
#
#   "long" — дата + идентификатор ряда + значение (tidy-формат)
#            date        | series_id | value
#            2015-01-01  | series_A  | 101.2
#            2015-01-01  | series_B  | 85.4
#
# Пример использования:
#   source("sa_ru_functions.R")
#   source("sa_ru_batch.R")
#
#   # Настройки по умолчанию + перегрузка для конкретных рядов
#   results <- sa_ru_batch(
#     df            = my_wide_df,
#     calendar_file = "russia_calendar.xlsx",
#     format        = "wide",
#     date_col      = "date",
#     default_config = list(calendar_mode = "basic", include_easter = TRUE,
#                           transform_function = "auto"),
#     series_configs = list(
#       CPI       = list(transform_function = "log"),
#       PPI       = list(calendar_mode = "extended", include_easter = FALSE),
#       IP_series = list(calendar_mode = "auto")
#     ),
#     verbose = TRUE
#   )
#
#   # Результат — именованный список с полным выводом sa_ru() для каждого ряда
#   results$CPI$data          # скорректированный ряд
#   results$CPI$aic           # AIC
#   results$summary           # сводная таблица по всем рядам
# ============================================================================

library(dplyr)
library(tidyr)
library(tibble)

# ----------------------------------------------------------------------------
# Вспомогательная: объединить default_config с индивидуальным series_config
# ----------------------------------------------------------------------------

.merge_config <- function(default_cfg, series_cfg) {
  # series_cfg перегружает default_cfg; прочие ключи берём из default_cfg
  if (is.null(series_cfg)) return(default_cfg)
  cfg <- default_cfg
  for (nm in names(series_cfg)) cfg[[nm]] <- series_cfg[[nm]]
  cfg
}

# ----------------------------------------------------------------------------
# Вспомогательная: привести df к long-формату
# ----------------------------------------------------------------------------

.to_long <- function(df, format, date_col, series_col, value_col) {
  if (format == "long") {
    stopifnot(all(c(date_col, series_col, value_col) %in% names(df)))
    return(df %>% rename(date = !!date_col, series_id = !!series_col, value = !!value_col))
  }
  # wide -> long
  stopifnot(date_col %in% names(df))
  value_cols <- setdiff(names(df), date_col)
  if (length(value_cols) == 0) stop("В df нет колонок со значениями (только дата).")
  df %>%
    rename(date = !!date_col) %>%
    tidyr::pivot_longer(cols = all_of(value_cols),
                        names_to  = "series_id",
                        values_to = "value")
}

# ----------------------------------------------------------------------------
# Основная функция
# ----------------------------------------------------------------------------

#' Пакетная сезонная корректировка нескольких рядов
#'
#' @param df data.frame — данные в wide или long формате.
#' @param calendar_file Путь к russia_calendar.xlsx.
#' @param format "wide" или "long". По умолчанию "wide".
#' @param date_col Название колонки с датами.
#' @param series_col Для format="long": колонка с именами рядов.
#' @param value_col Для format="long": колонка со значениями.
#' @param default_config Именованный список параметров sa_ru() по умолчанию.
#'   Доступные ключи: calendar_mode, include_easter, transform_function,
#'   use_outliers, outlier_types, outlier_critical, method,
#'   forecast_months, seasonality_alpha.
#' @param series_configs Именованный список списков — перегрузки default_config
#'   для конкретных рядов. Имена должны совпадать с именами рядов в df.
#' @param output_long Если TRUE, результат$combined — long data.frame
#'   со всеми рядами. Иначе только список.
#' @param fail_on_error Если TRUE, ошибка в одном ряду останавливает всё.
#'   Если FALSE — записывает ошибку и продолжает. По умолчанию FALSE.
#' @param verbose Печатать прогресс. По умолчанию TRUE.
#'
#' @return Список:
#'   \item{results}{Именованный список — полный вывод sa_ru() для каждого ряда.}
#'   \item{summary}{Сводная таблица (tibble) со всеми рядами и ключевыми метриками.}
#'   \item{combined}{Если output_long=TRUE: long data.frame со всеми SA-рядами.}
#'   \item{errors}{Именованный вектор ошибок для рядов, которые не оценились.}
sa_ru_batch <- function(df,
                        calendar_file,
                        format          = c("wide", "long"),
                        date_col        = "date",
                        series_col      = "series_id",
                        value_col       = "value",
                        default_config  = list(
                          calendar_mode      = "basic",
                          include_easter     = TRUE,
                          transform_function = "none",
                          use_outliers       = TRUE,
                          outlier_types      = "all",
                          outlier_critical   = NULL,
                          method             = "prefer_seats",
                          forecast_months    = 36,
                          seasonality_alpha  = 0.05
                        ),
                        series_configs  = list(),
                        output_long     = TRUE,
                        fail_on_error   = FALSE,
                        verbose         = TRUE) {

  format <- match.arg(format)
  say    <- function(...) if (isTRUE(verbose)) message(...)

  # --- 1. Приводим к long-формату
  long_df <- .to_long(df, format, date_col, series_col, value_col)
  series_names <- sort(unique(long_df$series_id))
  n_series <- length(series_names)
  say(sprintf("Пакетная обработка: %d рядов", n_series))

  # --- 2. Неизвестные имена в series_configs — предупреждаем
  unknown <- setdiff(names(series_configs), series_names)
  if (length(unknown) > 0) {
    warning("В series_configs есть ряды, которых нет в df: ",
            paste(unknown, collapse = ", "))
  }

  # --- 3. Обработка каждого ряда
  results   <- list()
  errors    <- character(0)
  summary_rows <- list()

  for (i in seq_along(series_names)) {
    sname <- series_names[i]
    say(sprintf("  [%d/%d] %s ...", i, n_series, sname))

    # Данные ряда
    sub_df <- long_df %>%
      filter(series_id == sname) %>%
      select(date, value) %>%
      as.data.frame()

    # Конфигурация: default + перегрузка
    cfg <- .merge_config(default_config, series_configs[[sname]])

    # Запуск sa_ru()
    res <- tryCatch({
      do.call(sa_ru, c(
        list(
          df            = sub_df,
          calendar_file = calendar_file,
          date_col      = "date",
          value_col     = "value",
          verbose       = FALSE
        ),
        cfg[intersect(names(cfg), c(
          "calendar_sheet", "calendar_date_col", "calendar_workday_col",
          "calendar_holiday_col", "calendar_easter_col",
          "calendar_mode", "include_easter", "use_outliers", "outlier_types",
          "outlier_critical", "transform_function", "method",
          "forecast_months", "seasonality_alpha",
          "center_start", "center_end"
        ))]
      ))
    }, error = function(e) e)

    if (inherits(res, "error")) {
      errmsg <- conditionMessage(res)
      say(sprintf("    ОШИБКА: %s", errmsg))
      errors[sname] <- errmsg
      if (isTRUE(fail_on_error)) stop(errmsg)

      summary_rows[[sname]] <- tibble(
        series_id            = sname,
        status               = "ERROR",
        chosen_model         = NA_character_,
        transform            = NA_character_,
        decomposition_method = NA_character_,
        arima                = NA_character_,
        seasonality_detected = NA,
        aic                  = NA_real_,
        n_obs                = nrow(sub_df),
        error_msg            = errmsg
      )
      next
    }

    results[[sname]] <- res

    # Метрики для сводки
    udg <- res$udg
    aic_val <- tryCatch(
      as.numeric(udg[intersect(c("aicc", "AICC", "aic", "AIC"), names(udg))[1]]),
      error = function(e) NA_real_
    )

    say(sprintf("    OK | модель=%s | transform=%s | сезонность=%s",
                res$chosen_model,
                ifelse(is.null(res$transform), "?", res$transform),
                ifelse(isTRUE(res$seasonality_detected), "да", "нет")))

    summary_rows[[sname]] <- tibble(
      series_id            = sname,
      status               = "OK",
      chosen_model         = as.character(res$chosen_model),
      transform            = as.character(res$transform),
      decomposition_method = as.character(res$decomposition_method),
      arima                = as.character(res$arima),
      seasonality_detected = isTRUE(res$seasonality_detected),
      aic                  = aic_val,
      n_obs                = nrow(res$data),
      error_msg            = NA_character_
    )
  }

  summary_tbl <- bind_rows(summary_rows)
  ok_count  <- sum(summary_tbl$status == "OK",    na.rm = TRUE)
  err_count <- sum(summary_tbl$status == "ERROR", na.rm = TRUE)
  say(sprintf("Готово: успешно=%d, ошибок=%d", ok_count, err_count))
  if (isTRUE(verbose)) print(summary_tbl)

  # --- 4. Объединённый long data.frame
  combined <- NULL
  if (isTRUE(output_long) && length(results) > 0) {
    combined <- bind_rows(lapply(names(results), function(nm) {
      results[[nm]]$data %>% mutate(series_id = nm, .before = 1)
    }))
  }

  list(
    results  = results,
    summary  = summary_tbl,
    combined = combined,
    errors   = errors
  )
}

# ----------------------------------------------------------------------------
# Утилита: сохранить результаты batch в Excel (один лист на ряд)
# ----------------------------------------------------------------------------

#' Сохранить результаты sa_ru_batch() в Excel-файл
#'
#' Требует пакет openxlsx или writexl.
#'
#' @param batch_result Вывод sa_ru_batch().
#' @param path Путь к выходному .xlsx файлу.
#' @param include_summary Добавить лист с общей сводкой. По умолчанию TRUE.
sa_ru_batch_to_excel <- function(batch_result,
                                 path = "sa_results.xlsx",
                                 include_summary = TRUE) {
  if (!requireNamespace("openxlsx", quietly = TRUE))
    stop("Нужен пакет openxlsx: install.packages('openxlsx').")

  wb <- openxlsx::createWorkbook()

  if (isTRUE(include_summary)) {
    openxlsx::addWorksheet(wb, "Summary")
    openxlsx::writeData(wb, "Summary", batch_result$summary)
  }

  for (nm in names(batch_result$results)) {
    sheet_nm <- substr(nm, 1, 31)  # Excel limit: 31 chars
    openxlsx::addWorksheet(wb, sheet_nm)
    openxlsx::writeData(wb, sheet_nm, batch_result$results[[nm]]$data)
  }

  openxlsx::saveWorkbook(wb, path, overwrite = TRUE)
  message("Сохранено: ", normalizePath(path))
  invisible(path)
}
