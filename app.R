# ----------------------------
# 0) AUTO-INSTALL PACKAGES
# ----------------------------
pkgs <- c("shiny","readxl","dplyr","stringr","tidyr","DT","openxlsx","plotly","shinyjs")
to_install <- pkgs[!pkgs %in% installed.packages()[,"Package"]]
if (length(to_install) > 0) install.packages(to_install, repos = "https://cloud.r-project.org")

suppressPackageStartupMessages({
  library(shiny)
  library(readxl)
  library(dplyr)
  library(stringr)
  library(tidyr)
  library(DT)
  library(openxlsx)
  library(plotly)
  library(shinyjs)
})

# ----------------------------
# CONFIG (DEFAULT mapping path)
# ----------------------------
MAP_PATH <- "C:/Users/poorn/OneDrive/Desktop/Hayleys/Mapping.xlsx"

# ----------------------------
# GLOBAL THEME COLORS
# ----------------------------
COL_FORECAST <- "#0A1F44"  # Forecast (dark navy)
COL_ACTUAL   <- "#4C6FA6"  # Actual (blue)
COL_STATUS_3 <- "#334155"  # slate
COL_STATUS_4 <- "#E11D48"  # red
COL_CAPACITY <- "#16A34A"  # Max Capacity (green)

# Quality status theme colors (for table)
QSTAT_BG <- c(
  "Forecast > Actual" = "#FFF7ED",
  "Forecast < Actual" = "#EFF6FF",
  "Forecast = Actual" = "#ECFDF5"
)
QSTAT_TXT <- c(
  "Forecast > Actual" = "#9A3412",
  "Forecast < Actual" = "#1D4ED8",
  "Forecast = Actual" = "#065F46"
)

# ----------------------------
# Helpers
# ----------------------------
clean_names_simple <- function(nm) {
  nm %>%
    as.character() %>%
    tolower() %>%
    str_replace_all("[^a-z0-9]+", "_") %>%
    str_replace_all("^_|_$", "")
}

make_names_unique <- function(nm) make.unique(nm, sep = "_")

safe_read_excel <- function(path) {
  x <- read_excel(path)
  names(x) <- make_names_unique(clean_names_simple(names(x)))
  x
}

norm_key <- function(x) x %>% as.character() %>% str_trim() %>% str_squish() %>% str_to_upper()

pick_col <- function(nms, candidates) {
  hit <- intersect(nms, candidates)
  if (length(hit) == 0) return(NA_character_)
  hit[1]
}

round_whole <- function(df) {
  df %>% mutate(across(where(is.numeric), ~as.numeric(round(.x, 0))))
}

add_total_row <- function(df, name_col) {
  num_cols <- names(df)[sapply(df, is.numeric)]
  total <- df %>% summarise(across(all_of(num_cols), ~sum(.x, na.rm = TRUE)))
  total[[name_col]] <- "TOTAL"
  total <- total %>% select(all_of(name_col), all_of(num_cols))
  bind_rows(df, total)
}

# Process levels
process_levels <- c(
  "Washing range","Preset","Sueding","Brushing","Singen","Bagging",
  "Stenter finish","Compactor finish","Calendar finish","Relax Dry",
  "Santex relax dry","Stenter Dry","Stenter Brushing","Stenter Final Finish",
  "Compactor final finish","Heat set","Pre bulk","Wider"
)
process_levels_clean <- clean_names_simple(process_levels)

# ----------------------------
# DT table helper (styled)
# ----------------------------
make_dt <- function(df,
                    page_len = 25,
                    first_col_for_total = NULL,
                    type = c("generic","quality","removed")) {
  
  type <- match.arg(type)
  
  dt <- datatable(
    df,
    rownames = FALSE,
    class = "stripe hover cell-border compact",
    options = list(
      pageLength = page_len,
      scrollX = TRUE,
      dom = "lftip",
      lengthMenu = list(c(6, 12, 25, 50, -1), c("6", "12", "25", "50", "All"))
    )
  )
  
  num_cols <- names(df)[sapply(df, is.numeric)]
  if (length(num_cols) > 0) {
    dt <- dt %>%
      formatRound(columns = num_cols, digits = 0) %>%
      formatStyle(columns = num_cols, `text-align` = "right")
  }
  
  dt <- dt %>%
    formatStyle(
      columns = names(df),
      `font-family` = "Arial",
      `font-size` = "13px",
      color = "#111827",
      `border-color` = "#E5E7EB",
      `padding` = "8px 10px"
    )
  
  if (!is.null(first_col_for_total) && first_col_for_total %in% names(df)) {
    dt <- dt %>%
      formatStyle(
        columns = names(df),
        target = "row",
        backgroundColor = styleEqual("TOTAL", "#FFF7ED"),
        color = styleEqual("TOTAL", "#9A3412"),
        fontWeight = styleEqual("TOTAL", "bold"),
        valueColumns = first_col_for_total
      )
  }
  
  if (type == "quality" && "Status" %in% names(df)) {
    dt <- dt %>%
      formatStyle(
        "Status",
        backgroundColor = styleEqual(
          c("Forecast > Actual","Forecast < Actual","Forecast = Actual","NOT MATCHED (Missing in Mapping)"),
          c("#FFF7ED", "#EFF6FF", "#ECFDF5", "#FEF2F2")
        ),
        color = styleEqual(
          c("Forecast > Actual","Forecast < Actual","Forecast = Actual","NOT MATCHED (Missing in Mapping)"),
          c("#9A3412", "#1D4ED8", "#065F46", "#991B1B")
        ),
        fontWeight = "bold"
      )
  }
  
  if (type == "removed" && "Removed Reason" %in% names(df)) {
    dt <- dt %>%
      formatStyle(
        "Removed Reason",
        backgroundColor = styleEqual("Forecast = 0 (Hidden from chart)", "#FEF3C7"),
        color = styleEqual("Forecast = 0 (Hidden from chart)", "#92400E"),
        fontWeight = "bold"
      )
  }
  
  dt
}

plotly_title <- function(title_text) {
  list(
    text = paste0("<b>", title_text, "</b>"),
    x = 0.01, xanchor = "left",
    y = 0.985, yanchor = "top",
    font = list(size = 20, color = COL_FORECAST)
  )
}

# ----------------------------
# Removed table (Forecast==0) for UI
# ----------------------------
get_removed_table <- function(df, name_col, fcol = "Forecast_VAL", acol = "Actual_VAL", top_n = 200) {
  d_all <- df %>%
    filter(.data[[name_col]] != "TOTAL") %>%
    transmute(
      Category = as.character(.data[[name_col]]),
      Forecast = as.numeric(.data[[fcol]]),
      Actual   = as.numeric(.data[[acol]])
    ) %>%
    mutate(
      Gap = abs(Actual - Forecast),
      Gap_pct = ifelse(Forecast == 0, NA_real_, (Gap / Forecast) * 100),
      reason = case_when(
        Forecast <= 0 | is.na(Forecast) ~ "Forecast = 0 (Hidden from chart)",
        TRUE ~ NA_character_
      )
    ) %>%
    filter(!is.na(reason)) %>%
    arrange(desc(Actual + Forecast), desc(Gap))
  
  if (nrow(d_all) > top_n) d_all <- d_all %>% slice(1:top_n)
  
  if (nrow(d_all) == 0) {
    return(data.frame(Note = "No categories removed (all have Forecast > 0)."))
  }
  
  d_all %>%
    mutate(
      Gap_pct = round(Gap_pct, 1),
      Forecast = round(Forecast, 0),
      Actual = round(Actual, 0),
      Gap = round(Gap, 0)
    ) %>%
    select(Category, Forecast, Actual, Gap, Gap_pct, reason) %>%
    rename(
      `Forecast (T)` = Forecast,
      `Actual (T)`   = Actual,
      `Gap (T)`      = Gap,
      `Gap%`         = Gap_pct,
      `Removed Reason` = reason
    )
}

# ----------------------------
# Chart ordering
# ----------------------------
compute_salesgroup_order <- function(df, name_col = "Sales_Group") {
  d <- df %>%
    filter(.data[[name_col]] != "TOTAL") %>%
    mutate(label = as.character(.data[[name_col]])) %>%
    distinct(label)
  
  present <- d$label
  fixed <- c("SG1","SG2","SG3","SG4","SG5","SG6")
  fixed_present <- fixed[fixed %in% present]
  
  others <- setdiff(present, fixed)
  others_sorted <- others[order(others)]
  c(fixed_present, others_sorted)
}

compute_chart_order <- function(df, name_col, fcol = "Forecast_VAL", acol = "Actual_VAL", top_n = 25) {
  d <- df %>%
    filter(.data[[name_col]] != "TOTAL") %>%
    mutate(
      label = as.character(.data[[name_col]]),
      f = as.numeric(.data[[fcol]]),
      a = as.numeric(.data[[acol]]),
      gap_abs = abs(a - f)
    ) %>%
    filter(!is.na(f), f > 0) %>%
    arrange(desc(gap_abs), desc(f + a))
  
  if (nrow(d) == 0) return(character(0))
  if (nrow(d) > top_n) d <- d %>% slice(1:top_n)
  d$label
}

# ----------------------------
# Capacity lookup helper (robust match)
# ----------------------------
lookup_capacity_by_label <- function(labels, max_caps) {
  if (is.null(max_caps) || length(max_caps) == 0) return(rep(NA_real_, length(labels)))
  cap_map <- as.numeric(max_caps)
  names(cap_map) <- norm_key(names(max_caps))
  lab_keys <- norm_key(labels)
  out <- unname(cap_map[lab_keys])
  as.numeric(out)
}

# ============================================================
# ✅ UPDATED PLOTLY FUNCTIONS (CLEAN HOVER + CAP GAP LABEL)
# ============================================================

# ----------------------------
# Plotly: Style 1 Overlay bars (+ optional Max Capacity behind)
#   ✅ Hover uses unit "T" only (not "GROSS (T)" / "NET (T)")
#   ✅ Red circle = Actual > Forecast
#   ✅ If capacity exists: the TOP label above GREEN bar shows (Capacity - Actual)
# ----------------------------
make_plotly_overlay_bars <- function(df, name_col, title_prefix = "", top_n = 25,
                                     fcol = "Forecast_VAL", acol = "Actual_VAL",
                                     yaxis_title = "T", yunit = "T",
                                     w_forecast = 0.68, w_actual = 0.34,
                                     label_order = NULL,
                                     max_caps = NULL) {
  
  d <- df %>%
    filter(.data[[name_col]] != "TOTAL") %>%
    mutate(
      label = as.character(.data[[name_col]]),
      f = as.numeric(.data[[fcol]]),
      a = as.numeric(.data[[acol]]),
      gap_abs = abs(a - f),
      gap_pct = ifelse(f == 0, NA_real_, (gap_abs / f) * 100),
      actual_gt = a > f,
      y_low = pmin(f, a, na.rm = TRUE),
      y_high = pmax(f, a, na.rm = TRUE)
    ) %>%
    filter(!is.na(f), f > 0)
  
  if (!is.null(label_order) && length(label_order) > 0) {
    d <- d %>% mutate(label = factor(label, levels = label_order)) %>% arrange(label)
  } else {
    d <- d %>% arrange(desc(gap_abs), desc(f + a))
  }
  
  if (nrow(d) == 0) {
    return(plotly_empty() %>% layout(title = plotly_title(paste0(title_prefix, " (No data)"))))
  }
  if (nrow(d) > top_n) d <- d %>% slice(1:top_n)
  
  # capacity mapping (optional)
  has_cap <- !is.null(max_caps) && length(max_caps) > 0
  if (has_cap) {
    d$cap <- lookup_capacity_by_label(d$label, max_caps)
    d$cap_gap <- d$cap - d$a
  } else {
    d$cap <- NA_real_
    d$cap_gap <- NA_real_
  }
  
  ymax <- max(c(d$f, d$a, d$cap), na.rm = TRUE) * 1.30
  
  # ---- Annotations ----
  # 1) Default: Gap% on top of Forecast/Actual (your earlier behavior)
  ann_gap_pct <- lapply(seq_len(nrow(d)), function(i) {
    list(
      x = as.character(d$label[i]),
      y = max(d$f[i], d$a[i], na.rm = TRUE),
      text = paste0(round(d$gap_pct[i], 1), "%"),
      xanchor = "center",
      yanchor = "bottom",
      showarrow = FALSE,
      font = list(size = 18, color = COL_FORECAST)
    )
  })
  
  # 2) NEW: If capacity exists, show (Capacity - Actual) above GREEN bar
  ann_cap_gap <- list()
  if (has_cap) {
    dcap <- d %>% filter(!is.na(cap))
    if (nrow(dcap) > 0) {
      ann_cap_gap <- lapply(seq_len(nrow(dcap)), function(i) {
        list(
          x = as.character(dcap$label[i]),
          y = dcap$cap[i],
          text = paste0("Cap gap: ", round(dcap$cap_gap[i], 0), " ", yunit),
          xanchor = "center",
          yanchor = "bottom",
          showarrow = FALSE,
          font = list(size = 15, color = COL_CAPACITY)
        )
      })
    }
  }
  
  p <- plot_ly()
  
  # ---- Max Capacity (behind) ----
  if (has_cap) {
    dcap <- d %>% filter(!is.na(cap))
    if (nrow(dcap) > 0) {
      p <- p %>%
        add_bars(
          data = dcap,
          x = ~label, y = ~cap,
          width = 0.92,
          name = "Max Capacity",
          marker = list(
            color = "rgba(22,163,74,0.28)",
            line  = list(color = COL_CAPACITY, width = 2.2)
          ),
          opacity = 0.30,
          showlegend = TRUE,
          hovertemplate = paste(
            "<b>%{x}</b><br>",
            "Max Capacity: %{y:.0f} ", yunit,
            "<extra></extra>"
          )
        )
    }
  }
  
  # ---- Forecast ----
  p <- p %>%
    add_bars(
      data = d, x = ~label, y = ~f,
      width = w_forecast,
      name = "Forecast",
      marker = list(color = COL_FORECAST, line = list(color = "black", width = 2.6)),
      opacity = 0.94,
      showlegend = TRUE,
      hovertemplate = paste(
        "<b>%{x}</b><br>",
        "Forecast: %{y:.0f} ", yunit,
        "<extra></extra>"
      )
    ) %>%
    # ---- Actual ----
  add_bars(
    data = d, x = ~label, y = ~a,
    width = w_actual,
    name = "Actual",
    marker = list(color = COL_ACTUAL, line = list(color = "black", width = 2.6)),
    opacity = 0.98,
    showlegend = TRUE,
    hovertemplate = paste(
      "<b>%{x}</b><br>",
      "Actual: %{y:.0f} ", yunit,
      "<extra></extra>"
    )
  ) %>%
    # connector line (no hover)
    add_segments(
      data = d,
      x = ~label, xend = ~label,
      y = ~y_low, yend = ~y_high,
      line = list(color = "black", width = 3.5),
      showlegend = FALSE,
      hoverinfo = "skip"
    )
  
  # red circle where Actual > Forecast (no hover)
  idx <- which(d$actual_gt)
  if (length(idx) > 0) {
    p <- p %>%
      add_markers(
        data = d[idx, , drop = FALSE],
        x = ~label, y = ~a,
        marker = list(symbol = "circle-open", size = 36, color = "red", line = list(width = 6)),
        inherit = FALSE, hoverinfo = "skip", showlegend = FALSE
      )
  }
  
  # combine annotations
  all_ann <- c(ann_gap_pct, ann_cap_gap)
  
  p %>%
    layout(
      title = plotly_title(title_prefix),
      barmode = "overlay",
      paper_bgcolor = "white",
      plot_bgcolor  = "white",
      xaxis = list(title = "Category", tickfont = list(size = 16, color = COL_FORECAST), tickangle = 35),
      yaxis = list(title = yaxis_title, titlefont = list(size = 20), tickfont = list(size = 16),
                   range = c(0, ymax), rangemode = "tozero"),
      margin = list(b = 220, t = 70),
      annotations = all_ann,
      bargap = 0.26,
      legend = list(orientation = "h", x = 0.01, y = -0.25)
    ) %>%
    config(displayModeBar = TRUE, scrollZoom = TRUE)
}

# ----------------------------
# Plotly: Style 2 Bracket bars
#   ✅ Hover uses unit "T" only
#   ✅ Top label shows ONLY Gap + % (no "GROSS (T)" / "NET (T)")
# ----------------------------
make_plotly_twobars_bracket <- function(df, name_col, title_prefix = "", top_n = 25,
                                        fcol = "Forecast_VAL", acol = "Actual_VAL",
                                        yaxis_title = "T", yunit = "T",
                                        label_order = NULL) {
  
  d <- df %>%
    filter(.data[[name_col]] != "TOTAL") %>%
    mutate(
      label = as.character(.data[[name_col]]),
      f = as.numeric(.data[[fcol]]),
      a = as.numeric(.data[[acol]]),
      gap_abs = abs(a - f),
      gap_pct = ifelse(f == 0, NA_real_, (gap_abs / f) * 100)
    ) %>%
    filter(!is.na(f), f > 0)
  
  if (!is.null(label_order) && length(label_order) > 0) {
    d <- d %>% mutate(label = factor(label, levels = label_order)) %>% arrange(label)
  } else {
    d <- d %>% arrange(desc(gap_abs), desc(f + a))
  }
  
  if (nrow(d) == 0) {
    return(plotly_empty() %>% layout(title = plotly_title(paste0(title_prefix, " (No data)"))))
  }
  if (nrow(d) > top_n) d <- d %>% slice(1:top_n)
  
  d$idx <- seq_len(nrow(d))
  x_f <- d$idx - 0.22
  x_a <- d$idx + 0.22
  top_y <- pmax(d$f, d$a, na.rm = TRUE)
  y_br  <- top_y * 1.06 + 1
  
  p <- plot_ly() %>%
    add_bars(
      data = d,
      x = x_f, y = ~f,
      width = 0.42,
      name = "Forecast",
      marker = list(color = COL_FORECAST, line = list(color = "black", width = 2.6)),
      opacity = 0.96,
      showlegend = TRUE,
      customdata = ~label,
      hovertemplate = paste(
        "<b>%{customdata}</b><br>",
        "Forecast: %{y:.0f} ", yunit,
        "<extra></extra>"
      )
    ) %>%
    add_bars(
      data = d,
      x = x_a, y = ~a,
      width = 0.30,
      name = "Actual",
      marker = list(color = COL_ACTUAL, line = list(color = "black", width = 2.6)),
      opacity = 0.96,
      showlegend = TRUE,
      customdata = ~label,
      hovertemplate = paste(
        "<b>%{customdata}</b><br>",
        "Actual: %{y:.0f} ", yunit,
        "<extra></extra>"
      )
    )
  
  shapes <- list()
  annotations <- list()
  
  for (i in seq_len(nrow(d))) {
    shapes[[length(shapes)+1]] <- list(type = "line", x0 = x_f[i], x1 = x_f[i], y0 = d$f[i], y1 = y_br[i],
                                       line = list(width = 2.5, color = "black"))
    shapes[[length(shapes)+1]] <- list(type = "line", x0 = x_a[i], x1 = x_a[i], y0 = d$a[i], y1 = y_br[i],
                                       line = list(width = 2.5, color = "black"))
    shapes[[length(shapes)+1]] <- list(type = "line", x0 = x_f[i], x1 = x_a[i], y0 = y_br[i], y1 = y_br[i],
                                       line = list(width = 3, color = "black"))
    
    # ✅ Only Gap + % (no "GROSS (T)" / "NET (T)")
    annotations[[length(annotations)+1]] <- list(
      x = d$idx[i], y = y_br[i],
      text = paste0("Gap: ", round(d$gap_abs[i], 0), " ", yunit, " (", round(d$gap_pct[i], 1), "%)"),
      showarrow = FALSE, yanchor = "bottom",
      font = list(size = 17, color = COL_FORECAST)
    )
  }
  
  p %>%
    layout(
      title = plotly_title(title_prefix),
      barmode = "group",
      bargap = 0.32,
      paper_bgcolor = "white",
      plot_bgcolor  = "white",
      xaxis = list(
        title = "Category",
        tickmode = "array",
        tickvals = d$idx,
        ticktext = as.character(d$label),
        tickangle = 35,
        tickfont = list(size = 16, color = COL_FORECAST)
      ),
      yaxis = list(title = yaxis_title, titlefont = list(size = 20), tickfont = list(size = 16),
                   rangemode = "tozero"),
      margin = list(b = 220, t = 70),
      shapes = shapes,
      annotations = annotations,
      legend = list(orientation = "h", x = 0.01, y = -0.25)
    ) %>%
    config(displayModeBar = TRUE, scrollZoom = TRUE)
}

# ============================================================
# Build summaries (Actual vs Forecast) with NET/GROSS basis
# ============================================================
build_all_summaries <- function(forecast_path, actual_path, days_in_month,
                                basis = c("GROSS","NET"),
                                gross_factor = (1/0.9) * 1.02,
                                map_path = MAP_PATH) {
  
  basis <- match.arg(basis)
  if (!file.exists(map_path)) stop(paste0("Mapping file not found at: ", map_path))
  
  fore_raw <- safe_read_excel(forecast_path)
  act_raw  <- safe_read_excel(actual_path)
  map_raw  <- safe_read_excel(map_path)
  
  f_sales <- pick_col(names(fore_raw), c("sales_group"))
  f_qual  <- pick_col(names(fore_raw), c("quality"))
  f_val   <- pick_col(names(fore_raw), c("forecast"))
  
  a_sales <- pick_col(names(act_raw), c("sales_group"))
  a_qual  <- pick_col(names(act_raw), c("quality"))
  a_val   <- pick_col(names(act_raw), c("actual"))
  
  if (is.na(f_sales) || is.na(f_qual) || is.na(f_val)) {
    stop("Forecast Excel must contain columns exactly named: Sales Group, Quality, Forecast")
  }
  if (is.na(a_sales) || is.na(a_qual) || is.na(a_val)) {
    stop("Actual Excel must contain columns exactly named: Sales Group, Quality, Actual")
  }
  
  m_sales <- pick_col(names(map_raw), c("sales_group"))
  m_qual  <- pick_col(names(map_raw), c("quality"))
  if (is.na(m_sales) || is.na(m_qual)) stop("Mapping file must contain: Sales Group and Quality")
  
  map_material_col  <- pick_col(names(map_raw), c("material_composition","material","composition","material_comp"))
  map_finish_col    <- pick_col(names(map_raw), c("finishing_machine","finishing_machines","finishing","finish_machine"))
  map_dye_col       <- pick_col(names(map_raw), c("dye_machine","dye_machines","dye","dyeing_machine"))
  map_processes_col <- pick_col(names(map_raw), c("processes","process","process_name","process_names"))
  map_preset_col    <- pick_col(names(map_raw), c("preset_machine","preset_machines","preset","preset_machine_name","preset_machine_type"))
  
  fore_keyed <- fore_raw %>%
    mutate(
      sales_group_key = norm_key(.data[[f_sales]]),
      quality_key     = norm_key(.data[[f_qual]]),
      Forecast_NET    = as.numeric(.data[[f_val]])
    ) %>%
    group_by(sales_group_key, quality_key) %>%
    summarise(Forecast_NET = sum(Forecast_NET, na.rm = TRUE), .groups = "drop")
  
  act_keyed <- act_raw %>%
    mutate(
      sales_group_key = norm_key(.data[[a_sales]]),
      quality_key     = norm_key(.data[[a_qual]]),
      Actual_NET      = as.numeric(.data[[a_val]])
    ) %>%
    group_by(sales_group_key, quality_key) %>%
    summarise(Actual_NET = sum(Actual_NET, na.rm = TRUE), .groups = "drop")
  
  map_keys <- map_raw %>%
    transmute(
      sales_group_key = norm_key(.data[[m_sales]]),
      quality_key     = norm_key(.data[[m_qual]])
    ) %>%
    distinct()
  
  map1 <- map_raw %>%
    mutate(
      sales_group_key = norm_key(.data[[m_sales]]),
      quality_key     = norm_key(.data[[m_qual]])
    ) %>%
    group_by(sales_group_key, quality_key) %>%
    slice(1) %>% ungroup()
  
  base <- full_join(fore_keyed, act_keyed, by = c("sales_group_key","quality_key")) %>%
    mutate(
      Forecast_NET = replace_na(Forecast_NET, 0),
      Actual_NET   = replace_na(Actual_NET, 0)
    )
  
  base2 <- base %>%
    left_join(map_keys %>% mutate(in_mapping = TRUE),
              by = c("sales_group_key","quality_key")) %>%
    mutate(in_mapping = replace_na(in_mapping, FALSE))
  
  factor_use <- if (basis == "GROSS") gross_factor else 1
  
  base2 <- base2 %>%
    mutate(
      Forecast_VAL = Forecast_NET * factor_use,
      Actual_VAL   = Actual_NET   * factor_use
    )
  
  unmatched_list <- base2 %>%
    filter(!in_mapping) %>%
    transmute(
      Sales_Group = sales_group_key,
      Quality     = quality_key,
      Forecast    = Forecast_VAL,
      Actual      = Actual_VAL
    ) %>%
    arrange(desc(Actual + Forecast)) %>%
    round_whole()
  
  quality_compare <- base2 %>%
    transmute(
      Sales_Group = sales_group_key,
      Quality     = quality_key,
      Forecast    = Forecast_VAL,
      Actual      = Actual_VAL,
      Status = case_when(
        in_mapping == FALSE ~ "NOT MATCHED (Missing in Mapping)",
        Forecast_VAL > Actual_VAL ~ "Forecast > Actual",
        Forecast_VAL < Actual_VAL ~ "Forecast < Actual",
        TRUE ~ "Forecast = Actual"
      )
    ) %>%
    round_whole()
  
  status_counts <- quality_compare %>%
    count(Status) %>%
    mutate(Status = factor(
      Status,
      levels = c("Forecast > Actual","Forecast < Actual","Forecast = Actual","NOT MATCHED (Missing in Mapping)")
    ))
  
  joined <- base2 %>%
    filter(in_mapping) %>%
    select(-in_mapping) %>%
    left_join(map1, by = c("sales_group_key","quality_key"))
  
  summarise_block <- function(df, group_col, out_name) {
    df %>%
      group_by(.grp = .data[[group_col]]) %>%
      summarise(
        Forecast_VAL = sum(Forecast_VAL, na.rm = TRUE),
        Actual_VAL   = sum(Actual_VAL,   na.rm = TRUE),
        .groups = "drop"
      ) %>%
      mutate(
        Gap_VAL = abs(Forecast_VAL - Actual_VAL),
        Gap_pct = ifelse(Forecast_VAL == 0, NA_real_, (Gap_VAL / Forecast_VAL) * 100)
      ) %>%
      rename(!!out_name := .grp) %>%
      round_whole()
  }
  
  sales_group_summary <- summarise_block(joined, "sales_group_key", "Sales_Group") %>%
    add_total_row("Sales_Group") %>%
    round_whole()
  
  # Processes: indicator cols OR processes column
  proc_cols_present <- intersect(names(joined), process_levels_clean)
  if (length(proc_cols_present) >= 3) {
    process_summary <- joined %>%
      pivot_longer(cols = all_of(proc_cols_present), names_to = "process_col", values_to = "flag") %>%
      mutate(
        flag2 = case_when(
          is.na(flag) ~ 0,
          is.numeric(flag) ~ as.integer(flag != 0),
          is.logical(flag) ~ as.integer(flag),
          TRUE ~ as.integer(str_trim(as.character(flag)) != "" &
                              str_to_upper(str_trim(as.character(flag))) != "N")
        )
      ) %>%
      filter(flag2 == 1) %>%
      group_by(process_col) %>%
      summarise(
        Forecast_VAL = sum(Forecast_VAL, na.rm = TRUE),
        Actual_VAL   = sum(Actual_VAL,   na.rm = TRUE),
        .groups = "drop"
      ) %>%
      mutate(Process = process_levels[match(process_col, process_levels_clean)]) %>%
      select(Process, Forecast_VAL, Actual_VAL)
  } else if (!is.na(map_processes_col)) {
    process_summary <- joined %>%
      mutate(processes_txt = as.character(.data[[map_processes_col]])) %>%
      separate_rows(processes_txt, sep = ",") %>%
      mutate(processes_txt = str_squish(str_trim(processes_txt))) %>%
      filter(!is.na(processes_txt), processes_txt != "") %>%
      group_by(processes_txt) %>%
      summarise(
        Forecast_VAL = sum(Forecast_VAL, na.rm = TRUE),
        Actual_VAL   = sum(Actual_VAL,   na.rm = TRUE),
        .groups = "drop"
      ) %>%
      rename(Process = processes_txt)
  } else stop("Mapping must have process info (indicator columns or 'Processes' column).")
  
  process_summary <- process_summary %>%
    mutate(
      Gap_VAL = abs(Forecast_VAL - Actual_VAL),
      Gap_pct = ifelse(Forecast_VAL == 0, NA_real_, (Gap_VAL / Forecast_VAL) * 100),
      Process = factor(Process, levels = process_levels)
    ) %>%
    arrange(Process) %>%
    round_whole()
  
  # Dye
  if (is.na(map_dye_col)) stop("Mapping missing Dye Machine column.")
  dye_summary <- joined %>%
    mutate(Dye_Machine = str_trim(as.character(.data[[map_dye_col]]))) %>%
    filter(!is.na(Dye_Machine), Dye_Machine != "") %>%
    group_by(Dye_Machine) %>%
    summarise(
      Forecast_VAL = sum(Forecast_VAL, na.rm = TRUE),
      Actual_VAL   = sum(Actual_VAL,   na.rm = TRUE),
      .groups = "drop"
    ) %>%
    mutate(
      Gap_VAL = abs(Forecast_VAL - Actual_VAL),
      Gap_pct = ifelse(Forecast_VAL == 0, NA_real_, (Gap_VAL / Forecast_VAL) * 100)
    ) %>%
    round_whole() %>%
    add_total_row("Dye_Machine") %>%
    round_whole()
  
  # Finishing
  if (is.na(map_finish_col)) stop("Mapping missing Finishing Machine column.")
  finish_summary <- joined %>%
    mutate(Finishing_Machine = str_trim(as.character(.data[[map_finish_col]]))) %>%
    filter(!is.na(Finishing_Machine), Finishing_Machine != "") %>%
    group_by(Finishing_Machine) %>%
    summarise(
      Forecast_VAL = sum(Forecast_VAL, na.rm = TRUE),
      Actual_VAL   = sum(Actual_VAL,   na.rm = TRUE),
      .groups = "drop"
    ) %>%
    mutate(
      Gap_VAL = abs(Forecast_VAL - Actual_VAL),
      Gap_pct = ifelse(Forecast_VAL == 0, NA_real_, (Gap_VAL / Forecast_VAL) * 100)
    ) %>%
    round_whole() %>%
    add_total_row("Finishing_Machine") %>%
    round_whole()
  
  # Material
  if (is.na(map_material_col)) stop("Mapping missing Material Composition column.")
  material_summary <- joined %>%
    mutate(Material = str_trim(as.character(.data[[map_material_col]]))) %>%
    filter(!is.na(Material), Material != "") %>%
    group_by(Material) %>%
    summarise(
      Forecast_VAL = sum(Forecast_VAL, na.rm = TRUE),
      Actual_VAL   = sum(Actual_VAL,   na.rm = TRUE),
      .groups = "drop"
    ) %>%
    mutate(
      Gap_VAL = abs(Forecast_VAL - Actual_VAL),
      Gap_pct = ifelse(Forecast_VAL == 0, NA_real_, (Gap_VAL / Forecast_VAL) * 100)
    ) %>%
    round_whole() %>%
    add_total_row("Material") %>%
    round_whole()
  
  # Preset Machines (exclude 0/"0")
  if (is.na(map_preset_col)) stop("Mapping missing Preset Machine column (e.g., Preset Machine).")
  preset_summary <- joined %>%
    mutate(
      Preset_Machine = str_trim(as.character(.data[[map_preset_col]])),
      Preset_Machine_key = norm_key(Preset_Machine)
    ) %>%
    filter(!is.na(Preset_Machine), Preset_Machine != "", Preset_Machine_key != "0") %>%
    group_by(Preset_Machine) %>%
    summarise(
      Forecast_VAL = sum(Forecast_VAL, na.rm = TRUE),
      Actual_VAL   = sum(Actual_VAL,   na.rm = TRUE),
      .groups = "drop"
    ) %>%
    mutate(
      Gap_VAL = abs(Forecast_VAL - Actual_VAL),
      Gap_pct = ifelse(Forecast_VAL == 0, NA_real_, (Gap_VAL / Forecast_VAL) * 100)
    ) %>%
    round_whole() %>%
    add_total_row("Preset_Machine") %>%
    round_whole()
  
  list(
    basis               = basis,
    quality_compare     = quality_compare,
    status_counts       = status_counts,
    unmatched_list      = unmatched_list,
    sales_group_summary = sales_group_summary,
    process_summary     = process_summary,
    preset_summary      = preset_summary,
    dye_summary         = dye_summary,
    finish_summary      = finish_summary,
    material_summary    = material_summary
  )
}

# ============================================================
# 6 MONTH PROJECTION HELPERS
# ============================================================
read_projection_6m <- function(proj_path) {
  x <- safe_read_excel(proj_path)
  
  sg_col <- pick_col(names(x), c("sales_group"))
  q_col  <- pick_col(names(x), c("quality"))
  if (is.na(sg_col) || is.na(q_col)) {
    stop("6 Month Projection file must contain columns: Sales Group, Quality (then 6 month columns).")
  }
  
  month_cols_all <- setdiff(names(x), c(sg_col, q_col))
  
  non_empty_cols <- month_cols_all[sapply(month_cols_all, function(cc) {
    v <- x[[cc]]
    any(!(is.na(v) | (as.character(v) %>% str_trim() == "")))
  })]
  
  if (length(non_empty_cols) < 6) {
    stop(paste0("6 Month Projection must have 6 month columns with values. Found only ", length(non_empty_cols), "."))
  }
  
  month_cols <- non_empty_cols[1:6]
  
  long <- x %>%
    mutate(
      Sales_Group = as.character(.data[[sg_col]]),
      Quality     = as.character(.data[[q_col]])
    ) %>%
    select(Sales_Group, Quality, all_of(month_cols)) %>%
    pivot_longer(cols = all_of(month_cols), names_to = "month_raw", values_to = "Forecast") %>%
    mutate(
      Forecast = suppressWarnings(as.numeric(Forecast)),
      Sales_Group_key = norm_key(Sales_Group),
      Quality_key     = norm_key(Quality),
      Month_label     = month_raw
    )
  
  month_order <- long %>%
    distinct(month_raw, Month_label) %>%
    mutate(idx = row_number()) %>%
    arrange(idx)
  
  long %>%
    left_join(month_order %>% select(month_raw, idx), by = "month_raw") %>%
    mutate(Month_label = factor(as.character(Month_label), levels = as.character(month_order$Month_label)))
}

get_projection_unmatched <- function(proj_long, map_path = MAP_PATH) {
  if (!file.exists(map_path)) stop(paste0("Mapping file not found at: ", map_path))
  map_raw <- safe_read_excel(map_path)
  
  m_sales <- pick_col(names(map_raw), c("sales_group"))
  m_qual  <- pick_col(names(map_raw), c("quality"))
  if (is.na(m_sales) || is.na(m_qual)) stop("Mapping file must contain: Sales Group and Quality")
  
  map_keys <- map_raw %>%
    transmute(
      Sales_Group_key = norm_key(.data[[m_sales]]),
      Quality_key     = norm_key(.data[[m_qual]])
    ) %>%
    distinct()
  
  proj_keys <- proj_long %>%
    transmute(Sales_Group_key = Sales_Group_key, Quality_key = Quality_key) %>%
    distinct()
  
  anti_join(proj_keys, map_keys, by = c("Sales_Group_key", "Quality_key")) %>%
    transmute(Sales_Group = Sales_Group_key, Quality = Quality_key) %>%
    arrange(Sales_Group, Quality)
}

get_mapping_one <- function(map_path = MAP_PATH) {
  if (!file.exists(map_path)) stop(paste0("Mapping file not found at: ", map_path))
  map_raw <- safe_read_excel(map_path)
  
  m_sales <- pick_col(names(map_raw), c("sales_group"))
  m_qual  <- pick_col(names(map_raw), c("quality"))
  if (is.na(m_sales) || is.na(m_qual)) stop("Mapping file must contain: Sales Group and Quality")
  
  map_raw %>%
    mutate(
      Sales_Group_key = norm_key(.data[[m_sales]]),
      Quality_key     = norm_key(.data[[m_qual]])
    ) %>%
    group_by(Sales_Group_key, Quality_key) %>%
    slice(1) %>%
    ungroup()
}

proj_sum_salesgroup <- function(proj_long, sg_keep = c("SG1","SG2","SG4","SG5","SG6")) {
  proj_long %>%
    filter(!is.na(Forecast)) %>%
    mutate(Sales_Group_key = norm_key(Sales_Group_key)) %>%
    filter(Sales_Group_key %in% sg_keep) %>%
    group_by(Month_label, Category = Sales_Group_key) %>%
    summarise(Forecast = sum(Forecast, na.rm = TRUE), .groups = "drop") %>%
    arrange(Month_label, Category)
}

proj_sum_processes <- function(proj_joined, map_raw) {
  proc_cols_present <- intersect(names(map_raw), process_levels_clean)
  map_processes_col <- pick_col(names(map_raw), c("processes","process","process_name","process_names"))
  
  if (length(proc_cols_present) >= 3) {
    proj_joined %>%
      pivot_longer(cols = all_of(proc_cols_present), names_to = "process_col", values_to = "flag") %>%
      mutate(
        flag2 = case_when(
          is.na(flag) ~ 0,
          is.numeric(flag) ~ as.integer(flag != 0),
          is.logical(flag) ~ as.integer(flag),
          TRUE ~ as.integer(str_trim(as.character(flag)) != "" &
                              str_to_upper(str_trim(as.character(flag))) != "N")
        )
      ) %>%
      filter(flag2 == 1) %>%
      mutate(Category = process_levels[match(process_col, process_levels_clean)]) %>%
      filter(!is.na(Category), Category != "") %>%
      group_by(Month_label, Category) %>%
      summarise(Forecast = sum(Forecast, na.rm = TRUE), .groups = "drop") %>%
      mutate(Category = factor(Category, levels = process_levels)) %>%
      arrange(Month_label, Category)
  } else if (!is.na(map_processes_col) && map_processes_col %in% names(proj_joined)) {
    proj_joined %>%
      mutate(processes_txt = as.character(.data[[map_processes_col]])) %>%
      separate_rows(processes_txt, sep = ",") %>%
      mutate(Category = str_squish(str_trim(processes_txt))) %>%
      filter(!is.na(Category), Category != "") %>%
      group_by(Month_label, Category) %>%
      summarise(Forecast = sum(Forecast, na.rm = TRUE), .groups = "drop") %>%
      arrange(Month_label, Category)
  } else {
    stop("Projection Processes: Mapping must have process info (indicator columns or 'Processes' column).")
  }
}

proj_sum_simple_category <- function(proj_joined, col_name, out_label = "Category", exclude0 = FALSE) {
  if (is.na(col_name) || !(col_name %in% names(proj_joined))) {
    stop(paste0("Projection Mapping missing required column: ", col_name))
  }
  d <- proj_joined %>%
    mutate(Category = str_trim(as.character(.data[[col_name]]))) %>%
    filter(!is.na(Category), Category != "")
  
  if (exclude0) {
    d <- d %>%
      mutate(Category_key = norm_key(Category)) %>%
      filter(Category_key != "0") %>%
      select(-Category_key)
  }
  
  d %>%
    group_by(Month_label, Category) %>%
    summarise(Forecast = sum(Forecast, na.rm = TRUE), .groups = "drop") %>%
    arrange(Month_label, Category) %>%
    rename(!!out_label := Category)
}

proj_make_wide_table <- function(d_long, category_name = "Category") {
  if (is.null(d_long) || nrow(d_long) == 0) return(data.frame(Note = "No data."))
  
  months <- levels(d_long$Month_label)
  if (is.null(months) || length(months) == 0) months <- unique(as.character(d_long$Month_label))
  
  dd <- d_long %>%
    mutate(
      Month = as.character(Month_label),
      Cat   = as.character(.data[[category_name]])
    ) %>%
    select(Cat, Month, Forecast) %>%
    group_by(Cat, Month) %>%
    summarise(Forecast = sum(Forecast, na.rm = TRUE), .groups = "drop")
  
  wide <- dd %>%
    mutate(Month = factor(Month, levels = months)) %>%
    arrange(Cat, Month) %>%
    pivot_wider(names_from = Month, values_from = Forecast, values_fill = 0)
  
  names(wide)[1] <- category_name
  round_whole(wide)
}

make_projection_line_plot <- function(d, title_text, top_n = 12, y_title = "Forecast (T)") {
  if (is.null(d) || nrow(d) == 0) {
    return(plotly_empty() %>% layout(title = plotly_title(paste0(title_text, " (No data)"))))
  }
  
  topcats <- d %>%
    group_by(Category) %>%
    summarise(Total = sum(Forecast, na.rm = TRUE), .groups = "drop") %>%
    arrange(desc(Total)) %>%
    slice_head(n = top_n) %>%
    pull(Category)
  
  dd <- d %>%
    filter(Category %in% topcats) %>%
    mutate(
      Month_label = factor(as.character(Month_label), levels = levels(d$Month_label)),
      Category = as.character(Category)
    )
  
  cats <- sort(unique(dd$Category))
  pal  <- grDevices::hcl.colors(max(length(cats), 3), palette = "Dark 3")
  col_map <- setNames(pal[seq_along(cats)], cats)
  
  plot_ly(
    data = dd,
    x = ~Month_label,
    y = ~Forecast,
    color = ~Category,
    colors = col_map,
    type = "scatter",
    mode = "lines+markers",
    marker = list(size = 8),
    line = list(width = 3),
    hovertemplate = "<b>%{customdata}</b><br>Month=%{x}<br>Forecast=%{y:.0f}<extra></extra>",
    customdata = ~Category
  ) %>%
    layout(
      title = plotly_title(title_text),
      xaxis = list(title = "Months", tickfont = list(size = 14)),
      yaxis = list(title = y_title, rangemode = "tozero", tickfont = list(size = 14)),
      paper_bgcolor = "white",
      plot_bgcolor  = "white",
      legend = list(orientation = "h", x = 0.01, y = -0.25),
      margin = list(t = 70, b = 120, l = 60, r = 20)
    ) %>%
    config(displayModeBar = TRUE, scrollZoom = FALSE)
}

chart_legend_note <- function(show_capacity = FALSE) {
  tagList(
    div(class="mini-note",
        tags$b("Chart legend (meaning):"),
        tags$ul(
          tags$li(tags$span(style=paste0("color:", COL_FORECAST, ";font-weight:900;"), "Forecast (Navy)"),
                  " = planned / expected production"),
          tags$li(tags$span(style=paste0("color:", COL_ACTUAL, ";font-weight:900;"), "Actual (Blue)"),
                  " = real production"),
          if (show_capacity) {
            tags$li(tags$span(style=paste0("color:", COL_CAPACITY, ";font-weight:900;"), "Max Capacity (Green background)"),
                    " = machine monthly capacity (top label shows Capacity - Actual)")
          },
          tags$li(tags$span(style="color:red;font-weight:900;", "Red circle"),
                  " = Actual is higher than Forecast")
        )
    )
  )
}

# ============================================================
# UI
# ============================================================
ui <- fluidPage(
  useShinyjs(),
  tags$head(tags$style(HTML("
    body { background-color: #f7f9fc; font-family: Arial; }
    .title-navy { color:#0A1F44; font-weight:900; }
    .btn, .btn-default { background:#0A1F44 !important; color:white !important; border:none !important; }
    .btn:hover { background:#4C6FA6 !important; }
    .nav-tabs > li > a { color:#0A1F44; font-weight:700; }
    .nav-tabs > li.active > a { background:#0A1F44 !important; color:white !important; }
    .eq-box { background:#ffffff; border:1px solid #e5e7eb; border-radius:10px; padding:10px 12px; }
    .mini-note { background:#ffffff; border:1px solid #e5e7eb; border-radius:10px; padding:10px 12px; margin-top:10px; }
    .purple-note { background:#F5F3FF; border:1px solid #E9D5FF; border-radius:10px; padding:10px 12px; margin-top:10px; }

    table.dataTable thead th {
      background: #0A1F44 !important;
      color: white !important;
      font-weight: 800 !important;
      border-bottom: 0 !important;
    }
    table.dataTable tbody tr:nth-child(odd)  { background-color: #F6F8FB !important; }
    table.dataTable tbody tr:nth-child(even) { background-color: #FFFFFF !important; }
    table.dataTable { border-radius: 10px; overflow: hidden; }
    .dataTables_wrapper .dataTables_filter input,
    .dataTables_wrapper .dataTables_length select {
      border: 1px solid #e5e7eb; border-radius: 8px; padding: 6px 10px;
    }
  "))),
  titlePanel(tags$span("HAYLEYS – Monthly Summary (Forecast vs Actual)", class = "title-navy")),
  
  tabsetPanel(
    id = "main_tab",
    
    tabPanel("Actual vs Forecast",
             sidebarLayout(
               sidebarPanel(
                 helpText("Mapping.xlsx is loaded automatically from your computer."),
                 hr(),
                 tags$div(
                   tags$b("Forecast file columns:"),
                   tags$ul(tags$li("Sales Group"), tags$li("Quality"), tags$li("Forecast"))
                 ),
                 tags$div(
                   tags$b("Actual file columns:"),
                   tags$ul(tags$li("Sales Group"), tags$li("Quality"), tags$li("Actual"))
                 ),
                 fileInput("forecast_file", "Upload Forecast Excel", accept = c(".xlsx")),
                 fileInput("actual_file", "Upload Actual Excel", accept = c(".xlsx")),
                 hr(),
                 radioButtons(
                   "analysis_basis",
                   "Analysis basis",
                   choices = c("GROSS" = "GROSS", "NET" = "NET"),
                   selected = "GROSS",
                   inline = TRUE
                 ),
                 helpText("If GROSS: Gross = (Net/0.9) * 1.02"),
                 hr(),
                 selectInput("month_name", "Month", choices = month.name,
                             selected = month.name[as.integer(format(Sys.Date(), "%m"))]),
                 dateInput("as_at_date", "As at date", value = Sys.Date()),
                 numericInput("days_in_month", "Days in month", value = 31, min = 28, max = 31),
                 hr(),
                 downloadButton("download_excel", "Download Output Excel")
               ),
               mainPanel(
                 tabsetPanel(
                   tabPanel("Tab 1: Quality",
                            plotOutput("total_plot", height = 320),
                            hr(),
                            plotlyOutput("status_plot", height = 380),
                            div(class="purple-note",
                                tags$b("Click a status bar to see the related qualities below."),
                                br(),
                                tags$span("Selected status will filter the table automatically.")
                            ),
                            br(),
                            DTOutput("status_detail_table")
                   ),
                   tabPanel("Unmatched List",
                            helpText("These Sales Group + Quality pairs are NOT in Mapping.xlsx. Please add them to Mapping."),
                            DTOutput("unmatched_table")
                   ),
                   tabPanel("Sales Group",
                            actionButton("btn_sales_style", "Switch Bar Style"),
                            br(), br(),
                            plotlyOutput("sales_plot", height = "560px"),
                            chart_legend_note(show_capacity = FALSE),
                            div(class="eq-box", HTML("<b>Gap%:</b> &nbsp; |Actual − Forecast| / Forecast × 100")),
                            div(class="mini-note",
                                tags$b("Removed categories (not shown in chart):"),
                                br(),
                                DTOutput("sales_removed_table")
                            ),
                            hr(),
                            DTOutput("sales_table")
                   ),
                   tabPanel("Processes",
                            actionButton("btn_process_style", "Switch Bar Style"),
                            br(), br(),
                            plotlyOutput("process_plot", height = "560px"),
                            chart_legend_note(show_capacity = FALSE),
                            div(class="eq-box", HTML("<b>Gap%:</b> &nbsp; |Actual − Forecast| / Forecast × 100")),
                            div(class="mini-note",
                                tags$b("Removed categories (not shown in chart):"),
                                br(),
                                DTOutput("process_removed_table")
                            ),
                            hr(),
                            DTOutput("process_table")
                   ),
                   tabPanel("Preset Machines",
                            actionButton("btn_preset_style", "Switch Bar Style"),
                            br(), br(),
                            plotlyOutput("preset_plot", height = "560px"),
                            chart_legend_note(show_capacity = TRUE),
                            div(class="eq-box", HTML("<b>Gap%:</b> &nbsp; |Actual − Forecast| / Forecast × 100")),
                            div(class="eq-box", HTML("<b>Monthly Max Capacity (T):</b> &nbsp; (kg/day × Days in month) / 1000")),
                            div(class="mini-note",
                                tags$b("Removed categories (not shown in chart):"),
                                br(),
                                DTOutput("preset_removed_table")
                            ),
                            hr(),
                            DTOutput("preset_table")
                   ),
                   tabPanel("Dye Machines",
                            actionButton("btn_dye_style", "Switch Bar Style"),
                            br(), br(),
                            plotlyOutput("dye_plot", height = "560px"),
                            chart_legend_note(show_capacity = TRUE),
                            div(class="eq-box", HTML("<b>Gap%:</b> &nbsp; |Actual − Forecast| / Forecast × 100")),
                            div(class="eq-box", HTML("<b>Monthly Max Capacity (T):</b> &nbsp; (kg/day × Days in month) / 1000")),
                            div(class="mini-note",
                                tags$b("Removed categories (not shown in chart):"),
                                br(),
                                DTOutput("dye_removed_table")
                            ),
                            hr(),
                            DTOutput("dye_table")
                   ),
                   tabPanel("Finishing Machines",
                            actionButton("btn_finish_style", "Switch Bar Style"),
                            br(), br(),
                            plotlyOutput("finish_plot", height = "560px"),
                            chart_legend_note(show_capacity = TRUE),
                            div(class="eq-box", HTML("<b>Gap%:</b> &nbsp; |Actual − Forecast| / Forecast × 100")),
                            div(class="eq-box", HTML("<b>Monthly Max Capacity (T):</b> &nbsp; (kg/day × Days in month) / 1000")),
                            div(class="mini-note",
                                tags$b("Removed categories (not shown in chart):"),
                                br(),
                                DTOutput("finish_removed_table")
                            ),
                            hr(),
                            DTOutput("finish_table")
                   ),
                   tabPanel("Material",
                            actionButton("btn_material_style", "Switch Bar Style"),
                            br(), br(),
                            plotlyOutput("material_plot", height = "560px"),
                            chart_legend_note(show_capacity = FALSE),
                            div(class="eq-box", HTML("<b>Gap%:</b> &nbsp; |Actual − Forecast| / Forecast × 100")),
                            div(class="mini-note",
                                tags$b("Removed categories (not shown in chart):"),
                                br(),
                                DTOutput("material_removed_table")
                            ),
                            hr(),
                            DTOutput("material_table")
                   )
                 )
               )
             )
    ),
    
    tabPanel("6 Month Projection",
             sidebarLayout(
               sidebarPanel(
                 helpText("Mapping.xlsx is loaded automatically from your computer."),
                 hr(),
                 tags$div(
                   tags$b("Projection file columns:"),
                   tags$ul(
                     tags$li("Sales Group"),
                     tags$li("Quality"),
                     tags$li("Then 6 month columns (ONLY 6 non-empty)")
                   )
                 ),
                 fileInput("proj6_file", "Upload 6 Month Projection Excel", accept = c(".xlsx")),
                 helpText("The app checks Sales Group + Quality pairs vs Mapping.xlsx and shows unmatched pairs.")
               ),
               mainPanel(
                 tabsetPanel(
                   tabPanel("Unmatched",
                            helpText("Sales Group + Quality pairs in projection but NOT in Mapping.xlsx."),
                            DTOutput("proj_unmatched_table")
                   ),
                   tabPanel("Sales Group",
                            helpText("Rows = Sales Groups, Columns = Months (6). Line chart colors represent categories."),
                            plotlyOutput("proj_salesgroup_plot", height = "560px"),
                            hr(),
                            DTOutput("proj_salesgroup_table")
                   ),
                   tabPanel("Processes",
                            helpText("Rows = Processes, Columns = Months (6). Line chart colors represent categories."),
                            plotlyOutput("proj_process_plot", height = "560px"),
                            hr(),
                            DTOutput("proj_process_table")
                   ),
                   tabPanel("Preset",
                            helpText("Rows = Preset Machines, Columns = Months (6). (0 excluded)."),
                            plotlyOutput("proj_preset_plot", height = "560px"),
                            hr(),
                            DTOutput("proj_preset_table")
                   ),
                   tabPanel("Dye",
                            helpText("Rows = Dye Machines, Columns = Months (6)."),
                            plotlyOutput("proj_dye_plot", height = "560px"),
                            hr(),
                            DTOutput("proj_dye_table")
                   ),
                   tabPanel("Finishing",
                            helpText("Rows = Finishing Machines, Columns = Months (6)."),
                            plotlyOutput("proj_finish_plot", height = "560px"),
                            hr(),
                            DTOutput("proj_finish_table")
                   ),
                   tabPanel("Material",
                            helpText("Rows = Material, Columns = Months (6)."),
                            plotlyOutput("proj_material_plot", height = "560px"),
                            hr(),
                            DTOutput("proj_material_table")
                   )
                 )
               )
             )
    )
  )
)

# ============================================================
# SERVER
# ============================================================
server <- function(input, output, session) {
  
  # axis title (kept), but unit text for hover/labels is only "T"
  y_axis_title <- reactive({
    if (input$analysis_basis == "NET") "NET (T)" else "GROSS (T)"
  })
  y_unit <- reactive("T")
  
  results <- reactive({
    req(input$forecast_file, input$actual_file)
    build_all_summaries(
      forecast_path = input$forecast_file$datapath,
      actual_path   = input$actual_file$datapath,
      days_in_month = input$days_in_month,
      basis         = input$analysis_basis,
      map_path      = MAP_PATH
    )
  })
  
  # ----------------------------
  # Style toggles
  # ----------------------------
  sales_style    <- reactiveVal(1)
  process_style  <- reactiveVal(1)
  preset_style   <- reactiveVal(1)
  dye_style      <- reactiveVal(1)
  finish_style   <- reactiveVal(1)
  material_style <- reactiveVal(1)
  
  observeEvent(input$btn_sales_style,    { sales_style(ifelse(sales_style() == 1, 2, 1)) })
  observeEvent(input$btn_process_style,  { process_style(ifelse(process_style() == 1, 2, 1)) })
  observeEvent(input$btn_preset_style,   { preset_style(ifelse(preset_style() == 1, 2, 1)) })
  observeEvent(input$btn_dye_style,      { dye_style(ifelse(dye_style() == 1, 2, 1)) })
  observeEvent(input$btn_finish_style,   { finish_style(ifelse(finish_style() == 1, 2, 1)) })
  observeEvent(input$btn_material_style, { material_style(ifelse(material_style() == 1, 2, 1)) })
  
  # ----------------------------
  # Total plot (base R)
  # ----------------------------
  output$total_plot <- renderPlot({
    req(results())
    sg <- results()$sales_group_summary
    total_row <- sg %>% filter(Sales_Group == "TOTAL")
    total_fore <- as.numeric(total_row$Forecast_VAL[1])
    total_act  <- as.numeric(total_row$Actual_VAL[1])
    
    vals <- c(total_fore, total_act)
    labs <- c("Forecast", "Actual")
    
    par(cex.main = 1.6, cex.axis = 1.2, cex.lab = 1.3)
    bp <- barplot(
      vals,
      names.arg = labs,
      main = paste0("Total ", y_axis_title()),
      ylim = c(0, max(vals, na.rm = TRUE) * 1.25),
      ylab = y_axis_title(),
      xlab = "",
      col = c(COL_FORECAST, COL_ACTUAL)
    )
    
    text(
      x = bp,
      y = vals * 0.5,
      labels = format(round(vals, 0), big.mark = ","),
      col = "white",
      cex = 1.4,
      font = 2
    )
    
    legend(
      "topright",
      legend = c("Forecast", "Actual"),
      fill = c(COL_FORECAST, COL_ACTUAL),
      bty = "n",
      cex = 1.05
    )
  })
  
  # ----------------------------
  # Quality Status clickable plotly + filtered table
  # ----------------------------
  selected_status <- reactiveVal(NULL)
  ALLOW_STATUS <- c("Forecast > Actual","Forecast < Actual","Forecast = Actual")
  
  output$status_plot <- renderPlotly({
    req(results())
    
    df_all <- results()$status_counts
    df <- df_all %>%
      filter(as.character(Status) %in% ALLOW_STATUS) %>%
      mutate(Status = factor(as.character(Status), levels = ALLOW_STATUS)) %>%
      arrange(Status)
    
    if (nrow(df) == 0) {
      return(plotly_empty() %>% layout(title = plotly_title("Quality Status Count")))
    }
    
    cols <- unname(QSTAT_TXT[as.character(df$Status)])
    
    plot_ly(
      data = df,
      x = ~as.character(Status),
      y = ~n,
      type = "bar",
      source = "status_bar",
      customdata = ~as.character(Status),
      hovertemplate = "<b>%{x}</b><br>N=%{y}<extra></extra>",
      marker = list(color = cols, line = list(color = "black", width = 1.4))
    ) %>%
      layout(
        title = plotly_title("Quality Status Count (Click a bar)"),
        xaxis = list(title = "", tickangle = 0, tickfont = list(size = 14)),
        yaxis = list(title = "Number of Pairs", rangemode = "tozero", tickfont = list(size = 14)),
        paper_bgcolor = "white",
        plot_bgcolor  = "white",
        margin = list(t = 60, b = 60, l = 60, r = 20)
      ) %>%
      config(displayModeBar = TRUE, scrollZoom = FALSE)
  })
  
  observeEvent(event_data("plotly_click", source = "status_bar"), {
    ed <- event_data("plotly_click", source = "status_bar")
    if (!is.null(ed) && nrow(ed) > 0) selected_status(as.character(ed$x[1]))
  })
  
  output$status_detail_table <- renderDT({
    req(results())
    qc <- results()$quality_compare
    
    st <- selected_status()
    if (is.null(st) || !(st %in% ALLOW_STATUS)) {
      df_show <- data.frame(Note = "Click one bar: Forecast > Actual / Forecast < Actual / Forecast = Actual")
      return(
        datatable(df_show, rownames = FALSE, options = list(dom = "t")) %>%
          formatStyle(columns = names(df_show),
                      backgroundColor = "#F5F3FF", color = "#4C1D95",
                      fontWeight = "bold")
      )
    }
    
    df_show <- qc %>%
      filter(Status == st) %>%
      transmute(Sales_Group, Quality, Forecast, Actual, Status) %>%
      arrange(desc(Forecast + Actual))
    
    if (nrow(df_show) == 0) {
      df_show <- data.frame(Note = paste0("No rows found for: ", st))
      return(
        datatable(df_show, rownames = FALSE, options = list(dom = "t")) %>%
          formatStyle(columns = names(df_show),
                      backgroundColor = "#F5F3FF", color = "#4C1D95",
                      fontWeight = "bold")
      )
    }
    
    bg  <- unname(QSTAT_BG[st])
    txt <- unname(QSTAT_TXT[st])
    
    datatable(
      df_show,
      rownames = FALSE,
      class = "stripe hover cell-border compact",
      options = list(pageLength = 12, scrollX = TRUE, dom = "lftip")
    ) %>%
      formatStyle(columns = names(df_show),
                  backgroundColor = bg, color = txt,
                  `border-color` = "#E5E7EB",
                  `font-family` = "Arial",
                  `font-size` = "13px",
                  `padding` = "8px 10px") %>%
      formatStyle("Status", fontWeight = "bold", color = txt) %>%
      formatRound(columns = c("Forecast","Actual"), digits = 0)
  })
  
  # ----------------------------
  # Tables (Actual vs Forecast)
  # ----------------------------
  output$unmatched_table <- renderDT({ req(results()); make_dt(results()$unmatched_list, page_len = 25, type = "generic") })
  output$sales_table     <- renderDT({ req(results()); make_dt(results()$sales_group_summary, page_len = 25, first_col_for_total = "Sales_Group") })
  output$process_table   <- renderDT({ req(results()); make_dt(results()$process_summary, page_len = 25, first_col_for_total = "Process") })
  output$preset_table    <- renderDT({ req(results()); make_dt(results()$preset_summary, page_len = 25, first_col_for_total = "Preset_Machine") })
  output$dye_table       <- renderDT({ req(results()); make_dt(results()$dye_summary, page_len = 25, first_col_for_total = "Dye_Machine") })
  output$finish_table    <- renderDT({ req(results()); make_dt(results()$finish_summary, page_len = 25, first_col_for_total = "Finishing_Machine") })
  output$material_table  <- renderDT({ req(results()); make_dt(results()$material_summary, page_len = 25, first_col_for_total = "Material") })
  
  # ----------------------------
  # Removed tables (Forecast==0)
  # ----------------------------
  output$sales_removed_table    <- renderDT({ req(results()); make_dt(get_removed_table(results()$sales_group_summary, "Sales_Group"), page_len = 6, type = "removed") })
  output$process_removed_table  <- renderDT({ req(results()); make_dt(get_removed_table(results()$process_summary, "Process"), page_len = 6, type = "removed") })
  output$preset_removed_table   <- renderDT({ req(results()); make_dt(get_removed_table(results()$preset_summary, "Preset_Machine"), page_len = 6, type = "removed") })
  output$dye_removed_table      <- renderDT({ req(results()); make_dt(get_removed_table(results()$dye_summary, "Dye_Machine"), page_len = 6, type = "removed") })
  output$finish_removed_table   <- renderDT({ req(results()); make_dt(get_removed_table(results()$finish_summary, "Finishing_Machine"), page_len = 6, type = "removed") })
  output$material_removed_table <- renderDT({ req(results()); make_dt(get_removed_table(results()$material_summary, "Material"), page_len = 6, type = "removed") })
  
  # ----------------------------
  # Plotly charts (Actual vs Forecast) with style switch + capacity
  # ----------------------------
  output$sales_plot <- renderPlotly({
    req(results())
    ord <- compute_salesgroup_order(results()$sales_group_summary, "Sales_Group")
    if (sales_style() == 1) {
      make_plotly_overlay_bars(results()$sales_group_summary, "Sales_Group",
                               title_prefix = paste0("Sales Group (", y_axis_title(), ")"),
                               label_order = ord,
                               yaxis_title = y_axis_title(),
                               yunit = y_unit()
      )
    } else {
      make_plotly_twobars_bracket(results()$sales_group_summary, "Sales_Group",
                                  title_prefix = paste0("Sales Group – Difference (", y_axis_title(), ")"),
                                  label_order = ord,
                                  yaxis_title = y_axis_title(),
                                  yunit = y_unit()
      )
    }
  })
  
  output$process_plot <- renderPlotly({
    req(results())
    ord <- compute_chart_order(results()$process_summary, "Process")
    if (process_style() == 1) {
      make_plotly_overlay_bars(results()$process_summary, "Process",
                               title_prefix = paste0("Processes (", y_axis_title(), ")"),
                               label_order = ord,
                               yaxis_title = y_axis_title(),
                               yunit = y_unit()
      )
    } else {
      make_plotly_twobars_bracket(results()$process_summary, "Process",
                                  title_prefix = paste0("Processes – Difference (", y_axis_title(), ")"),
                                  label_order = ord,
                                  yaxis_title = y_axis_title(),
                                  yunit = y_unit()
      )
    }
  })
  
  # PRESET: capacity behind + cap gap label
  output$preset_plot <- renderPlotly({
    req(results())
    ord <- compute_chart_order(results()$preset_summary, "Preset_Machine")
    
    per_day_caps_kg_preset <- c(
      `TY-04`  = 7000, `TY 04` = 7000, `TY04` = 7000, `TY_04` = 7000,
      `10BAY`  = 10000, `10 BAY` = 10000, `10-BAY` = 10000, `10_BAY` = 10000,
      `TY-03`  = 5000, `TY 03` = 5000, `TY03` = 5000, `TY_03` = 5000
    )
    max_caps_T_preset <- (per_day_caps_kg_preset * input$days_in_month) / 1000
    
    if (preset_style() == 1) {
      make_plotly_overlay_bars(results()$preset_summary, "Preset_Machine",
                               title_prefix = paste0("Preset Machines (", y_axis_title(), ")"),
                               label_order = ord,
                               yaxis_title = y_axis_title(),
                               yunit = y_unit(),
                               max_caps = max_caps_T_preset
      )
    } else {
      make_plotly_twobars_bracket(results()$preset_summary, "Preset_Machine",
                                  title_prefix = paste0("Preset Machines – Difference (", y_axis_title(), ")"),
                                  label_order = ord,
                                  yaxis_title = y_axis_title(),
                                  yunit = y_unit()
      )
    }
  })
  
  output$dye_plot <- renderPlotly({
    req(results())
    ord <- compute_chart_order(results()$dye_summary, "Dye_Machine")
    
    per_day_caps_kg <- c(
      `SC Snaging Free` = 5000,
      `DC`             = 1000,
      `TH`             = 12000,
      `SC`             = 7400,
      `FG`             = 2000,
      `SC 12 &  SC 04` = 2700
    )
    max_caps_T <- (per_day_caps_kg * input$days_in_month) / 1000
    
    if (dye_style() == 1) {
      make_plotly_overlay_bars(results()$dye_summary, "Dye_Machine",
                               title_prefix = paste0("Dye Machines (", y_axis_title(), ")"),
                               label_order = ord,
                               yaxis_title = y_axis_title(),
                               yunit = y_unit(),
                               max_caps = max_caps_T
      )
    } else {
      make_plotly_twobars_bracket(results()$dye_summary, "Dye_Machine",
                                  title_prefix = paste0("Dye Machines – Difference (", y_axis_title(), ")"),
                                  label_order = ord,
                                  yaxis_title = y_axis_title(),
                                  yunit = y_unit()
      )
    }
  })
  
  output$finish_plot <- renderPlotly({
    req(results())
    ord <- compute_chart_order(results()$finish_summary, "Finishing_Machine")
    
    per_day_caps_kg_finish <- c(
      `T/Y 01 & T/Y 02` = 10000,
      `B 01 & B 02`     = 10000,
      `10 BAY & T/Y 05` = 15000,
      `Santex - Compact & B 02 - Compact`  = 22000,
      `Santex - Compact`= 18000,
      `T/Y 03`          = 5000
      )
    max_caps_T_finish <- (per_day_caps_kg_finish * input$days_in_month) / 1000
    
    if (finish_style() == 1) {
      make_plotly_overlay_bars(results()$finish_summary, "Finishing_Machine",
                               title_prefix = paste0("Finishing Machines (", y_axis_title(), ")"),
                               label_order = ord,
                               yaxis_title = y_axis_title(),
                               yunit = y_unit(),
                               max_caps = max_caps_T_finish
      )
    } else {
      make_plotly_twobars_bracket(results()$finish_summary, "Finishing_Machine",
                                  title_prefix = paste0("Finishing – Difference (", y_axis_title(), ")"),
                                  label_order = ord,
                                  yaxis_title = y_axis_title(),
                                  yunit = y_unit()
      )
    }
  })
  
  output$material_plot <- renderPlotly({
    req(results())
    ord <- compute_chart_order(results()$material_summary, "Material")
    if (material_style() == 1) {
      make_plotly_overlay_bars(results()$material_summary, "Material",
                               title_prefix = paste0("Material (", y_axis_title(), ")"),
                               label_order = ord,
                               yaxis_title = y_axis_title(),
                               yunit = y_unit()
      )
    } else {
      make_plotly_twobars_bracket(results()$material_summary, "Material",
                                  title_prefix = paste0("Material – Difference (", y_axis_title(), ")"),
                                  label_order = ord,
                                  yaxis_title = y_axis_title(),
                                  yunit = y_unit()
      )
    }
  })
  
  # ----------------------------
  # Download Excel (Actual vs Forecast)
  # ----------------------------
  output$download_excel <- downloadHandler(
    filename = function() {
      paste0("Output_", input$month_name, "_", format(input$as_at_date, "%Y-%m-%d"), "_", input$analysis_basis, ".xlsx")
    },
    content = function(file) {
      req(results())
      wb <- createWorkbook()
      
      addWorksheet(wb, "Quality_Comparison");  writeDataTable(wb, "Quality_Comparison", results()$quality_compare)
      addWorksheet(wb, "Sales_Group");         writeDataTable(wb, "Sales_Group", results()$sales_group_summary)
      addWorksheet(wb, "Processes");           writeDataTable(wb, "Processes", results()$process_summary)
      addWorksheet(wb, "Preset_Machines");     writeDataTable(wb, "Preset_Machines", results()$preset_summary)
      addWorksheet(wb, "Dye_Machines");        writeDataTable(wb, "Dye_Machines", results()$dye_summary)
      addWorksheet(wb, "Finishing_Machines");  writeDataTable(wb, "Finishing_Machines", results()$finish_summary)
      addWorksheet(wb, "Material");            writeDataTable(wb, "Material", results()$material_summary)
      
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
  
  # ============================================================
  # 6 Month Projection server (unchanged)
  # ============================================================
  proj_long <- reactive({
    req(input$proj6_file)
    validate(need(file.exists(MAP_PATH), paste0("Mapping file not found: ", MAP_PATH)))
    tryCatch(
      read_projection_6m(input$proj6_file$datapath),
      error = function(e) stop(paste0("Projection file error: ", conditionMessage(e)))
    )
  })
  
  proj_unmatched <- reactive({
    req(proj_long())
    tryCatch(
      get_projection_unmatched(proj_long(), map_path = MAP_PATH),
      error = function(e) stop(paste0("Mapping check error: ", conditionMessage(e)))
    )
  })
  
  output$proj_unmatched_table <- renderDT({
    req(input$proj6_file)
    validate(need(file.exists(MAP_PATH), paste0("Mapping file not found: ", MAP_PATH)))
    
    df <- proj_unmatched()
    if (nrow(df) == 0) {
      return(
        datatable(
          data.frame(Note = "All Sales Group + Quality pairs match Mapping.xlsx ✅"),
          rownames = FALSE, options = list(dom = "t")
        ) %>%
          formatStyle(columns = "Note",
                      backgroundColor = "#ECFDF5", color = "#065F46",
                      fontWeight = "bold", `font-size` = "14px")
      )
    }
    make_dt(df, page_len = 25, type = "generic")
  })
  
  proj_map_one <- reactive({
    req(input$proj6_file)
    validate(need(file.exists(MAP_PATH), paste0("Mapping file not found: ", MAP_PATH)))
    tryCatch(
      get_mapping_one(MAP_PATH),
      error = function(e) stop(paste0("Mapping read error: ", conditionMessage(e)))
    )
  })
  
  proj_joined <- reactive({
    req(proj_long(), proj_map_one())
    proj_long() %>%
      inner_join(proj_map_one(),
                 by = c("Sales_Group_key" = "Sales_Group_key", "Quality_key" = "Quality_key"))
  })
  
  proj_map_cols <- reactive({
    mr <- proj_map_one()
    list(
      material  = pick_col(names(mr), c("material_composition","material","composition","material_comp")),
      finishing = pick_col(names(mr), c("finishing_machine","finishing_machines","finishing","finish_machine")),
      dye       = pick_col(names(mr), c("dye_machine","dye_machines","dye","dyeing_machine")),
      preset    = pick_col(names(mr), c("preset_machine","preset_machines","preset","preset_machine_name","preset_machine_type"))
    )
  })
  
  proj_salesgroup_df <- reactive({
    req(proj_long())
    proj_sum_salesgroup(proj_long(), sg_keep = c("SG1","SG2","SG4","SG5","SG6"))
  })
  
  output$proj_salesgroup_plot <- renderPlotly({
    req(input$proj6_file)
    validate(need(file.exists(MAP_PATH), paste0("Mapping file not found: ", MAP_PATH)))
    tryCatch(
      make_projection_line_plot(
        d = proj_salesgroup_df() %>% select(Month_label, Category, Forecast),
        title_text = "Projection – Sales Group",
        top_n = 12
      ),
      error = function(e) plotly_empty() %>% layout(title = plotly_title(paste0("Projection chart error: ", conditionMessage(e))))
    )
  })
  
  output$proj_salesgroup_table <- renderDT({
    req(proj_salesgroup_df())
    wide <- proj_make_wide_table(
      d_long = proj_salesgroup_df() %>% mutate(Category = as.character(Category)),
      category_name = "Category"
    )
    make_dt(wide, page_len = 25, type = "generic")
  })
  
  proj_process_df <- reactive({
    req(proj_joined(), proj_map_one())
    tryCatch(
      proj_sum_processes(proj_joined(), map_raw = proj_map_one()),
      error = function(e) stop(paste0("Projection Processes error: ", conditionMessage(e)))
    ) %>%
      mutate(Category = as.character(Category))
  })
  
  output$proj_process_plot <- renderPlotly({
    req(input$proj6_file)
    validate(need(file.exists(MAP_PATH), paste0("Mapping file not found: ", MAP_PATH)))
    tryCatch(
      make_projection_line_plot(
        d = proj_process_df() %>% select(Month_label, Category, Forecast),
        title_text = "Projection – Processes",
        top_n = 12
      ),
      error = function(e) plotly_empty() %>% layout(title = plotly_title(paste0("Process chart error: ", conditionMessage(e))))
    )
  })
  
  output$proj_process_table <- renderDT({
    req(proj_process_df())
    wide <- proj_make_wide_table(d_long = proj_process_df(), category_name = "Category")
    names(wide)[1] <- "Process"
    make_dt(wide, page_len = 25, type = "generic")
  })
  
  proj_preset_df <- reactive({
    req(proj_joined(), proj_map_cols())
    cols <- proj_map_cols()
    tryCatch(
      proj_sum_simple_category(proj_joined(), col_name = cols$preset, out_label = "Preset_Machine", exclude0 = TRUE) %>%
        rename(Category = Preset_Machine),
      error = function(e) stop(paste0("Projection Preset error: ", conditionMessage(e)))
    )
  })
  
  output$proj_preset_plot <- renderPlotly({
    req(input$proj6_file)
    validate(need(file.exists(MAP_PATH), paste0("Mapping file not found: ", MAP_PATH)))
    tryCatch(
      make_projection_line_plot(
        d = proj_preset_df() %>% select(Month_label, Category, Forecast),
        title_text = "Projection – Preset",
        top_n = 12
      ),
      error = function(e) plotly_empty() %>% layout(title = plotly_title(paste0("Preset chart error: ", conditionMessage(e))))
    )
  })
  
  output$proj_preset_table <- renderDT({
    req(proj_preset_df())
    wide <- proj_make_wide_table(d_long = proj_preset_df(), category_name = "Category")
    names(wide)[1] <- "Preset_Machine"
    make_dt(wide, page_len = 25, type = "generic")
  })
  
  proj_dye_df <- reactive({
    req(proj_joined(), proj_map_cols())
    cols <- proj_map_cols()
    tryCatch(
      proj_sum_simple_category(proj_joined(), col_name = cols$dye, out_label = "Dye_Machine", exclude0 = FALSE) %>%
        rename(Category = Dye_Machine),
      error = function(e) stop(paste0("Projection Dye error: ", conditionMessage(e)))
    )
  })
  
  output$proj_dye_plot <- renderPlotly({
    req(input$proj6_file)
    validate(need(file.exists(MAP_PATH), paste0("Mapping file not found: ", MAP_PATH)))
    tryCatch(
      make_projection_line_plot(
        d = proj_dye_df() %>% select(Month_label, Category, Forecast),
        title_text = "Projection – Dye",
        top_n = 12
      ),
      error = function(e) plotly_empty() %>% layout(title = plotly_title(paste0("Dye chart error: ", conditionMessage(e))))
    )
  })
  
  output$proj_dye_table <- renderDT({
    req(proj_dye_df())
    wide <- proj_make_wide_table(d_long = proj_dye_df(), category_name = "Category")
    names(wide)[1] <- "Dye_Machine"
    make_dt(wide, page_len = 25, type = "generic")
  })
  
  proj_finish_df <- reactive({
    req(proj_joined(), proj_map_cols())
    cols <- proj_map_cols()
    tryCatch(
      proj_sum_simple_category(proj_joined(), col_name = cols$finishing, out_label = "Finishing_Machine", exclude0 = FALSE) %>%
        rename(Category = Finishing_Machine),
      error = function(e) stop(paste0("Projection Finishing error: ", conditionMessage(e)))
    )
  })
  
  output$proj_finish_plot <- renderPlotly({
    req(input$proj6_file)
    validate(need(file.exists(MAP_PATH), paste0("Mapping file not found: ", MAP_PATH)))
    tryCatch(
      make_projection_line_plot(
        d = proj_finish_df() %>% select(Month_label, Category, Forecast),
        title_text = "Projection – Finishing",
        top_n = 12
      ),
      error = function(e) plotly_empty() %>% layout(title = plotly_title(paste0("Finishing chart error: ", conditionMessage(e))))
    )
  })
  
  output$proj_finish_table <- renderDT({
    req(proj_finish_df())
    wide <- proj_make_wide_table(d_long = proj_finish_df(), category_name = "Category")
    names(wide)[1] <- "Finishing_Machine"
    make_dt(wide, page_len = 25, type = "generic")
  })
  
  proj_material_df <- reactive({
    req(proj_joined(), proj_map_cols())
    cols <- proj_map_cols()
    tryCatch(
      proj_sum_simple_category(proj_joined(), col_name = cols$material, out_label = "Material", exclude0 = FALSE) %>%
        rename(Category = Material),
      error = function(e) stop(paste0("Projection Material error: ", conditionMessage(e)))
    )
  })
  
  output$proj_material_plot <- renderPlotly({
    req(input$proj6_file)
    validate(need(file.exists(MAP_PATH), paste0("Mapping file not found: ", MAP_PATH)))
    tryCatch(
      make_projection_line_plot(
        d = proj_material_df() %>% select(Month_label, Category, Forecast),
        title_text = "Projection – Material",
        top_n = 12
      ),
      error = function(e) plotly_empty() %>% layout(title = plotly_title(paste0("Material chart error: ", conditionMessage(e))))
    )
  })
  
  output$proj_material_table <- renderDT({
    req(proj_material_df())
    wide <- proj_make_wide_table(d_long = proj_material_df(), category_name = "Category")
    names(wide)[1] <- "Material"
    make_dt(wide, page_len = 25, type = "generic")
  })
}

shinyApp(ui, server)
