library(tidyverse)
library(pdftools)
library(openxlsx)

path_folder_source <- "~/Data/esistafe/relatoriofase2026/"


extract_sistafe_table <- function(path_pdf) {
  raw_text <- pdftools::pdf_text(path_pdf)

  # ---- helpers ----
  normalize_pt_date <- function(x) {
    x <- as.character(x)
    x <- stringr::str_replace_all(x, "\\s*", "")
    dplyr::na_if(x, "")
  }

  extract_header_date <- function(txt, label_regex) {
    m <- stringr::str_match(
      txt,
      paste0(
        "(?s)\\b",
        label_regex,
        "\\s*:?\\s*(\\d{2}\\s*/\\s*\\d{2}\\s*/\\s*\\d{4})\\b"
      )
    )
    normalize_pt_date(m[, 2])
  }

  # ---- unidade_gestao ----
  unidade_gestao <- raw_text[1] |>
    stringr::str_extract("Gestão:\\s*(.+)") |>
    stringr::str_remove("Gestão:\\s*") |>
    stringr::str_trim()

  # ---- header dates (preferred for SALDO rows) ----
  header_data_chr <- extract_header_date(raw_text[1], "Data(?!\\s*Final)")
  header_data_final_chr <- extract_header_date(raw_text[1], "Data\\s*Final")

  header_data <- suppressWarnings(lubridate::dmy(header_data_chr))
  header_data_final <- suppressWarnings(lubridate::dmy(header_data_final_chr))

  # ---- header saldo (fallback when no transactions) ----
  saldo_hdr <- raw_text[1] |>
    stringr::str_extract("Saldo:\\s*([\\d\\.]+,\\d{2})") |>
    stringr::str_remove("Saldo:\\s*")

  saldo_hdr_dc <- raw_text[1] |>
    stringr::str_extract("Saldo:\\s*[\\d\\.]+,\\d{2}\\s*([CD])") |>
    stringr::str_extract("[CD]$")

  saldo_hdr_num <- readr::parse_number(
    saldo_hdr,
    locale = readr::locale(decimal_mark = ",", grouping_mark = ".")
  )

  # ---- Extract transaction lines ----
  lines <- raw_text |>
    stringr::str_split("\n") |>
    unlist() |>
    stringr::str_subset("^\\d{2}\\s*/\\s*\\d{2}\\s*/\\s*\\d{4}") |>
    stringr::str_squish()

  # ---- No transactions: SALDO rows only, using header dates ----
  if (length(lines) == 0) {
    message(
      "No transactions found in: ",
      basename(path_pdf),
      " — returning saldos only"
    )

    return(
      dplyr::bind_rows(
        tibble::tibble(
          unidade_gestao = unidade_gestao,
          data = header_data,
          tipo = "SALDO_INICIAL",
          codigo_documento = NA_character_,
          valor_lancamento = 0,
          dc1 = NA_character_,
          saldo_atual = saldo_hdr_num,
          dc2 = saldo_hdr_dc,
          saldo_inicial_fim = saldo_hdr_num
        ),
        tibble::tibble(
          unidade_gestao = unidade_gestao,
          data = header_data_final,
          tipo = "SALDO_FINAL",
          codigo_documento = NA_character_,
          valor_lancamento = 0,
          dc1 = NA_character_,
          saldo_atual = saldo_hdr_num,
          dc2 = saldo_hdr_dc,
          saldo_inicial_fim = saldo_hdr_num
        )
      )
    )
  }

  # ---- Parse transactions ----
  df <- tibble::tibble(raw = lines) |>
    tidyr::separate(
      raw,
      into = c(
        "data",
        "codigo_documento",
        "valor_lancamento",
        "dc1",
        "saldo_atual",
        "dc2"
      ),
      sep = "\\s+",
      fill = "right"
    ) |>
    dplyr::mutate(
      # normalize date like "01 / 12 / 2025" -> "01/12/2025"
      data = normalize_pt_date(data),
      data = lubridate::dmy(data),

      valor_lancamento = readr::parse_number(
        valor_lancamento,
        locale = readr::locale(decimal_mark = ",", grouping_mark = ".")
      ),
      saldo_atual = readr::parse_number(
        saldo_atual,
        locale = readr::locale(decimal_mark = ",", grouping_mark = ".")
      ),

      # C => negative, D => positive
      valor_lancamento = dplyr::if_else(
        dc1 == "C",
        -valor_lancamento,
        valor_lancamento
      ),

      unidade_gestao = unidade_gestao,
      tipo = "MOVIMENTO",
      saldo_inicial_fim = NA_real_
    ) |>
    dplyr::select(
      unidade_gestao,
      data,
      tipo,
      codigo_documento,
      valor_lancamento,
      dc1,
      saldo_atual,
      dc2,
      saldo_inicial_fim
    )

  # Fallback if header dates are missing for any reason
  data_inicio <- if (!is.na(header_data)) header_data else df$data[1]
  data_fim <- if (!is.na(header_data_final)) {
    header_data_final
  } else {
    df$data[nrow(df)]
  }

  # ---- SALDO INICIAL (opening balance from first movement) ----
  saldo_inicial_calc <- df$saldo_atual[1] - df$valor_lancamento[1]

  saldo_inicial_row <- tibble::tibble(
    unidade_gestao = unidade_gestao,
    data = data_inicio,
    tipo = "SALDO_INICIAL",
    codigo_documento = NA_character_,
    valor_lancamento = 0,
    dc1 = NA_character_,
    saldo_atual = saldo_inicial_calc,
    dc2 = df$dc2[1],
    saldo_inicial_fim = saldo_inicial_calc
  )

  # ---- SALDO FINAL (closing balance from last movement saldo_atual) ----
  saldo_final_val <- df$saldo_atual[nrow(df)]

  saldo_final_row <- tibble::tibble(
    unidade_gestao = unidade_gestao,
    data = data_fim,
    tipo = "SALDO_FINAL",
    codigo_documento = NA_character_,
    valor_lancamento = 0,
    dc1 = NA_character_,
    saldo_atual = saldo_final_val,
    dc2 = df$dc2[nrow(df)],
    saldo_inicial_fim = saldo_final_val
  )

  dplyr::bind_rows(saldo_inicial_row, df, saldo_final_row)
}


list_pdf <- list.files(
  path = path_folder_source,
  pattern = "\\.pdf$",
  full.names = TRUE,
  recursive = TRUE
) |>
  str_subset(
    "CENTRAL USD|EXTRACTO DA CONTA FOREX EUR|EXTRACTO DA CONTA FOREX USD",
    negate = TRUE
  )


df <- list_pdf |>
  set_names(basename) |>
  map(extract_sistafe_table) |>
  list_rbind(names_to = "source_file")


# ---- Create date-range text from min/max(df$data) ----
date_min <- suppressWarnings(min(df$data, na.rm = TRUE))
date_max <- suppressWarnings(max(df$data, na.rm = TRUE))

# if everything is NA, keep a safe suffix
date_range_txt <- if (is.finite(date_min) && is.finite(date_max)) {
  paste0(format(date_min, "%Y-%m-%d"), "_a_", format(date_max, "%Y-%m-%d"))
} else {
  "sem_datas"
}

# ---- Write outputs with the date-range in the filename ----
csv_out <- file.path(
  "Dataout",
  paste0("razao_contabilistico_acumulado_", date_range_txt, ".csv")
)
xlsx_out <- file.path(
  "Dataout",
  paste0("razao_contabilistico_acumulado_", date_range_txt, ".xlsx")
)

write_csv(df, csv_out, na = "")
openxlsx::write.xlsx(df, file = xlsx_out, na.string = "")
