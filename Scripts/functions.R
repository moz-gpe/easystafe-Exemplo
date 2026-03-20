#' Extrair metadados do ficheiro e-SISTAFE
#'
#' Esta função lê o nome de um ficheiro do e-SISTAFE e extrai metadados
#' estruturados, incluindo:
#'
#' * tipo de reporte (Funcionamento, Investimento Interno, Investimento Externo);
#' * data de referência (a primeira data YYYYMMDD presente no nome);
#' * data de extração (a segunda data YYYYMMDD, se existir);
#' * ano e mês de referência, com nome do mês em Português.
#'
#' A função procura padrões \code{\\d{8}} no nome do ficheiro.
#' O primeiro padrão encontrado é assumido como data de referência
#' e o segundo como data de extração.
#'
#' @param caminho Caminho ou nome do ficheiro a partir do qual extrair metadados.
#'
#' @return Um tibble com as seguintes colunas:
#' \describe{
#'   \item{file_name}{Nome do ficheiro.}
#'   \item{reporte_tipo}{Classificação do tipo de reporte.}
#'   \item{data_reporte}{Data de referência (classe \code{Date}).}
#'   \item{data_extraido}{Data de extração (classe \code{Date}).}
#'   \item{ano}{Ano extraído da data de referência.}
#'   \item{mes}{Nome do mês (Português) correspondente à data de referência.}
#' }
#'
#' @examples
#' \dontrun{
#' extrair_meta_extracto(
#'   "OrcamentoFuncionamento_20240101_20240115.xlsx"
#' )
#' }
#'
#' @export

# extrair_meta_extracto <- function(caminho) {
#   # ------------------------------------------------------------
#   # Extract file name
#   # ------------------------------------------------------------
#   fname <- base::basename(caminho)
#
#   # ------------------------------------------------------------
#   # Report type classification
#   # ------------------------------------------------------------
#   if (stringr::str_detect(fname, "InvestimentoCompExterna")) {
#     report_type <- "Investimento Externo"
#   } else if (stringr::str_detect(fname, "InvestimentoCompInterna")) {
#     report_type <- "Investimento Interno"
#   } else if (stringr::str_detect(fname, "OrcamentoFuncionamento")) {
#     report_type <- "Funcionamento"
#   } else {
#     report_type <- NA_character_
#   }
#
#   # ------------------------------------------------------------
#   # Extract dates (YYYYMMDD patterns)
#   # ------------------------------------------------------------
#   dates <- stringr::str_extract_all(fname, "\\d{8}")[[1]]
#
#   ref_date <- dplyr::coalesce(dates[1], NA_character_)
#   extract_date <- dplyr::coalesce(dates[2], NA_character_)
#
#   # Convert to real Date objects
#   ref_dt <- base::as.Date(ref_date, format = "%Y%m%d")
#   extract_dt <- base::as.Date(extract_date, format = "%Y%m%d")
#
#   # ------------------------------------------------------------
#   # Portuguese month names (ASCII-safe)
#   # ------------------------------------------------------------
#   meses_pt <- c(
#     "Janeiro",
#     "Fevereiro",
#     "Mar\u00E7o",
#     "Abril",
#     "Maio",
#     "Junho",
#     "Julho",
#     "Agosto",
#     "Setembro",
#     "Outubro",
#     "Novembro",
#     "Dezembro"
#   )
#
#   # ------------------------------------------------------------
#   # Extract year + month
#   # ------------------------------------------------------------
#   ano <- if (!base::is.na(ref_dt)) lubridate::year(ref_dt) else NA_integer_
#   mes <- if (!base::is.na(ref_dt)) {
#     meses_pt[lubridate::month(ref_dt)]
#   } else {
#     NA_character_
#   }
#
#   # ------------------------------------------------------------
#   # Return metadata tibble
#   # ------------------------------------------------------------
#   tibble::tibble(
#     file_name = fname,
#     reporte_tipo = report_type,
#     data_reporte = ref_dt,
#     data_extraido = extract_dt,
#     ano = ano,
#     mes = mes
#   )
# }


extrair_meta_extracto <- function(caminho) {

  extrair_um <- function(c) {
    # ------------------------------------------------------------
    # Extract file name
    # ------------------------------------------------------------
    fname <- base::basename(c)
    # ------------------------------------------------------------
    # Report type classification
    # ------------------------------------------------------------
    if (stringr::str_detect(fname, "InvestimentoCompExterna")) {
      report_type <- "Investimento Externo"
    } else if (stringr::str_detect(fname, "InvestimentoCompInterna")) {
      report_type <- "Investimento Interno"
    } else if (stringr::str_detect(fname, "OrcamentoFuncionamento")) {
      report_type <- "Funcionamento"
    } else {
      report_type <- NA_character_
    }
    # ------------------------------------------------------------
    # Extract dates (YYYYMMDD patterns)
    # ------------------------------------------------------------
    dates <- stringr::str_extract_all(fname, "\\d{8}")[[1]]
    ref_date    <- dplyr::coalesce(dates[1], NA_character_)
    extract_date <- dplyr::coalesce(dates[2], NA_character_)
    # Convert to real Date objects
    ref_dt     <- base::as.Date(ref_date,     format = "%Y%m%d")
    extract_dt <- base::as.Date(extract_date, format = "%Y%m%d")
    # ------------------------------------------------------------
    # Portuguese month names (ASCII-safe)
    # ------------------------------------------------------------
    meses_pt <- c(
      "Janeiro", "Fevereiro", "Mar\u00E7o", "Abril",
      "Maio", "Junho", "Julho", "Agosto",
      "Setembro", "Outubro", "Novembro", "Dezembro"
    )
    # ------------------------------------------------------------
    # Extract year + month
    # ------------------------------------------------------------
    ano <- if (!base::is.na(ref_dt)) lubridate::year(ref_dt)          else NA_integer_
    mes <- if (!base::is.na(ref_dt)) meses_pt[lubridate::month(ref_dt)] else NA_character_
    # ------------------------------------------------------------
    # Return metadata tibble
    # ------------------------------------------------------------
    tibble::tibble(
      file_name     = fname,
      reporte_tipo  = report_type,
      data_reporte  = ref_dt,
      data_extraido = extract_dt,
      ano           = ano,
      mes           = mes
    )
  }

  purrr::map(caminho, extrair_um) |> list_rbind()

}

processar_extracto_sistafe <- function(
    source_path,
    ugb_lookup,
    include_percent = TRUE,
    include_meta    = TRUE,
    quiet           = TRUE
) {

  # --- Mensagens internas ---
  msg <- function(...) {
    if (!quiet) message(...)
  }

  # --- 1. Carregar ficheiros ---
  msg("A carregar ficheiros...")

  df <- map(source_path, ~read_excel(.x, col_types = "text")) |>
    set_names(basename(source_path)) |>
    list_rbind(names_to = "file_name")

  msg(glue("Ficheiros carregados: {n_distinct(df$file_name)} | Linhas: {nrow(df)}"))

  # --- 2. Adicionar ou remover metadados ---
  if (include_meta) {
    msg("A extrair e adicionar metadados...")

    paths_meta <- extrair_meta_extracto(source_path)

    df <- df |>
      left_join(paths_meta, by = "file_name") |>
      relocate(names(paths_meta)[-1], .after = file_name)
  }

  # --- 3. Renomear colunas ---
  msg("A limpar nomes de colunas...")

  df_limpeza_1 <- clean_names(df)

  # --- 4. Remover colunas percent ---
  msg("A remover colunas percent...")

  df_limpeza_2 <- df_limpeza_1 |>
    select(!ends_with("percent"))

  # --- 5. Extrair código UGB ---
  msg("A extrair código UGB...")

  df_limpeza_3 <- df_limpeza_2 |>
    mutate(
      across(dotacao_inicial:liq_ad_fundos_via_directa_lafvd, as.numeric),
      ugb_id = substr(ugb, 1, 9)
    ) |>
    relocate(ugb_id, .after = ugb)

  # --- 6. Filtrar UGBs de educação ---
  msg("A filtrar UGBs de educação...")

  vec_ugb <- ugb_lookup |>
    clean_names() |>
    select(ugb_nome = ugb_3) |>
    distinct(ugb_nome) |>
    pull()

  df_limpeza_4 <- df_limpeza_3 |>
    mutate(mec_ugb_class = ifelse(ugb %in% vec_ugb, "Keep", "Remove")) |>
    filter(mec_ugb_class == "Keep") |>
    select(-mec_ugb_class)

  # --- 7. Remover linhas com CED e funcao/programa/FR em branco ---
  msg("A remover linhas com CED e campos-chave em branco...")

  df_limpeza_5 <- df_limpeza_4 |>
    filter(!is.na(ced) | (!is.na(funcao) & !is.na(programa) & !is.na(fr))) |>
    mutate(data_tipo = if_else(is.na(ced), "Metrica", "Valor")) |>
    relocate(data_tipo, .before = ced)

  # --- 8. Classificar grupos CED e remover grupo D ---
  msg("A classificar grupos CED e remover grupo D...")

  df_limpeza_6 <- df_limpeza_5 |>
    mutate(
      ced_group = case_when(
        !str_ends(ced, "00")                                                  ~ "A",
        str_ends(ced, "00") & !str_ends(ced, "000") & !str_ends(ced, "0000") ~ "B",
        str_ends(ced, "000") & !str_ends(ced, "0000")                        ~ "C",
        str_ends(ced, "0000")                                                 ~ "D",
        TRUE                                                                  ~ NA_character_
      )
    ) |>
    filter(is.na(ced_group) | ced_group != "D")

  # --- 9. Criar variáveis hierárquicas ---
  msg("A criar variáveis hierárquicas...")

  df_limpeza_7 <- df_limpeza_6 |>
    mutate(
      ced_b4   = str_sub(ced, 1, 4),
      ced_b3   = str_sub(ced, 1, 3),
      id_ced_b4 = str_c(ugb_id, funcao, programa, fr, ced_b4, sep = " | "),
      id_ced_b3 = str_c(ugb_id, funcao, programa, fr, ced_b3, sep = " | ")
    ) |>
    unite(ugb_funcao_prog_fr, ugb_id, funcao, programa, fr, sep = " | ", remove = FALSE, na.rm = FALSE) |>
    relocate(c(ced_b4, ced_b3), .after = ced) |>
    relocate(ced_group, .before = ced) |>
    relocate(data_tipo, .after = ced_b3) |>
    relocate(ugb_funcao_prog_fr, .before = everything())

  # --- 10. Definir colunas numéricas ---
  num_cols <- df_limpeza_7 |>
    select(dotacao_inicial:liq_ad_fundos_via_directa_lafvd) |>
    names()

  # --- 11. Subtração hierárquica: Passo 1 (A -> B dentro de ced_b4) ---
  msg("A executar subtração hierárquica — Passo 1 (A → B)...")

  df_step1 <- df_limpeza_7 |>
    filter(data_tipo == "Valor") |>
    group_by(ugb_funcao_prog_fr, ced_b4) |>
    mutate(
      across(
        all_of(num_cols),
        ~ if_else(ced_group == "B", .x - sum(.x[ced_group == "A"], na.rm = TRUE), .x)
      )
    ) |>
    ungroup()

  # --- 12. Subtração hierárquica: Passo 2 (B ajustado -> C dentro de ced_b3) ---
  msg("A executar subtração hierárquica — Passo 2 (B → C)...")

  df_step2 <- df_step1 |>
    group_by(ugb_funcao_prog_fr, ced_b3) |>
    mutate(
      across(
        all_of(num_cols),
        ~ if_else(ced_group == "C", .x - sum(.x[ced_group == "B"], na.rm = TRUE), .x)
      )
    ) |>
    ungroup()

  # --- 13. Subtração hierárquica: Passo 3 (A directo -> C dentro de ced_b3) ---
  msg("A executar subtração hierárquica — Passo 3 (A directo → C)...")

  df_limpeza_9 <- df_step2 |>
    group_by(ugb_funcao_prog_fr, ced_b3) |>
    mutate(
      across(
        all_of(num_cols),
        ~ if_else(ced_group == "C", .x - sum(.x[ced_group == "A"], na.rm = TRUE), .x)
      )
    ) |>
    ungroup()

  # --- 14. Seleccionar colunas finais e restaurar estrutura original ---
  msg("A finalizar estrutura do dataset...")

  df_limpeza_final <- df_limpeza_9 |>
    select(any_of(names(df_limpeza_1))) |>
    mutate(
      dc_da_percent   = NA_real_,
      afdp_da_percent = NA_real_,
      laf_af_percent  = NA_real_
    ) |>
    select(all_of(names(df_limpeza_1)))

  # --- 15. Incluir ou excluir colunas percent ---
  if (!include_percent) {
    df_limpeza_final <- df_limpeza_final |>
      select(!ends_with("percent"))
  }

  # --- 16. Remover file_name se metadados não incluídos ---
  if (!include_meta) {
    df_limpeza_final <- df_limpeza_final |>
      select(-file_name)
  }

  msg("Concluído.")

  # --- Resumo final ---
  n_files <- n_distinct(df$file_name)
  message(glue("Processamento concluído: {n_files} ficheiro(s) processado(s) com sucesso."))

  return(df_limpeza_final)

}



gravar_extracto_sistafe <- function(
    df,
    output_folder = "Dataout",
    quiet         = TRUE
) {

  # --- Mensagens internas ---
  msg <- function(...) {
    if (!quiet) message(...)
  }

  # --- Verificar que colunas de metadados existem ---
  required_cols <- c("reporte_tipo", "ano", "mes")
  missing_cols  <- setdiff(required_cols, names(df))

  if (length(missing_cols) > 0) {
    stop(glue("Colunas de metadados em falta: {paste(missing_cols, collapse = ', ')}. ",
              "Certifique-se de que include_meta = TRUE foi usado ao processar o ficheiro."))
  }

  # --- Extrair valores únicos dos metadados ---
  reporte_tipo <- df |> distinct(reporte_tipo) |> pull() |> paste(collapse = "-")
  ano          <- df |> distinct(ano)          |> pull() |> paste(collapse = "-")
  mes          <- df |> distinct(mes)          |> pull() |> paste(collapse = "-")
  today        <- format(Sys.Date(), "%Y%m%d")

  # --- Construir nome do ficheiro ---
  file_name <- glue("{reporte_tipo}_{ano}_{mes}_{today}.xlsx")
  file_path <- file.path(output_folder, file_name)

  # --- Criar pasta se não existir ---
  if (!dir.exists(output_folder)) {
    msg(glue("Pasta '{output_folder}' não encontrada — a criar..."))
    dir.create(output_folder, recursive = TRUE)
  }

  # --- Guardar ficheiro ---
  msg(glue("A guardar ficheiro: {file_path}"))
  write_xlsx(df, file_path)
  msg("Concluído.")

  # --- Retornar caminho invisível para uso posterior se necessário ---
  invisible(file_path)

}
