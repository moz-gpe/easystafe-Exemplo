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

extrair_meta_extracto <- function(caminho) {
  # ------------------------------------------------------------
  # Extract file name
  # ------------------------------------------------------------
  fname <- base::basename(caminho)

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

  ref_date <- dplyr::coalesce(dates[1], NA_character_)
  extract_date <- dplyr::coalesce(dates[2], NA_character_)

  # Convert to real Date objects
  ref_dt <- base::as.Date(ref_date, format = "%Y%m%d")
  extract_dt <- base::as.Date(extract_date, format = "%Y%m%d")

  # ------------------------------------------------------------
  # Portuguese month names (ASCII-safe)
  # ------------------------------------------------------------
  meses_pt <- c(
    "Janeiro",
    "Fevereiro",
    "Mar\u00E7o",
    "Abril",
    "Maio",
    "Junho",
    "Julho",
    "Agosto",
    "Setembro",
    "Outubro",
    "Novembro",
    "Dezembro"
  )

  # ------------------------------------------------------------
  # Extract year + month
  # ------------------------------------------------------------
  ano <- if (!base::is.na(ref_dt)) lubridate::year(ref_dt) else NA_integer_
  mes <- if (!base::is.na(ref_dt)) {
    meses_pt[lubridate::month(ref_dt)]
  } else {
    NA_character_
  }

  # ------------------------------------------------------------
  # Return metadata tibble
  # ------------------------------------------------------------
  tibble::tibble(
    file_name = fname,
    reporte_tipo = report_type,
    data_reporte = ref_dt,
    data_extraido = extract_dt,
    ano = ano,
    mes = mes
  )
}
