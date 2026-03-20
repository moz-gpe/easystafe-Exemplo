#' Extrair metadados de ficheiros de extracto e-SISTAFE
#'
#' Extrai metadados relevantes a partir dos nomes de ficheiros de exportação
#' do e-SISTAFE, incluindo o tipo de relatório, datas de referência e de
#' extracção, ano e mês em português. Suporta um ou múltiplos ficheiros.
#'
#' @param caminho Um vector de caracteres com um ou mais caminhos completos ou
#'   relativos para ficheiros de exportação e-SISTAFE. Os nomes dos ficheiros
#'   devem seguir a convenção de nomenclatura padrão do e-SISTAFE, contendo
#'   padrões de data no formato \code{YYYYMMDD}.
#'
#' @return Um tibble com uma linha por ficheiro e as seguintes colunas:
#' \describe{
#'   \item{file_name}{Nome do ficheiro sem o caminho completo.}
#'   \item{reporte_tipo}{Tipo de relatório classificado a partir do nome do
#'     ficheiro. Um de \code{"Funcionamento"}, \code{"Investimento Externo"},
#'     \code{"Investimento Interno"}, ou \code{NA} se não reconhecido.}
#'   \item{data_reporte}{Data de referência do relatório como objecto
#'     \code{Date}, extraída do primeiro padrão \code{YYYYMMDD} no nome
#'     do ficheiro.}
#'   \item{data_extraido}{Data de extracção do ficheiro como objecto
#'     \code{Date}, extraída do segundo padrão \code{YYYYMMDD} no nome
#'     do ficheiro.}
#'   \item{ano}{Ano da data de referência como inteiro.}
#'   \item{mes}{Mês da data de referência em português (ex. \code{"Janeiro"}).}
#' }
#'
#' @details
#' A classificação do tipo de relatório é feita por detecção de padrões no
#' nome do ficheiro:
#' \itemize{
#'   \item \code{"InvestimentoCompExterna"} → \code{"Investimento Externo"}
#'   \item \code{"InvestimentoCompInterna"} → \code{"Investimento Interno"}
#'   \item \code{"OrcamentoFuncionamento"}  → \code{"Funcionamento"}
#' }
#' Se nenhum padrão for reconhecido, \code{reporte_tipo} é \code{NA}.
#'
#' As datas são extraídas pelo padrão regex \code{\\d{8}} — espera-se que o
#' primeiro match corresponda à data de referência do relatório e o segundo
#' à data de extracção.
#'
#' @examples
#' \dontrun{
#' # Ficheiro único
#' extrair_meta_extracto("Data/DemonstrativoConsolidadoOrcamentoFuncionamento_20251231_20260205.xlsx")
#'
#' # Múltiplos ficheiros
#' path_files <- list.files("Data/", pattern = "\\.xlsx$", full.names = TRUE)
#' extrair_meta_extracto(path_files)
#' }
#'
#' @importFrom purrr map
#' @importFrom dplyr coalesce
#' @importFrom stringr str_detect str_extract_all
#' @importFrom tibble tibble
#' @importFrom lubridate year month
#'
#' @export

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




#' Processar extractos de exportação e-SISTAFE
#'
#' Carrega, limpa e processa um ou mais ficheiros de exportação do e-SISTAFE
#' no formato Excel, aplicando uma sequência de transformações que inclui
#' renomeação de colunas, filtragem de UGBs de educação, classificação e
#' subtracção hierárquica de códigos CED, e restauração da estrutura original
#' de colunas. Devolve um dataframe final desduplicado e pronto para análise.
#'
#' @param source_path Um vector de caracteres com um ou mais caminhos para
#'   ficheiros de exportação e-SISTAFE no formato \code{.xlsx}.
#' @param ugb_lookup Um dataframe com a tabela de referência de UGBs de
#'   educação, carregado a partir do ficheiro Excel de códigos UGB (folha
#'   \code{"UGBS"}). Deve conter uma coluna \code{ugb_3} com os nomes
#'   completos dos UGBs válidos.
#' @param include_percent Lógico. Se \code{TRUE} (padrão), as colunas
#'   \code{percent} são incluídas no output (preenchidas com \code{NA}).
#'   Se \code{FALSE}, essas colunas são removidas do resultado final.
#' @param include_meta Lógico. Se \code{TRUE} (padrão), os metadados
#'   extraídos do nome do ficheiro (tipo de relatório, ano, mês, datas) são
#'   adicionados ao dataframe imediatamente após a coluna \code{file_name}.
#'   Se \code{FALSE}, os metadados não são adicionados e a coluna
#'   \code{file_name} é também removida do resultado final.
#' @param quiet Lógico. Se \code{TRUE} (padrão), as mensagens de progresso
#'   são suprimidas. Se \code{FALSE}, é emitida uma mensagem por cada etapa
#'   do processamento. Independentemente deste parâmetro, é sempre emitida
#'   uma mensagem final com o número de ficheiros processados.
#'
#' @return Um tibble com uma linha por entrada CED deduplificada, contendo
#'   as colunas originais do extracto e-SISTAFE após limpeza e subtracção
#'   hierárquica. As colunas de percentagem são sempre incluídas na estrutura
#'   original (preenchidas com \code{NA}) salvo se \code{include_percent = FALSE}.
#'
#' @details
#' O processamento segue as seguintes etapas principais:
#' \enumerate{
#'   \item Carregamento e combinação de todos os ficheiros em \code{source_path}.
#'   \item Adição opcional de metadados via \code{extrair_meta_extracto()}.
#'   \item Limpeza de nomes de colunas com \code{janitor::clean_names()}.
#'   \item Remoção de colunas \code{percent}.
#'   \item Conversão de colunas numéricas e extracção do código \code{ugb_id}.
#'   \item Filtragem de UGBs válidos de educação a partir de \code{ugb_lookup}.
#'   \item Remoção de linhas com CED e campos-chave em branco.
#'   \item Classificação de grupos CED (A, B, C, D) e remoção do grupo D.
#'   \item Criação de variáveis hierárquicas auxiliares.
#'   \item Subtracção hierárquica em três passos para eliminar dupla contagem:
#'     \itemize{
#'       \item Passo 1: Subtrair grupo A do grupo B (dentro de \code{ced_b4}).
#'       \item Passo 2: Subtrair grupo B ajustado do grupo C (dentro de \code{ced_b3}).
#'       \item Passo 3: Subtrair grupo A directamente do grupo C (dentro de \code{ced_b3}).
#'     }
#'   \item Restauração da estrutura original de colunas.
#' }
#'
#' @examples
#' \dontrun{
#' ugb_raw    <- readxl::read_excel("Data/ugb/Codigos de UGBs.xlsx", sheet = "UGBS")
#' path_files <- list.files("Data/", pattern = "\\.xlsx$", full.names = TRUE)
#'
#' # Padrão — com metadados e colunas percent
#' df <- processar_extracto_sistafe(
#'   source_path = path_files,
#'   ugb_lookup  = ugb_raw
#' )
#'
#' # Sem metadados, sem colunas percent
#' df <- processar_extracto_sistafe(
#'   source_path     = path_files,
#'   ugb_lookup      = ugb_raw,
#'   include_percent = FALSE,
#'   include_meta    = FALSE
#' )
#'
#' # Com mensagens de progresso
#' df <- processar_extracto_sistafe(
#'   source_path = path_files,
#'   ugb_lookup  = ugb_raw,
#'   quiet       = FALSE
#' )
#' }
#'
#' @importFrom purrr map
#' @importFrom readxl read_excel
#' @importFrom dplyr n_distinct mutate across relocate filter select left_join
#'   if_else case_when group_by ungroup bind_rows distinct
#' @importFrom tidyr unite
#' @importFrom stringr str_ends str_sub str_c
#' @importFrom janitor clean_names
#' @importFrom glue glue
#' @importFrom tibble tibble
#'
#' @export

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

  df <- purrr::map(source_path, ~readxl::read_excel(.x, col_types = "text")) |>
    purrr::set_names(base::basename(source_path)) |>
    purrr::list_rbind(names_to = "file_name")

  msg(glue::glue("Ficheiros carregados: {dplyr::n_distinct(df$file_name)} | Linhas: {nrow(df)}"))

  # --- 2. Adicionar ou remover metadados ---
  if (include_meta) {
    msg("A extrair e adicionar metadados...")

    paths_meta <- extrair_meta_extracto(source_path)

    df <- df |>
      dplyr::left_join(paths_meta, by = "file_name") |>
      dplyr::relocate(names(paths_meta)[-1], .after = file_name)
  }

  # --- 3. Renomear colunas ---
  msg("A limpar nomes de colunas...")

  df_limpeza_1 <- janitor::clean_names(df)

  # --- 4. Remover colunas percent ---
  msg("A remover colunas percent...")

  df_limpeza_2 <- df_limpeza_1 |>
    dplyr::select(!dplyr::ends_with("percent"))

  # --- 5. Extrair código UGB ---
  msg("A extrair código UGB...")

  df_limpeza_3 <- df_limpeza_2 |>
    dplyr::mutate(
      dplyr::across(dotacao_inicial:liq_ad_fundos_via_directa_lafvd, as.numeric),
      ugb_id = base::substr(ugb, 1, 9)
    ) |>
    dplyr::relocate(ugb_id, .after = ugb)

  # --- 6. Filtrar UGBs de educação ---
  msg("A filtrar UGBs de educação...")

  vec_ugb <- ugb_lookup |>
    janitor::clean_names() |>
    dplyr::select(ugb_nome = ugb_3) |>
    dplyr::distinct(ugb_nome) |>
    dplyr::pull()

  df_limpeza_4 <- df_limpeza_3 |>
    dplyr::mutate(mec_ugb_class = base::ifelse(ugb %in% vec_ugb, "Keep", "Remove")) |>
    dplyr::filter(mec_ugb_class == "Keep") |>
    dplyr::select(-mec_ugb_class)

  # --- 7. Remover linhas com CED e funcao/programa/FR em branco ---
  msg("A remover linhas com CED e campos-chave em branco...")

  df_limpeza_5 <- df_limpeza_4 |>
    dplyr::filter(!base::is.na(ced) | (!base::is.na(funcao) & !base::is.na(programa) & !base::is.na(fr))) |>
    dplyr::mutate(data_tipo = dplyr::if_else(base::is.na(ced), "Metrica", "Valor")) |>
    dplyr::relocate(data_tipo, .before = ced)

  # --- 8. Classificar grupos CED e remover grupo D ---
  msg("A classificar grupos CED e remover grupo D...")

  df_limpeza_6 <- df_limpeza_5 |>
    dplyr::mutate(
      ced_group = dplyr::case_when(
        !stringr::str_ends(ced, "00")                                                  ~ "A",
        stringr::str_ends(ced, "00") & !stringr::str_ends(ced, "000") & !stringr::str_ends(ced, "0000") ~ "B",
        stringr::str_ends(ced, "000") & !stringr::str_ends(ced, "0000")               ~ "C",
        stringr::str_ends(ced, "0000")                                                 ~ "D",
        TRUE                                                                            ~ NA_character_
      )
    ) |>
    dplyr::filter(base::is.na(ced_group) | ced_group != "D")

  # --- 9. Criar variáveis hierárquicas ---
  msg("A criar variáveis hierárquicas...")

  df_limpeza_7 <- df_limpeza_6 |>
    dplyr::mutate(
      ced_b4    = stringr::str_sub(ced, 1, 4),
      ced_b3    = stringr::str_sub(ced, 1, 3),
      id_ced_b4 = stringr::str_c(ugb_id, funcao, programa, fr, ced_b4, sep = " | "),
      id_ced_b3 = stringr::str_c(ugb_id, funcao, programa, fr, ced_b3, sep = " | ")
    ) |>
    tidyr::unite(ugb_funcao_prog_fr, ugb_id, funcao, programa, fr, sep = " | ", remove = FALSE, na.rm = FALSE) |>
    dplyr::relocate(c(ced_b4, ced_b3), .after = ced) |>
    dplyr::relocate(ced_group, .before = ced) |>
    dplyr::relocate(data_tipo, .after = ced_b3) |>
    dplyr::relocate(ugb_funcao_prog_fr, .before = dplyr::everything())

  # --- 10. Definir colunas numéricas ---
  num_cols <- df_limpeza_7 |>
    dplyr::select(dotacao_inicial:liq_ad_fundos_via_directa_lafvd) |>
    base::names()

  # --- 11. Subtração hierárquica: Passo 1 (A -> B dentro de ced_b4) ---
  msg("A executar subtração hierárquica — Passo 1 (A → B)...")

  df_step1 <- df_limpeza_7 |>
    dplyr::filter(data_tipo == "Valor") |>
    dplyr::group_by(ugb_funcao_prog_fr, ced_b4) |>
    dplyr::mutate(
      dplyr::across(
        dplyr::all_of(num_cols),
        ~ dplyr::if_else(ced_group == "B", .x - base::sum(.x[ced_group == "A"], na.rm = TRUE), .x)
      )
    ) |>
    dplyr::ungroup()

  # --- 12. Subtração hierárquica: Passo 2 (B ajustado -> C dentro de ced_b3) ---
  msg("A executar subtração hierárquica — Passo 2 (B → C)...")

  df_step2 <- df_step1 |>
    dplyr::group_by(ugb_funcao_prog_fr, ced_b3) |>
    dplyr::mutate(
      dplyr::across(
        dplyr::all_of(num_cols),
        ~ dplyr::if_else(ced_group == "C", .x - base::sum(.x[ced_group == "B"], na.rm = TRUE), .x)
      )
    ) |>
    dplyr::ungroup()

  # --- 13. Subtração hierárquica: Passo 3 (A directo -> C dentro de ced_b3) ---
  msg("A executar subtração hierárquica — Passo 3 (A directo → C)...")

  df_limpeza_9 <- df_step2 |>
    dplyr::group_by(ugb_funcao_prog_fr, ced_b3) |>
    dplyr::mutate(
      dplyr::across(
        dplyr::all_of(num_cols),
        ~ dplyr::if_else(ced_group == "C", .x - base::sum(.x[ced_group == "A"], na.rm = TRUE), .x)
      )
    ) |>
    dplyr::ungroup()

  # --- 14. Seleccionar colunas finais e restaurar estrutura original ---
  msg("A finalizar estrutura do dataset...")

  df_limpeza_final <- df_limpeza_9 |>
    dplyr::select(dplyr::any_of(base::names(df_limpeza_1))) |>
    dplyr::mutate(
      dc_da_percent   = NA_real_,
      afdp_da_percent = NA_real_,
      laf_af_percent  = NA_real_
    ) |>
    dplyr::select(dplyr::all_of(base::names(df_limpeza_1)))

  # --- 15. Incluir ou excluir colunas percent ---
  if (!include_percent) {
    df_limpeza_final <- df_limpeza_final |>
      dplyr::select(!dplyr::ends_with("percent"))
  }

  # --- 16. Remover file_name se metadados não incluídos ---
  if (!include_meta) {
    df_limpeza_final <- df_limpeza_final |>
      dplyr::select(-file_name)
  }

  msg("Concluído.")

  # --- Resumo final ---
  n_files <- dplyr::n_distinct(df$file_name)
  message(glue::glue("Processamento concluído: {n_files} ficheiro(s) processado(s) com sucesso."))

  return(df_limpeza_final)

}



#' Gravar extracto processado do e-SISTAFE em Excel
#'
#' Grava um dataframe processado do e-SISTAFE num ficheiro Excel, construindo
#' automaticamente o nome do ficheiro a partir dos metadados do relatório
#' (tipo, ano e mês) e da data actual. Cria a pasta de destino se não existir.
#'
#' @param df Um dataframe processado por \code{processar_extracto_sistafe()}
#'   com \code{include_meta = TRUE}. Deve conter as colunas \code{reporte_tipo},
#'   \code{ano} e \code{mes}.
#' @param output_folder Caractere. Caminho para a pasta de destino onde o
#'   ficheiro Excel será gravado. Por padrão \code{"Dataout"}. A pasta é
#'   criada automaticamente se não existir.
#' @param quiet Lógico. Se \code{TRUE} (padrão), as mensagens de progresso
#'   são suprimidas. Se \code{FALSE}, são emitidas mensagens sobre a criação
#'   da pasta e o caminho do ficheiro gravado.
#'
#' @return O caminho completo do ficheiro gravado, retornado de forma invisível.
#'   Pode ser capturado com \code{path <- gravar_extracto_sistafe(df)} para
#'   uso posterior se necessário.
#'
#' @details
#' O nome do ficheiro é construído automaticamente no formato:
#' \code{<reporte_tipo>_<ano>_<mes>_<YYYYMMDD>.xlsx}
#'
#' Por exemplo: \code{Funcionamento_2025_Dezembro_20260320.xlsx}
#'
#' Se o dataframe contiver múltiplos valores para \code{reporte_tipo},
#' \code{ano} ou \code{mes} (por exemplo, quando se combinam vários meses),
#' os valores são concatenados com \code{"-"} no nome do ficheiro.
#'
#' Esta função requer que \code{processar_extracto_sistafe()} tenha sido
#' chamado com \code{include_meta = TRUE}. Se as colunas de metadados
#' estiverem em falta, a função para com uma mensagem de erro informativa.
#'
#' @examples
#' \dontrun{
#' # Gravar com pasta padrão
#' gravar_extracto_sistafe(df)
#'
#' # Gravar numa pasta personalizada
#' gravar_extracto_sistafe(df, output_folder = "Data/processed")
#'
#' # Gravar com mensagens de progresso
#' gravar_extracto_sistafe(df, quiet = FALSE)
#'
#' # Capturar o caminho do ficheiro gravado
#' path <- gravar_extracto_sistafe(df, quiet = FALSE)
#' }
#'
#' @importFrom dplyr distinct pull
#' @importFrom glue glue
#' @importFrom writexl write_xlsx
#'
#' @export
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
  missing_cols  <- base::setdiff(required_cols, base::names(df))

  if (base::length(missing_cols) > 0) {
    stop(glue::glue(
      "Colunas de metadados em falta: {paste(missing_cols, collapse = ', ')}. ",
      "Certifique-se de que include_meta = TRUE foi usado ao processar o ficheiro."
    ))
  }

  # --- Remover trailing slash se presente ---
  output_folder <- base::gsub("/$", "", output_folder)

  # --- Extrair valores únicos dos metadados ---
  reporte_tipo <- df |> dplyr::distinct(reporte_tipo) |> dplyr::pull() |> base::paste(collapse = "-")
  ano          <- df |> dplyr::distinct(ano)          |> dplyr::pull() |> base::paste(collapse = "-")
  mes          <- df |> dplyr::distinct(mes)          |> dplyr::pull() |> base::paste(collapse = "-")
  today        <- base::format(base::Sys.Date(), "%Y%m%d")

  # --- Construir nome do ficheiro ---
  file_name <- glue::glue("{reporte_tipo}_{ano}_{mes}_{today}.xlsx")
  file_path <- base::file.path(output_folder, file_name)

  # --- Criar pasta se não existir ---
  if (!base::dir.exists(output_folder)) {
    msg(glue::glue("Pasta '{output_folder}' não encontrada — a criar..."))
    base::dir.create(output_folder, recursive = TRUE)
  }

  # --- Guardar ficheiro ---
  msg(glue::glue("A guardar ficheiro: {file_path}"))
  writexl::write_xlsx(df, file_path)
  msg("Concluído.")

  # --- Retornar caminho invisível para uso posterior se necessário ---
  base::invisible(file_path)

}
