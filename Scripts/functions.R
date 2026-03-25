
#' Processar extractos de exportacao e-SISTAFE
#'
#' Carrega, limpa e processa um ou mais ficheiros de exportacao do e-SISTAFE
#' no formato Excel, aplicando uma sequencia de transformacoes que inclui
#' renomeacao de colunas, filtragem de UGBs de educacao, classificacao e
#' subtraccao hierarquica de codigos CED, e restauracao da estrutura original
#' de colunas. Devolve um dataframe final desduplicado e pronto para analise.
#'
#' @param source_path Um vector de caracteres com um ou mais caminhos para
#'   ficheiros de exportacao e-SISTAFE no formato \code{.xlsx}.
#' @param ugb_lookup Um dataframe com a tabela de referencia de UGBs de
#'   educacao, carregado a partir do ficheiro Excel de codigos UGB (folha
#'   \code{"UGBS"}). Deve conter uma coluna \code{ugb_3} com os nomes
#'   completos dos UGBs validos.
#' @param include_percent Logico. Se \code{TRUE} (padrao), as colunas
#'   \code{percent} sao incluidas no output (preenchidas com \code{NA}).
#'   Se \code{FALSE}, essas colunas sao removidas do resultado final.
#' @param include_meta Logico. Se \code{TRUE} (padrao), os metadados
#'   extraidos do nome do ficheiro (tipo de relatorio, ano, mes, datas) sao
#'   adicionados ao dataframe imediatamente apos a coluna \code{file_name}.
#'   Se \code{FALSE}, os metadados nao sao adicionados e a coluna
#'   \code{file_name} e tambem removida do resultado final.
#' @param include_metrica Logico. Se \code{FALSE} (padrao), as linhas do tipo
#'   \code{"Metrica"} sao excluidas do output final, mantendo apenas as linhas
#'   \code{"Valor"} apos subtraccao hierarquica. Se \code{TRUE}, as linhas
#'   \code{"Metrica"} sao reincluidas no output final apos o processamento,
#'   util para comparacoes e validacao. A coluna \code{data_tipo} e sempre
#'   incluida no output, independentemente deste parametro.
#' @param quiet Logico. Se \code{TRUE} (padrao), as mensagens de progresso
#'   sao suprimidas. Se \code{FALSE}, e emitida uma mensagem por cada etapa
#'   do processamento. Independentemente deste parametro, e sempre emitida
#'   uma mensagem final com o numero de ficheiros processados.
#'
#' @return Um tibble com uma linha por entrada CED deduplificada, contendo
#'   as colunas originais do extracto e-SISTAFE apos limpeza e subtraccao
#'   hierarquica. A coluna \code{data_tipo} esta sempre presente e posicionada
#'   imediatamente antes de \code{ugb}. As colunas de percentagem sao sempre
#'   incluidas na estrutura original (preenchidas com \code{NA}) salvo se
#'   \code{include_percent = FALSE}.
#'
#' @details
#' O processamento segue as seguintes etapas principais:
#' \enumerate{
#'   \item Carregamento e combinacao de todos os ficheiros em \code{source_path}.
#'   \item Adicao opcional de metadados via \code{extrair_meta_extracto()}.
#'   \item Limpeza de nomes de colunas com \code{janitor::clean_names()}.
#'   \item Remocao de colunas \code{percent}.
#'   \item Conversao de colunas numericas e extraccao do codigo \code{ugb_id}.
#'   \item Filtragem de UGBs validos de educacao a partir de \code{ugb_lookup}.
#'   \item Remocao de linhas com CED e campos-chave em branco.
#'   \item Classificacao de grupos CED (A, B, C, D) e remocao do grupo D.
#'   \item Criacao de variaveis hierarquicas auxiliares.
#'   \item Separacao de linhas \code{"Metrica"} e \code{"Valor"} antes da
#'     subtraccao hierarquica.
#'   \item Subtraccao hierarquica em tres passos para eliminar dupla contagem
#'     (aplicada apenas a linhas \code{"Valor"}):
#'     \itemize{
#'       \item Passo 1: Subtrair grupo A do grupo B (dentro de \code{ced_b4}).
#'       \item Passo 2: Subtrair grupo B ajustado do grupo C (dentro de \code{ced_b3}).
#'       \item Passo 3: Subtrair grupo A directamente do grupo C (dentro de \code{ced_b3}).
#'     }
#'   \item Reinclusao opcional das linhas \code{"Metrica"} via \code{include_metrica}.
#'   \item Seleccao das colunas finais a partir de um vector explicito,
#'     garantindo que \code{data_tipo} e sempre incluido antes de \code{ugb}.
#' }
#'
#' @examples
#' \dontrun{
#' ugb_raw    <- readxl::read_excel("Data/ugb/Codigos de UGBs.xlsx", sheet = "UGBS")
#' path_files <- list.files("Data/", pattern = "\\.xlsx$", full.names = TRUE)
#'
#' # Padrao -- com metadados e colunas percent, sem linhas Metrica
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
#' # Com linhas Metrica incluidas para comparacao
#' df <- processar_extracto_sistafe(
#'   source_path      = path_files,
#'   ugb_lookup       = ugb_raw,
#'   include_metrica  = TRUE
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

processar_extracto_esistafe <- function(
    source_path,
    ugb_lookup,
    include_percent  = TRUE,
    include_meta     = TRUE,
    include_metrica  = FALSE,
    quiet            = TRUE
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
  msg("A extrair c\u00f3digo UGB...")

  df_limpeza_3 <- df_limpeza_2 |>
    dplyr::mutate(
      dplyr::across(dotacao_inicial:liq_ad_fundos_via_directa_lafvd, as.numeric),
      ugb_id = base::substr(ugb, 1, 9)
    ) |>
    dplyr::relocate(ugb_id, .after = ugb)

  # --- 6. Filtrar UGBs de educação ---
  msg("A filtrar UGBs de educa\u00e7\u00e3o...")

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
        !stringr::str_ends(ced, "00")                                                                    ~ "A",
        stringr::str_ends(ced, "00") & !stringr::str_ends(ced, "000") & !stringr::str_ends(ced, "0000") ~ "B",
        stringr::str_ends(ced, "000") & !stringr::str_ends(ced, "0000")                                 ~ "C",
        stringr::str_ends(ced, "0000")                                                                   ~ "D",
        TRUE                                                                                              ~ NA_character_
      )
    ) |>
    dplyr::filter(base::is.na(ced_group) | ced_group != "D")

  # --- 9. Criar variáveis hierárquicas ---
  msg("A criar vari\u00e1veis hier\u00e1rquicas...")

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

  # --- 10b. Separar linhas Metrica antes da subtraccao hierarquica ---
  msg("A separar linhas Metrica e Valor...")

  df_metrica <- df_limpeza_7 |>
    dplyr::filter(data_tipo == "Metrica")

  # --- 11. Subtração hierárquica: Passo 1 (A -> B dentro de ced_b4) ---
  # Nota: apenas linhas "Valor" entram na subtraccao -- comportamento agora explicito
  msg("A executar subtra\u00e7\u00e3o hier\u00e1rquica \u2014 Passo 1 (A \u2192 B)...")

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
  msg("A executar subtra\u00e7\u00e3o hier\u00e1rquica \u2014 Passo 2 (B \u2192 C)...")

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
  msg("A executar subtra\u00e7\u00e3o hier\u00e1rquica \u2014 Passo 3 (A directo \u2192 C)...")

  df_limpeza_9 <- df_step2 |>
    dplyr::group_by(ugb_funcao_prog_fr, ced_b3) |>
    dplyr::mutate(
      dplyr::across(
        dplyr::all_of(num_cols),
        ~ dplyr::if_else(ced_group == "C", .x - base::sum(.x[ced_group == "A"], na.rm = TRUE), .x)
      )
    ) |>
    dplyr::ungroup()

  # --- 13b. Reincluir linhas Metrica se solicitado ---
  if (include_metrica) {
    msg("A reincluir linhas Metrica...")
    df_limpeza_9 <- dplyr::bind_rows(df_limpeza_9, df_metrica)
  }

  # --- 14. Seleccionar colunas finais a partir de vector explicito ---
  # data_tipo e sempre incluido, posicionado antes de ugb.
  # percent e file_name sao incluidos ou excluidos conforme os argumentos.
  msg("A finalizar estrutura do dataset...")

  final_cols <- c(
    # metadados de ficheiro (removidos se include_meta = FALSE)
    "file_name", "reporte_tipo", "data_reporte", "data_extraido",
    "ano", "mes",
    # classificacao da linha -- sempre presente
    "data_tipo",
    "ugb_id",
    # identificadores orcamentais
    "ugb", "funcao", "programa", "fr", "ced",
    # colunas numericas
    "dotacao_inicial",
    "dotacao_revista",
    "dotacao_actualizada_da",
    "dotacao_disponivel",
    "dotacao_cabimentada_dc",
    "ad_fundos_concedidos_af",
    "despesa_paga_via_directa_dp",
    "ad_fundos_desp_paga_vd_afdp",
    "ad_fundos_liquidados_laf",
    "despesa_liquidada_via_directa_lvd",
    "liq_ad_fundos_via_directa_lafvd",
    "dc_da_percent",
    "afdp_da_percent",
    "laf_af_percent"
  )

  df_limpeza_final <- df_limpeza_9 |>
    dplyr::mutate(
      dc_da_percent   = NA_real_,
      afdp_da_percent = NA_real_,
      laf_af_percent  = NA_real_
    ) |>
    dplyr::select(dplyr::any_of(final_cols)) |>
    dplyr::relocate(data_tipo, .after = mes) |>
    dplyr::arrange(ugb)

  # --- 15. Excluir colunas percent se solicitado ---
  if (!include_percent) {
    df_limpeza_final <- df_limpeza_final |>
      dplyr::select(!dplyr::ends_with("percent"))
  }

  # --- 16. Remover file_name e metadados se include_meta = FALSE ---
  if (!include_meta) {
    df_limpeza_final <- df_limpeza_final |>
      dplyr::select(-dplyr::any_of(c("file_name", "reporte_tipo", "data_reporte", "data_extraido", "ano", "mes", "ugb_id")))
  }

  msg("Conclu\u00eddo.")

  # --- Resumo final ---
  n_files <- dplyr::n_distinct(df$file_name)
  message(glue::glue("Processamento conclu\u00eddo: {n_files} ficheiro(s) processado(s) com sucesso."))

  return(df_limpeza_final)

}












#' Processar extractos do Razão C do e-SISTAFE a partir de ficheiros PDF
#'
#' Lê todos os ficheiros PDF de uma pasta, extrai as transacções e saldos
#' de cada extracto do Razão C do e-SISTAFE, e combina os resultados num
#' único tibble. Ficheiros com formato FOREX (USD/EUR) são excluídos por
#' padrão.
#'
#' @param source_path Caractere. Caminho para a pasta que contém os ficheiros
#'   PDF a processar. Obrigatório.
#' @param exclude_pattern Caractere. Expressão regular para excluir ficheiros
#'   pelo nome. Por padrão exclui ficheiros FOREX:
#'   \code{"CENTRAL USD|EXTRACTO DA CONTA FOREX EUR|EXTRACTO DA CONTA FOREX USD"}.
#'   Para não excluir nenhum ficheiro, usar \code{NULL}.
#' @param recursive Lógico. Se \code{TRUE}, a pesquisa de ficheiros PDF
#'   inclui subpastas. Por padrão \code{FALSE}.
#' @param quiet Lógico. Se \code{TRUE} (padrão), suprime as mensagens emitidas
#'   por ficheiro durante o processamento (por exemplo, quando um PDF não
#'   contém transacções). Se \code{FALSE}, as mensagens são apresentadas.
#'
#' @return Um tibble com uma linha por registo (movimentos, saldo inicial e
#'   saldo final) de todos os PDFs processados, contendo as colunas:
#'   \describe{
#'     \item{source_file}{Nome do ficheiro PDF de origem.}
#'     \item{unidade_gestao}{Nome da unidade de gestão extraído do cabeçalho.}
#'     \item{data}{Data do registo (\code{Date}).}
#'     \item{tipo}{Tipo de registo: \code{"MOVIMENTO"}, \code{"SALDO_INICIAL"} ou \code{"SALDO_FINAL"}.}
#'     \item{codigo_documento}{Código do documento (apenas em movimentos).}
#'     \item{valor_lancamento}{Valor do lançamento em MZN, negativo para créditos (C).}
#'     \item{dc1}{Indicador débito/crédito do lançamento (\code{"D"} ou \code{"C"}).}
#'     \item{saldo_atual}{Saldo acumulado após o lançamento.}
#'     \item{dc2}{Indicador débito/crédito do saldo.}
#'     \item{saldo_inicial_fim}{Valor do saldo inicial ou final (apenas nessas linhas).}
#'   }
#'
#' @details
#' A lógica de extracção trata os seguintes casos:
#' \itemize{
#'   \item PDFs com transacções: extrai movimentos linha a linha e calcula
#'     saldos inicial e final.
#'   \item PDFs sem transacções: retorna apenas as linhas SALDO_INICIAL e
#'     SALDO_FINAL com base nos valores do cabeçalho.
#'   \item Datas com espaços irregulares (ex: \code{"01 / 12 / 2025"}): são
#'     normalizadas automaticamente.
#'   \item Valores em formato português (ponto como separador de milhares,
#'     vírgula como decimal): convertidos correctamente.
#'   \item Créditos (C) são convertidos para valores negativos.
#' }
#'
#' O intervalo de datas do conjunto processado é guardado como atributo do
#' tibble retornado, acessível via \code{attr(df, "date_range_txt")}.
#'
#' @examples
#' \dontrun{
#' df_razao <- processar_extracto_razao_c(
#'   source_path = path_folder_source
#' )
#'
#' # Com mensagens visíveis e subpastas incluídas
#' df_razao <- processar_extracto_razao_c(
#'   source_path = path_folder_source,
#'   recursive   = TRUE,
#'   quiet       = FALSE
#' )
#'
#' # Sem exclusão de ficheiros FOREX
#' df_razao <- processar_extracto_razao_c(
#'   source_path     = path_folder_source,
#'   exclude_pattern = NULL
#' )
#' }
#'
#' @export
processar_extracto_razao_c <- function(
    source_path,
    exclude_pattern = "CENTRAL USD|EXTRACTO DA CONTA FOREX EUR|EXTRACTO DA CONTA FOREX USD",
    recursive       = FALSE,
    quiet           = TRUE
) {

  # ---- Helper interno: extrair tabela de um PDF ----
  extract_sistafe_table <- function(path_pdf) {

    raw_text <- pdftools::pdf_text(path_pdf)

    # -- helpers --
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

    # -- unidade_gestao --
    unidade_gestao <- raw_text[1] |>
      stringr::str_extract("Gestão:\\s*(.+)") |>
      stringr::str_remove("Gestão:\\s*") |>
      stringr::str_trim()

    # -- header dates --
    header_data_chr       <- extract_header_date(raw_text[1], "Data(?!\\s*Final)")
    header_data_final_chr <- extract_header_date(raw_text[1], "Data\\s*Final")

    header_data       <- suppressWarnings(lubridate::dmy(header_data_chr))
    header_data_final <- suppressWarnings(lubridate::dmy(header_data_final_chr))

    # -- header saldo (fallback when no transactions) --
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

    # -- extract transaction lines --
    lines <- raw_text |>
      stringr::str_split("\n") |>
      unlist() |>
      stringr::str_subset("^\\d{2}\\s*/\\s*\\d{2}\\s*/\\s*\\d{4}") |>
      stringr::str_squish()

    # -- no transactions: return saldo rows only --
    if (length(lines) == 0) {
      if (!quiet) {
        message(
          "Sem transacções em: ",
          basename(path_pdf),
          " — a retornar apenas saldos"
        )
      }

      return(
        dplyr::bind_rows(
          tibble::tibble(
            unidade_gestao   = unidade_gestao,
            data             = header_data,
            tipo             = "SALDO_INICIAL",
            codigo_documento = NA_character_,
            valor_lancamento = 0,
            dc1              = NA_character_,
            saldo_atual      = saldo_hdr_num,
            dc2              = saldo_hdr_dc,
            saldo_inicial_fim = saldo_hdr_num
          ),
          tibble::tibble(
            unidade_gestao   = unidade_gestao,
            data             = header_data_final,
            tipo             = "SALDO_FINAL",
            codigo_documento = NA_character_,
            valor_lancamento = 0,
            dc1              = NA_character_,
            saldo_atual      = saldo_hdr_num,
            dc2              = saldo_hdr_dc,
            saldo_inicial_fim = saldo_hdr_num
          )
        )
      )
    }

    # -- parse transactions --
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
        sep   = "\\s+",
        fill  = "right"
      ) |>
      dplyr::mutate(
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

        valor_lancamento = dplyr::if_else(
          dc1 == "C",
          -valor_lancamento,
          valor_lancamento
        ),

        unidade_gestao    = unidade_gestao,
        tipo              = "MOVIMENTO",
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

    # -- saldo inicial --
    data_inicio      <- if (!is.na(header_data)) header_data else df$data[1]
    saldo_inicial_calc <- df$saldo_atual[1] - df$valor_lancamento[1]

    saldo_inicial_row <- tibble::tibble(
      unidade_gestao    = unidade_gestao,
      data              = data_inicio,
      tipo              = "SALDO_INICIAL",
      codigo_documento  = NA_character_,
      valor_lancamento  = 0,
      dc1               = NA_character_,
      saldo_atual       = saldo_inicial_calc,
      dc2               = df$dc2[1],
      saldo_inicial_fim = saldo_inicial_calc
    )

    # -- saldo final --
    data_fim       <- if (!is.na(header_data_final)) header_data_final else df$data[nrow(df)]
    saldo_final_val <- df$saldo_atual[nrow(df)]

    saldo_final_row <- tibble::tibble(
      unidade_gestao    = unidade_gestao,
      data              = data_fim,
      tipo              = "SALDO_FINAL",
      codigo_documento  = NA_character_,
      valor_lancamento  = 0,
      dc1               = NA_character_,
      saldo_atual       = saldo_final_val,
      dc2               = df$dc2[nrow(df)],
      saldo_inicial_fim = saldo_final_val
    )

    dplyr::bind_rows(saldo_inicial_row, df, saldo_final_row)
  }

  # ---- Validação de argumentos ----
  if (!dir.exists(source_path)) {
    cli::cli_abort("A pasta {.path {source_path}} não existe.")
  }

  # ---- Listar ficheiros PDF ----
  list_pdf <- list.files(
    path       = source_path,
    pattern    = "\\.pdf$",
    full.names = TRUE,
    recursive  = recursive
  )

  if (!is.null(exclude_pattern)) {
    list_pdf <- stringr::str_subset(list_pdf, exclude_pattern, negate = TRUE)
  }

  if (length(list_pdf) == 0) {
    cli::cli_abort("Nenhum ficheiro PDF encontrado em {.path {source_path}}.")
  }

  # ---- Processar PDFs ----
  df <- list_pdf |>
    purrr::set_names(basename) |>
    purrr::map(extract_sistafe_table) |>
    purrr::list_rbind(names_to = "source_file")

  # ---- Calcular intervalo de datas ----
  date_min <- suppressWarnings(min(df$data, na.rm = TRUE))
  date_max <- suppressWarnings(max(df$data, na.rm = TRUE))

  date_range_txt <- if (is.finite(date_min) && is.finite(date_max)) {
    paste0(format(date_min, "%Y-%m-%d"), "_a_", format(date_max, "%Y-%m-%d"))
  } else {
    "sem_datas"
  }

  attr(df, "date_range_txt") <- date_range_txt

  df
}
















#' Gravar extracto do Razão C processado em Excel
#'
#' Grava um dataframe processado por \code{processar_extracto_razao_c()} num
#' ficheiro Excel, construindo automaticamente o nome do ficheiro a partir do
#' intervalo de datas do relatório e da data actual. Cria a pasta de destino
#' se não existir.
#'
#' @param df Um tibble processado por \code{processar_extracto_razao_c()}.
#'   Deve conter o atributo \code{date_range_txt} gerado por essa função.
#' @param output_folder Caractere. Caminho para a pasta de destino onde o
#'   ficheiro Excel sera gravado. Por padrao \code{"Dataout"}. A pasta e
#'   criada automaticamente se nao existir.
#' @param quiet Logico. Se \code{TRUE} (padrao), as mensagens de progresso
#'   sao suprimidas. Se \code{FALSE}, sao emitidas mensagens sobre a criacao
#'   da pasta e o caminho do ficheiro gravado.
#'
#' @return O caminho completo do ficheiro gravado, retornado de forma invisivel.
#'   Pode ser capturado com \code{path <- gravar_extracto_razao_c(df)} para
#'   uso posterior se necessario.
#'
#' @details
#' O nome do ficheiro e construido automaticamente no formato:
#' \code{Razao_C_<data_inicio>_a_<data_fim>_<YYYYMMDD>.xlsx}
#'
#' Por exemplo: \code{Razao_C_2025-01-01_a_2025-12-31_20260323.xlsx}
#'
#' Se o atributo \code{date_range_txt} nao estiver presente no dataframe
#' (por exemplo, se o objeto foi modificado apos o processamento), o nome
#' do ficheiro usa \code{"sem_datas"} como sufixo.
#'
#' @examples
#' \dontrun{
#' # Gravar com pasta padrao
#' gravar_extracto_razao_c(df_razao)
#'
#' # Gravar numa pasta personalizada
#' gravar_extracto_razao_c(df_razao, output_folder = "Data/processed")
#'
#' # Gravar com mensagens de progresso
#' gravar_extracto_razao_c(df_razao, quiet = FALSE)
#'
#' # Capturar o caminho do ficheiro gravado
#' path <- gravar_extracto_razao_c(df_razao, quiet = FALSE)
#' }
#'
#' @importFrom glue glue
#' @importFrom writexl write_xlsx
#'
#' @export
gravar_extracto_razao_c <- function(
    df,
    output_folder = "Dataout",
    quiet         = TRUE
) {

  # --- Mensagens internas ---
  msg <- function(...) {
    if (!quiet) message(...)
  }

  # --- Recuperar date_range_txt do atributo ---
  date_range_txt <- attr(df, "date_range_txt")

  if (is.null(date_range_txt)) {
    message(
      "Atributo 'date_range_txt' nao encontrado — ",
      "a usar 'sem_datas' no nome do ficheiro."
    )
    date_range_txt <- "sem_datas"
  }

  # --- Remover trailing slash se presente ---
  output_folder <- base::gsub("/$", "", output_folder)

  # --- Construir nome do ficheiro ---
  today     <- base::format(base::Sys.Date(), "%Y%m%d")
  file_name <- glue::glue("Razao-Cont_{date_range_txt}_{today}.xlsx")
  file_path <- base::file.path(output_folder, file_name)

  # --- Criar pasta se nao existir ---
  if (!base::dir.exists(output_folder)) {
    msg(glue::glue("Pasta '{output_folder}' nao encontrada — a criar..."))
    base::dir.create(output_folder, recursive = TRUE)
  }

  # --- Guardar ficheiro ---
  msg(glue::glue("A guardar ficheiro: {file_path}"))
  writexl::write_xlsx(df, file_path)
  msg("Concluido.")

  # --- Retornar caminho invisivel para uso posterior ---
  base::invisible(file_path)
}
