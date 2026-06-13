library(shiny)
library(bslib)
library(dplyr)
library(duckdb)
library(DBI)
library(pins)
library(DT)
library(querychat)

# --- Data ----------------------------------------------------------------

board <- board_folder(
  "C:/Users/Joe/Documents/R/moz-data-pins/Data",
  versioned = TRUE
)

financial_cols <- c(
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
  "liq_ad_fundos_via_directa_lafvd"
)

df_raw <- pin_read(board, "demostrativo_consolidado_compilado") |>
  filter(data_tipo == "Valor") |>
  mutate(across(all_of(financial_cols), \(x) round(x / 1000, 0)))

con <- dbConnect(duckdb())
dbWriteTable(con, "orcamento", df_raw)

fmt_mil <- function(x) {
  format(round(x), big.mark = ".", decimal.mark = ",", scientific = FALSE)
}

qc <- querychat(
  con,
  "orcamento",
  client = ellmer::chat_anthropic(),
  data_description = "
    Budget execution data for the Mozambican education sector, sourced from the
    e-SISTAFE system. All financial values are in millions of Meticais (M Mts.).
    Key columns:
      - ano: fiscal year (integer)
      - mes: month name in Portuguese (e.g. 'Janeiro', 'Fevereiro', ..., 'Dezembro')
      - reporte_tipo: budget type — 'Funcionamento', 'Investimento Interno', 'Investimento Externo'
      - ambito: scope — 'Provincial' or 'Central'
      - provincia: province name
      - descricao: institution name (Central institutions)
      - dotacao_actualizada_da: updated budget allocation (M Mts.)
      - ad_fundos_desp_paga_vd_afdp: paid expenditure (M Mts.)
      - exec_rate: execution rate as a decimal (paid / allocation)
  ",
  greeting = "
    Olá! Pode explorar os dados de execução orçamental do sector da Educação.
    Alguns exemplos para começar:
    - \"Mostra a execução total do último mês disponível\"
    - \"Compara a execução por província em 2026\"
    - \"Qual é a taxa de execução do Investimento Interno por ano?\"
  "
)

qc_model <- tryCatch(qc$client()$get_model(), error = function(e) "—")

# --- UI ------------------------------------------------------------------

ui <- page_sidebar(
  title = "Execução Orçamental — Sector da Educação",
  theme = bs_theme(bootswatch = "flatly"),

  sidebar = qc$sidebar(),

  layout_columns(
    col_widths = c(4, 4, 4),
    value_box(
      title = "Orçamento Actualizado",
      value = textOutput("vb_dotacao"),
      theme = "primary"
    ),
    value_box(
      title = "Despesa Paga",
      value = textOutput("vb_paga"),
      theme = "success"
    ),
    value_box(
      title = "Taxa de Execução",
      value = textOutput("vb_exec"),
      theme = "info"
    )
  ),

  layout_columns(
    col_widths = c(8, 4),
    card(
      card_header("Execução por Tipo de Reporte"),
      DTOutput("tbl_reporte")
    ),
    card(
      card_header("SQL gerado"),
      verbatimTextOutput("sql_query"),
      card_footer(paste("Modelo:", qc_model))
    )
  )
)

# --- Server --------------------------------------------------------------

server <- function(input, output, session) {
  qc_vals <- qc$server()

  totals <- reactive({
    df <- qc_vals$df()
    req(nrow(df) > 0)
    df |>
      summarise(
        dotacao = sum(dotacao_actualizada_da, na.rm = TRUE),
        paga = sum(ad_fundos_desp_paga_vd_afdp, na.rm = TRUE)
      ) |>
      mutate(exec_pct = if_else(dotacao > 0, paga / dotacao * 100, NA_real_))
  })

  output$vb_dotacao <- renderText(paste(fmt_mil(totals()$dotacao), "M Mts."))
  output$vb_paga <- renderText(paste(fmt_mil(totals()$paga), "M Mts."))
  output$vb_exec <- renderText(paste0(round(totals()$exec_pct, 1), "%"))

  output$tbl_reporte <- renderDT({
    qc_vals$df() |>
      group_by(`Tipo de Reporte` = reporte_tipo) |>
      summarise(
        `Orçamento (M Mts.)` = sum(dotacao_actualizada_da, na.rm = TRUE),
        `Pago (M Mts.)` = sum(ad_fundos_desp_paga_vd_afdp, na.rm = TRUE),
        .groups = "drop"
      ) |>
      mutate(
        `Exec %` = round(
          if_else(
            `Orçamento (M Mts.)` > 0,
            `Pago (M Mts.)` / `Orçamento (M Mts.)` * 100,
            NA_real_
          ),
          1
        )
      ) |>
      datatable(options = list(dom = "t", paging = FALSE), rownames = FALSE)
  })

  output$sql_query <- renderText({
    sql <- qc_vals$sql()
    if (is.null(sql) || sql == "") "Nenhuma consulta ainda." else sql
  })

  onStop(function() {
    dbDisconnect(con, shutdown = TRUE)
    qc$cleanup()
  })
}

shinyApp(ui, server)
