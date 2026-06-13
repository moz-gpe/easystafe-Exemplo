library(shiny)
library(bslib)
library(dplyr)
library(duckdb)
library(DBI)
library(pins)
library(glue)
library(DT)

# --- Data ----------------------------------------------------------------

board <- board_folder(
  "C:/Users/Joe/Documents/R/moz-data-pins/Data",
  versioned = TRUE
)

financial_cols <- c(
  "dotacao_inicial", "dotacao_revista", "dotacao_actualizada_da",
  "dotacao_disponivel", "dotacao_cabimentada_dc", "ad_fundos_concedidos_af",
  "despesa_paga_via_directa_dp", "ad_fundos_desp_paga_vd_afdp",
  "ad_fundos_liquidados_laf", "despesa_liquidada_via_directa_lvd",
  "liq_ad_fundos_via_directa_lafvd"
)

df_raw <- pin_read(board, "demostrativo_consolidado_compilado") |>
  filter(data_tipo == "Valor") |>
  mutate(across(all_of(financial_cols), \(x) round(x / 1000, 0)))

con <- dbConnect(duckdb())
dbWriteTable(con, "orcamento", df_raw)

mes_order <- c(
  "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
  "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
)

fmt_mil <- function(x)
  format(round(x), big.mark = ".", decimal.mark = ",", scientific = FALSE)

# --- UI ------------------------------------------------------------------

ui <- page_sidebar(
  title = "Execução Orçamental — Sector da Educação",
  theme = bs_theme(bootswatch = "flatly"),

  sidebar = sidebar(
    selectInput("ano", "Ano",
      choices = sort(unique(df_raw$ano), decreasing = TRUE)
    ),
    selectInput("mes", "Mês", choices = NULL)
  ),

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

  card(
    card_header("Execução por Tipo de Reporte"),
    DTOutput("tbl_reporte")
  )
)

# --- Server --------------------------------------------------------------

server <- function(input, output, session) {

  observe({
    meses <- df_raw |>
      filter(ano == as.integer(input$ano)) |>
      pull(mes) |>
      unique()
    meses <- meses[order(match(meses, mes_order))]
    updateSelectInput(session, "mes", choices = meses, selected = tail(meses, 1))
  })

  df_filt <- reactive({
    req(input$ano, input$mes)
    dbGetQuery(con, glue_sql(
      "SELECT * FROM orcamento WHERE ano = {as.integer(input$ano)} AND mes = {input$mes}",
      .con = con
    ))
  })

  totals <- reactive({
    df_filt() |>
      summarise(
        dotacao = sum(dotacao_actualizada_da,      na.rm = TRUE),
        paga    = sum(ad_fundos_desp_paga_vd_afdp, na.rm = TRUE)
      ) |>
      mutate(exec_pct = if_else(dotacao > 0, paga / dotacao * 100, NA_real_))
  })

  output$vb_dotacao <- renderText(paste(fmt_mil(totals()$dotacao), "M Mts."))
  output$vb_paga    <- renderText(paste(fmt_mil(totals()$paga),    "M Mts."))
  output$vb_exec    <- renderText(paste0(round(totals()$exec_pct, 1), "%"))

  output$tbl_reporte <- renderDT({
    df_filt() |>
      group_by(`Tipo de Reporte` = reporte_tipo) |>
      summarise(
        `Orçamento (M Mts.)` = sum(dotacao_actualizada_da,      na.rm = TRUE),
        `Pago (M Mts.)`      = sum(ad_fundos_desp_paga_vd_afdp, na.rm = TRUE),
        .groups = "drop"
      ) |>
      mutate(
        `Exec %` = round(
          if_else(`Orçamento (M Mts.)` > 0,
                  `Pago (M Mts.)` / `Orçamento (M Mts.)` * 100,
                  NA_real_),
          1
        )
      ) |>
      datatable(options = list(dom = "t", paging = FALSE), rownames = FALSE)
  })

  onStop(function() dbDisconnect(con, shutdown = TRUE))
}

shinyApp(ui, server)
