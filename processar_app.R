# Pipeline e-SISTAFE — Shiny app
# Runs: shiny::runApp("processar_app.R")
# Requires: callr, shinyjs (install.packages(c("callr","shinyjs")))

library(shiny)
library(bslib)
library(shinyjs)

# ── constants ──────────────────────────────────────────────────────────────

BOARD_PATH  <- "C:/Users/Joe/Documents/R/moz-data-pins/Data"
META_LOOKUP <- "Documents/esistafe_dim_lookup.xlsx"
DATA_BASE   <- "Data/demonstrativo_consolidado"
REPORT_QMD  <- "Relatório de Execução Orçamental.qmd"

MESES_PT <- c(
  "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
  "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
)

# ── helpers ────────────────────────────────────────────────────────────────

scan_months <- function() {
  if (!dir.exists(DATA_BASE)) return(character(0))
  d <- list.dirs(DATA_BASE, recursive = FALSE, full.names = FALSE)
  d <- d[grepl("^\\d{6}$", d)]
  d <- d[substr(d, 1, 4) %in% c("2025", "2026")]
  has_data <- vapply(d, function(yyyymm) {
    files <- list.files(file.path(DATA_BASE, yyyymm), pattern = "DemonstrativoConsolidado")
    length(files) > 0
  }, logical(1))
  sort(d[has_data], decreasing = TRUE)
}

month_label <- function(yyyymm) {
  yr <- substr(yyyymm, 1, 4)
  mo <- as.integer(substr(yyyymm, 5, 6))
  paste0(yr, "-", MESES_PT[mo])
}

read_log <- function(path) {
  raw <- paste(readLines(path, warn = FALSE), collapse = "\n")
  gsub("\033\\[[0-9;]*m", "", raw)   # strip ANSI colour codes
}

log_style <- paste(
  "font-family: monospace; font-size: 0.78rem; white-space: pre-wrap;",
  "max-height: 380px; overflow-y: auto;",
  "background: #f8f9fa; padding: 10px; border-radius: 4px; min-height: 60px;"
)

# ── UI ─────────────────────────────────────────────────────────────────────

ui <- page_sidebar(
  title   = "Pipeline e-SISTAFE",
  theme   = bs_theme(bootswatch = "flatly"),
  fillable = FALSE,
  useShinyjs(),

  sidebar = sidebar(
    width = 270,

    tags$b("Períodos a processar"),
    uiOutput("month_ui"),
    div(class = "mt-1 small",
      actionLink("btn_all",  "todos"),
      " · ",
      actionLink("btn_none", "nenhum")
    ),

    hr(),

    actionButton(
      "btn_process", "Processar dados",
      icon  = icon("play"),
      class = "btn-primary w-100 mb-2"
    ),
    actionButton(
      "btn_report", "Gerar Relatório",
      icon  = icon("file"),
      class = "btn-success w-100"
    ),

    hr(),

    actionButton(
      "btn_scan", "Actualizar lista",
      icon  = icon("sync"),
      class = "btn-outline-secondary btn-sm w-100"
    )
  ),

  navset_card_underline(
    id = "tabs",

    nav_panel("Processamento", value = "tab_proc",

      card(
        card_header("Log de Processamento"),
        div(style = log_style, textOutput("proc_log"))
      ),

      uiOutput("proc_status_card"),

      card(
        card_header("Verificação de Qualidade (sitrep)"),
        verbatimTextOutput("sitrep_out")
      )
    ),

    nav_panel("Relatório", value = "tab_rep",

      uiOutput("rep_status_card"),

      card(
        card_header("Log de Renderização"),
        div(style = log_style, textOutput("rep_log"))
      )
    )
  )
)

# ── Server ─────────────────────────────────────────────────────────────────

server <- function(input, output, session) {

  shinyjs::disable("btn_report")

  # months ----------------------------------------------------------------

  months_rv <- reactiveVal(scan_months())
  observeEvent(input$btn_scan, months_rv(scan_months()))

  output$month_ui <- renderUI({
    m <- months_rv()
    if (!length(m))
      return(p(class = "text-muted small",
               paste("Sem pastas em", DATA_BASE)))
    ch <- setNames(file.path(DATA_BASE, m), month_label(m))
    checkboxGroupInput("months", NULL, choices = ch, selected = character(0))
  })

  observeEvent(input$btn_all, {
    m  <- months_rv()
    ch <- setNames(file.path(DATA_BASE, m), month_label(m))
    updateCheckboxGroupInput(session, "months", selected = ch)
  })
  observeEvent(input$btn_none,
    updateCheckboxGroupInput(session, "months", selected = character(0))
  )

  # pipeline --------------------------------------------------------------

  proc_job  <- reactiveVal(NULL)
  proc_logf <- reactiveVal(NULL)
  proc_log  <- reactiveVal("Aguardando...")
  proc_done <- reactiveVal(FALSE)
  proc_res  <- reactiveVal(NULL)

  observeEvent(input$btn_process, {
    req(length(input$months) > 0)
    proc_done(FALSE)
    proc_res(NULL)
    proc_log("")
    shinyjs::disable("btn_report")
    nav_select("tabs", "tab_proc")

    lf <- tempfile(fileext = ".txt")
    file.create(lf)
    proc_logf(lf)

    paths   <- input$months
    meta    <- META_LOOKUP
    board_p <- BOARD_PATH
    outdir  <- "Dataout/"
    wd      <- getwd()

    job <- callr::r_bg(
      func = function(paths, meta, board_p, outdir) {
        suppressPackageStartupMessages({
          library(tidyverse)
          library(readxl)
          library(writexl)
          library(janitor)
          library(arrow)
          library(pins)
          library(easystafe)
        })

        message("[1/5] A carregar lookups...")
        lookups <- carregar_lookups_esistafe(meta)

        message("[2/5] A processar ", length(paths), " período(s)...")
        df <- map(paths, \(p) {
          message("  → ", basename(p))
          processar_extracto_esistafe(
            source_path           = p,
            df_ugb_lookup         = lookups$ugb,
            include_percent       = FALSE,
            include_file_metadata = TRUE,
            include_metrica       = TRUE,
            correct_negatives     = TRUE,
            quiet                 = FALSE
          )
        }) |>
          list_rbind() |>
          adicionar_lookups_esistafe(lookups)
        message("  Total: ", nrow(df), " linhas")

        message("[3/5] A verificar qualidade (sitrep)...")
        gaps <- sitrep_processamento_esistafe(df, lookups, lookup_path = meta)

        message("[4/5] A gravar ficheiros...")
        df_db <- config_para_duckdb(df)
        gravar_esistafe(df, output_folder = outdir, quiet = TRUE)

        message("[5/5] A escrever pinos...")
        board <- board_folder(board_p, versioned = TRUE)
        pin_write(
          board, df,
          name        = "demostrativo_consolidado_compilado",
          type        = "parquet",
          description = "Demonstrativo Consolidado compilado (easystafe)."
        )
        pin_write(
          board, df_db,
          name        = "esistafe_dm_duckdb",
          type        = "parquet",
          description = "Demonstrativo Consolidado — formato DuckDB (easystafe)."
        )

        message("✓ Processamento concluído.")
        list(gaps = gaps, nrows = nrow(df))
      },
      args   = list(paths = paths, meta = meta, board_p = board_p, outdir = outdir),
      stdout = lf,
      stderr = lf,
      wd     = wd
    )

    proc_job(job)
  })

  observe({
    job <- proc_job()
    if (is.null(job) || proc_done()) return()
    invalidateLater(800)
    proc_log(read_log(proc_logf()))
    if (!job$is_alive()) {
      proc_done(TRUE)
      res <- tryCatch(
        job$get_result(),
        error = function(e) list(error = conditionMessage(e))
      )
      proc_res(res)
      if (is.null(res$error)) shinyjs::enable("btn_report")
    }
  })

  output$proc_log <- renderText(proc_log())

  output$proc_status_card <- renderUI({
    job <- proc_job()
    if (is.null(job)) return(NULL)
    done <- proc_done()
    res  <- proc_res()
    card(card_body(
      if (!done)
        div(class = "d-flex align-items-center gap-2",
          div(class = "spinner-border spinner-border-sm text-primary"),
          span("A processar dados..."))
      else if (!is.null(res$error))
        div(class = "text-danger", strong("Erro: "), res$error)
      else
        div(class = "text-success",
          strong(icon("circle-check"), " Concluído — "),
          format(res$nrows, big.mark = "."), " linhas processadas.")
    ))
  })

  output$sitrep_out <- renderPrint({
    req(proc_done())
    res <- proc_res()
    if (is.null(res) || !is.null(res$error)) return(invisible(NULL))
    print(res$gaps)
  })

  # report ----------------------------------------------------------------

  rep_job  <- reactiveVal(NULL)
  rep_logf <- reactiveVal(NULL)
  rep_log  <- reactiveVal("")
  rep_done <- reactiveVal(FALSE)
  rep_ok   <- reactiveVal(FALSE)
  rep_err  <- reactiveVal(NULL)

  observeEvent(input$btn_report, {
    rep_done(FALSE)
    rep_ok(FALSE)
    rep_err(NULL)
    rep_log("")
    nav_select("tabs", "tab_rep")

    lf <- tempfile(fileext = ".txt")
    file.create(lf)
    rep_logf(lf)

    # absolute path avoids working-directory ambiguity in the subprocess
    qmd_abs <- normalizePath(REPORT_QMD, mustWork = FALSE)

    job <- callr::r_bg(
      func = function(qmd) {
        message("A renderizar: ", qmd)
        quarto::quarto_render(input = qmd)
        message("✓ Relatório gerado.")
      },
      args   = list(qmd = qmd_abs),
      stdout = lf,
      stderr = lf,
      wd     = getwd()
    )

    rep_job(job)
  })

  observe({
    job <- rep_job()
    if (is.null(job) || rep_done()) return()
    invalidateLater(800)
    rep_log(read_log(rep_logf()))
    if (!job$is_alive()) {
      rep_done(TRUE)
      result <- tryCatch(
        { job$get_result(); list(ok = TRUE,  err = NULL) },
        error = function(e) list(ok = FALSE, err = conditionMessage(e))
      )
      rep_ok(result$ok)
      rep_err(result$err)
    }
  })

  output$rep_log <- renderText(rep_log())

  output$rep_status_card <- renderUI({
    job <- rep_job()
    if (is.null(job))
      return(card(card_body(
        p(class = "text-muted",
          "Clique em 'Gerar Relatório' para iniciar a renderização."))))
    done <- rep_done()
    ok   <- rep_ok()
    err  <- rep_err()
    card(card_body(
      if (!done)
        div(class = "d-flex align-items-center gap-2",
          div(class = "spinner-border spinner-border-sm text-success"),
          span("A renderizar relatório..."))
      else if (ok) {
        html <- normalizePath(sub("\\.qmd$", ".html", REPORT_QMD), mustWork = FALSE)
        url  <- paste0("file:///", gsub("\\\\", "/", html))
        div(class = "alert alert-success mb-0",
          strong(icon("circle-check"), " Relatório gerado! "),
          tags$a(href = url, target = "_blank", "Abrir no browser"))
      } else
        div(class = "alert alert-danger mb-0",
          tags$b("Erro: "),
          if (!is.null(err)) err else "Erro desconhecido. Ver log abaixo.")
    ))
  })
}

shinyApp(ui, server)
