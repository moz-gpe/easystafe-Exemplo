# Pipeline e-SISTAFE — Shiny app
# Runs: shiny::runApp("processar_app.R")
# Requires: callr, shinyjs (install.packages(c("callr","shinyjs")))

library(shiny)
library(bslib)
library(shinyjs)

# Rendering the report used to drop quarto intermediates (relatorio_tmp.html,
# *_files/, *.rmarkdown) into this app's working directory. Under RStudio that
# file churn tripped Shiny's autoreload, restarting the session mid-render
# (symptom: the page "flashes" and jumps back to the first tab). Disable it —
# this is a pipeline-runner app, not something we hot-edit while it's running.
options(shiny.autoreload = FALSE)

# ── constants ──────────────────────────────────────────────────────────────

BOARD_PATH  <- "C:/Users/Joe/Documents/R/moz-data-pins/Data"
META_LOOKUP <- "documents/esistafe_dim_lookup.xlsx"
DATA_BASE   <- "data/demonstrativo_consolidado"
REPORT_QMD  <- "Relatório de Execução Orçamental.qmd"

MESES_PT <- c(
  "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
  "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
)

# Serve the finished report over the app's own http origin. browseURL() can't
# be used: it runs on the R server and asks the IDE (Positron/RStudio) to
# launch an external browser, which fails. Instead we expose a static dir and
# hand the client an in-app URL it can open in a new tab. The report is a
# single self-contained HTML, so one file is all we serve (ASCII name, since
# the project-dir copy has accented characters that complicate URLs).
REPORT_WWW  <- file.path(tempdir(), "relatorio_www")
REPORT_FILE <- "relatorio.html"
REPORT_URL  <- paste0("relatorio/", REPORT_FILE)   # client-visible path
dir.create(REPORT_WWW, showWarnings = FALSE, recursive = TRUE)
addResourcePath("relatorio", REPORT_WWW)

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

strip_ansi <- function(x) gsub("\033\\[[0-9;]*m", "", x)

# Drain stdout + stderr pipes into a reactiveVal, appending to existing text.
#
# processx note (verified on Windows): reading a pipe with `timeout = 0`
# returns NOTHING even when data is waiting — that silent-empty read was the
# root cause of the blank log. The reliable non-blocking pattern is:
#   * while the job is alive: poll_io(0) first (moves pipe data into the
#     connection buffer and reports which streams are "ready"), then read
#     ONLY the ready streams with the default timeout — proven to return in
#     a few ms, so it never freezes the Shiny reactive loop.
#   * once the job has exited (final = TRUE): read unconditionally with the
#     default timeout. EOF is guaranteed, so the read returns immediately and
#     captures the last line (which the poll-gated path races past at exit).
drain_pipes <- function(job, log_rv, final = FALSE) {
  if (final) {
    out <- tryCatch(job$read_output_lines(), error = function(e) character(0))
    err <- tryCatch(job$read_error_lines(), error = function(e) character(0))
  } else {
    st  <- tryCatch(job$poll_io(0), error = function(e) c(output = "", error = ""))
    out <- if (identical(st[["output"]], "ready"))
             tryCatch(job$read_output_lines(), error = function(e) character(0)) else character(0)
    err <- if (identical(st[["error"]], "ready"))
             tryCatch(job$read_error_lines(), error = function(e) character(0)) else character(0)
  }
  lines <- strip_ansi(c(out, err))
  lines <- lines[nchar(lines) > 0]
  if (!length(lines)) return(invisible(NULL))
  prev <- log_rv()
  log_rv(if (nchar(prev) > 0) paste(prev, paste(lines, collapse = "\n"), sep = "\n")
         else paste(lines, collapse = "\n"))
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

  # best-effort auto-open of the finished report in a new browser tab; the
  # visible link in the status card is the reliable fallback if blocked
  tags$head(tags$script(HTML(
    "Shiny.addCustomMessageHandler('open_report', function(url){ window.open(url, '_blank'); });"
  ))),

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

    paths   <- input$months
    meta    <- META_LOOKUP
    board_p <- BOARD_PATH
    outdir  <- "dataout/"

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
      stdout = "|",
      stderr = "|",
      wd     = getwd()
    )

    proc_job(job)
  })

  observe({
    job <- proc_job()
    if (is.null(job) || proc_done()) return()
    invalidateLater(800)
    drain_pipes(job, proc_log)
    if (!job$is_alive()) {
      drain_pipes(job, proc_log, final = TRUE)
      res <- tryCatch(
        job$get_result(),
        error = function(e) list(error = conditionMessage(e))
      )
      proc_res(res)
      proc_done(TRUE)
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
  rep_log  <- reactiveVal("")
  rep_done <- reactiveVal(FALSE)
  rep_ok   <- reactiveVal(FALSE)
  rep_err  <- reactiveVal(NULL)
  rep_url  <- reactiveVal(NULL)

  observeEvent(input$btn_report, {
    rep_done(FALSE)
    rep_ok(FALSE)
    rep_err(NULL)
    rep_url(NULL)
    rep_log("")
    nav_select("tabs", "tab_rep")

    qmd_abs <- normalizePath(REPORT_QMD, mustWork = FALSE)

    job <- callr::r_bg(
      func = function(qmd) {
        if (!file.exists(qmd))
          stop("QMD não encontrado: ", qmd)

        out_html <- sub("\\.qmd$", ".html", qmd)

        # Render in an ISOLATED temp dir — never in the app's working directory.
        # Two reasons:
        #   1. Quarto on Windows rejects non-ASCII/special characters in the
        #      input filename, so we need a plain-ASCII copy anyway.
        #   2. Writing quarto intermediates into the watched app dir restarts
        #      the Shiny session mid-render. Keeping all churn under tempdir()
        #      means the app dir only ever sees the single final HTML copy.
        # The QMD reads its data via absolute pins paths and uses only inline
        # CSS, so it has no dependence on being rendered from the project dir.
        rdir      <- file.path(tempdir(), paste0("relatorio_", as.integer(Sys.time())))
        dir.create(rdir, showWarnings = FALSE, recursive = TRUE)
        tmp_qmd   <- file.path(rdir, "relatorio.qmd")
        tmp_html  <- file.path(rdir, "relatorio.html")
        on.exit(unlink(rdir, recursive = TRUE), add = TRUE)

        # Inject `embed-resources: true` so quarto emits ONE self-contained
        # HTML (no _files/ sidecar), making the copy back a single-file op.
        src <- readLines(qmd, warn = FALSE)
        hit <- grep("^\\s*html:\\s*$", src)[1]
        if (!is.na(hit))
          src <- append(src, "    embed-resources: true", after = hit)
        writeLines(src, tmp_qmd)

        message("A renderizar relatório (em ", rdir, ")...")
        quarto::quarto_render(input = tmp_qmd)
        message("Render concluído. A verificar output...")

        if (!file.exists(tmp_html)) {
          html_files <- list.files(rdir, pattern = "\\.html$")
          stop("HTML não encontrado após renderização. Ficheiros HTML encontrados: ",
               if (length(html_files)) paste(html_files, collapse = ", ") else "(nenhum)")
        }

        file.copy(tmp_html, out_html, overwrite = TRUE)
        message("✓ Relatório gerado: ", basename(out_html))
      },
      args   = list(qmd = qmd_abs),
      stdout = "|",
      stderr = "|",
      wd     = getwd()
    )

    rep_job(job)
  })

  observe({
    job <- rep_job()
    if (is.null(job) || rep_done()) return()
    invalidateLater(800)
    drain_pipes(job, rep_log)
    if (!job$is_alive()) {
      drain_pipes(job, rep_log, final = TRUE)
      result <- tryCatch(
        { job$get_result(); list(ok = TRUE,  err = NULL) },
        error = function(e) list(ok = FALSE, err = conditionMessage(e))
      )
      rep_ok(result$ok)
      rep_err(result$err)
      rep_done(TRUE)
    }
  })

  # fires once when rep_done flips to TRUE — publish the report and notify
  observeEvent(rep_done(), {
    req(rep_done())
    if (rep_ok()) {
      html   <- normalizePath(sub("\\.qmd$", ".html", REPORT_QMD), mustWork = FALSE)
      served <- file.path(REPORT_WWW, REPORT_FILE)
      if (file.exists(html) && file.copy(html, served, overwrite = TRUE)) {
        rep_url(REPORT_URL)
        # try to pop it open automatically; link in the card is the fallback
        session$sendCustomMessage("open_report", REPORT_URL)
        showNotification(
          "Relatório gerado. Se não abrir automaticamente, use o link no painel.",
          type = "message", duration = 10
        )
      } else {
        showNotification(
          "Relatório renderizado mas não foi possível publicá-lo para visualização.",
          type = "warning", duration = 15
        )
      }
    } else {
      err_msg <- if (!is.null(rep_err())) rep_err() else "Erro desconhecido."
      showNotification(paste0("Erro: ", err_msg), type = "error", duration = 20)
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
      else if (ok)
        div(class = "alert alert-success mb-0",
          strong(icon("circle-check"), " Relatório gerado. "),
          if (!is.null(rep_url()))
            tags$a(href = rep_url(), target = "_blank", class = "alert-link",
                   "Abrir relatório numa nova janela."))
      else
        div(class = "alert alert-danger mb-0",
          tags$b("Erro: "),
          if (!is.null(err)) err else "Erro desconhecido. Ver log abaixo.")
    ))
  })
}

shinyApp(ui, server)
