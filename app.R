# Pipeline e-SISTAFE — Shiny app
# Deploy: Posit Connect Cloud. Users upload e-SISTAFE Excel files; the app
# detects periods from filenames, processes, and serves downloads.
# Requires: callr, shinyjs (install.packages(c("callr","shinyjs")))

library(shiny)
library(bslib)
library(shinyjs)

# Rendering the report used to drop quarto intermediates into this app's
# working directory. Under RStudio that file churn tripped Shiny's autoreload,
# restarting the session mid-render. Disable it — this is a pipeline-runner
# app, not something we hot-edit while it's running.
options(shiny.autoreload = FALSE)

# ── constants ──────────────────────────────────────────────────────────────

META_LOOKUP <- "documents/esistafe_dim_lookup.xlsx"
REPORT_QMD  <- "Relatório de Execução Orçamental.qmd"

MESES_PT <- c(
  "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
  "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
)

# Serve rendered reports over the app's own HTTP origin. Each session writes
# to a unique filename so concurrent users don't overwrite each other.
REPORT_WWW <- file.path(tempdir(), "relatorio_www")
dir.create(REPORT_WWW, showWarnings = FALSE, recursive = TRUE)
addResourcePath("relatorio", REPORT_WWW)

# ── helpers ────────────────────────────────────────────────────────────────

# Extract YYYYMM from an e-SISTAFE filename by parsing the first YYYYMMDD
# pattern (data_reporte). Returns NA if no 8-digit date is found.
extract_yyyymm <- function(fname) {
  m <- regmatches(fname, gregexpr("\\d{8}", fname))[[1]]
  if (!length(m)) return(NA_character_)
  substr(m[1], 1, 6)
}

yyyymm_label <- function(yyyymm) {
  yr <- substr(yyyymm, 1, 4)
  mo <- as.integer(substr(yyyymm, 5, 6))
  paste0(yr, "-", MESES_PT[mo])
}

strip_ansi <- function(x) gsub("\033\\[[0-9;]*m", "", x)

# Copy uploaded files into per-period YYYYMM subdirs under base_dir.
# Returns a named character vector: names = YYYYMM, values = folder paths.
# Files whose names don't contain a parseable date are silently skipped.
stage_files <- function(files_df, base_dir) {
  dirs <- list()
  for (i in seq_len(nrow(files_df))) {
    fname  <- files_df$name[i]
    fpath  <- files_df$datapath[i]
    yyyymm <- extract_yyyymm(fname)
    if (is.na(yyyymm)) next
    dir_path <- file.path(base_dir, yyyymm)
    dir.create(dir_path, showWarnings = FALSE, recursive = TRUE)
    file.copy(fpath, file.path(dir_path, fname), overwrite = TRUE)
    dirs[[yyyymm]] <- dir_path
  }
  unlist(dirs)
}

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
    err <- tryCatch(job$read_error_lines(),  error = function(e) character(0))
  } else {
    st  <- tryCatch(job$poll_io(0), error = function(e) c(output = "", error = ""))
    out <- if (identical(st[["output"]], "ready"))
             tryCatch(job$read_output_lines(), error = function(e) character(0)) else character(0)
    err <- if (identical(st[["error"]], "ready"))
             tryCatch(job$read_error_lines(),  error = function(e) character(0)) else character(0)
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
  title    = "Pipeline e-SISTAFE",
  theme    = bs_theme(bootswatch = "flatly"),
  fillable = FALSE,
  useShinyjs(),

  tags$head(tags$script(HTML(
    "Shiny.addCustomMessageHandler('open_report', function(url){ window.open(url, '_blank'); });"
  ))),

  sidebar = sidebar(
    width = 270,

    fileInput(
      "files", "Ficheiros e-SISTAFE (.xlsx)",
      multiple    = TRUE,
      accept      = ".xlsx",
      buttonLabel = "Seleccionar...",
      placeholder = "Nenhum ficheiro seleccionado"
    ),

    uiOutput("periods_ui"),

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

  # Session-scoped paths — unique per user so concurrent sessions don't
  # collide on shared temp directories.
  out_dir     <- file.path(tempdir(), paste0("esistafe_out_",   session$token))
  stage_base  <- file.path(tempdir(), paste0("esistafe_stage_", session$token))
  rep_file    <- paste0("relatorio_", session$token, ".html")
  rep_url_path <- paste0("relatorio/", rep_file)
  rep_dest    <- file.path(REPORT_WWW, rep_file)
  dir.create(out_dir,    recursive = TRUE, showWarnings = FALSE)
  dir.create(stage_base, recursive = TRUE, showWarnings = FALSE)

  # ── file upload & period detection ───────────────────────────────────────

  output$periods_ui <- renderUI({
    req(input$files)
    yyyymms <- vapply(input$files$name, extract_yyyymm, character(1))
    valid   <- sort(yyyymms[!is.na(yyyymms)])

    if (!length(valid))
      return(p(class = "text-warning small",
               "Não foi possível detectar períodos nos nomes dos ficheiros."))

    tbl   <- sort(table(valid))
    items <- mapply(
      function(yyyymm, n) tags$li(paste0(yyyymm_label(yyyymm), " — ", n, " ficheiro(s)")),
      names(tbl), as.integer(tbl),
      SIMPLIFY = FALSE
    )
    tagList(
      p(class = "small mb-1",
        strong(length(tbl), " período(s) detectado(s):")),
      tags$ul(class = "small ps-3", items)
    )
  })

  # ── pipeline ──────────────────────────────────────────────────────────────

  proc_job  <- reactiveVal(NULL)
  proc_log  <- reactiveVal("Aguardando...")
  proc_done <- reactiveVal(FALSE)
  proc_res  <- reactiveVal(NULL)

  observeEvent(input$btn_process, {
    req(input$files)
    proc_done(FALSE)
    proc_res(NULL)
    proc_log("")
    shinyjs::disable("btn_report")
    nav_select("tabs", "tab_proc")

    staged <- stage_files(input$files, stage_base)
    req(length(staged) > 0)

    paths  <- unname(staged)
    meta   <- normalizePath(META_LOOKUP, mustWork = FALSE)
    outdir <- out_dir

    job <- callr::r_bg(
      func = function(paths, meta, outdir) {
        suppressPackageStartupMessages({
          library(tidyverse)
          library(readxl)
          library(writexl)
          library(janitor)
          library(arrow)
          library(easystafe)
        })

        message("[1/4] A carregar lookups...")
        lookups <- carregar_lookups_esistafe(meta)

        message("[2/4] A processar ", length(paths), " período(s)...")
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
          adicionar_lookups_esistafe(lookups) |>
          escolaR::padronizar_geo_dim(col_provincia = "provincia", col_distrito = "distrito")
        message("  Total: ", nrow(df), " linhas")

        message("[3/4] A verificar qualidade (sitrep)...")
        gaps <- sitrep_processamento_esistafe(df, lookups, lookup_path = meta)

        message("[4/4] A gravar ficheiros...")
        if (!dir.exists(outdir)) dir.create(outdir, recursive = TRUE)
        arrow::write_parquet(df, file.path(outdir, "e-SISTAFE_processado.parquet"))
        writexl::write_xlsx(df, file.path(outdir, "e-SISTAFE_processado.xlsx"))

        message("✓ Processamento concluído.")
        list(gaps = gaps, nrows = nrow(df))
      },
      args   = list(paths = paths, meta = meta, outdir = outdir),
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
        div(
          div(class = "text-success mb-2",
            strong(icon("circle-check"), " Concluído — "),
            format(res$nrows, big.mark = "."), " linhas processadas."),
          div(class = "d-flex gap-2",
            downloadButton("dl_parquet", "Parquet",
                           class = "btn-outline-primary btn-sm"),
            downloadButton("dl_excel",   "Excel",
                           class = "btn-outline-success btn-sm")
          )
        )
    ))
  })

  output$dl_parquet <- downloadHandler(
    filename = "e-SISTAFE_processado.parquet",
    content  = function(file) {
      file.copy(file.path(out_dir, "e-SISTAFE_processado.parquet"), file)
    }
  )

  output$dl_excel <- downloadHandler(
    filename = "e-SISTAFE_processado.xlsx",
    content  = function(file) {
      file.copy(file.path(out_dir, "e-SISTAFE_processado.xlsx"), file)
    }
  )

  output$sitrep_out <- renderPrint({
    req(proc_done())
    res <- proc_res()
    if (is.null(res) || !is.null(res$error)) return(invisible(NULL))
    print(res$gaps)
  })

  # ── report ────────────────────────────────────────────────────────────────

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

    qmd_abs     <- normalizePath(REPORT_QMD, mustWork = FALSE)
    abs_parquet <- normalizePath(
      file.path(out_dir, "e-SISTAFE_processado.parquet"),
      mustWork = FALSE, winslash = "/"
    )

    job <- callr::r_bg(
      func = function(qmd, abs_parquet, report_dest) {
        if (!file.exists(qmd))
          stop("QMD não encontrado: ", qmd)

        # Render in an isolated temp dir — never in the app's working directory.
        # Two reasons:
        #   1. Quarto on Windows rejects non-ASCII/special characters in the
        #      input filename, so we need a plain-ASCII copy anyway.
        #   2. Writing quarto intermediates into the watched app dir restarts
        #      the Shiny session mid-render. Keeping all churn under tempdir()
        #      means the app dir only ever sees the single final HTML copy.
        rdir     <- file.path(tempdir(), paste0("relatorio_", as.integer(Sys.time())))
        dir.create(rdir, showWarnings = FALSE, recursive = TRUE)
        tmp_qmd  <- file.path(rdir, "relatorio.qmd")
        tmp_html <- file.path(rdir, "relatorio.html")
        on.exit(unlink(rdir, recursive = TRUE), add = TRUE)

        # Inject `embed-resources: true` so quarto emits ONE self-contained
        # HTML (no _files/ sidecar), making the copy back a single-file op.
        # Also inject the absolute parquet path so the report resolves it
        # correctly when rendered from an isolated temp directory.
        src <- readLines(qmd, warn = FALSE)
        hit <- grep("^\\s*html:\\s*$", src)[1]
        if (!is.na(hit))
          src <- append(src, "    embed-resources: true", after = hit)
        src <- sub(
          'parquet_path <- "dataout/e-SISTAFE_processado.parquet"',
          paste0('parquet_path <- "', abs_parquet, '"'),
          src, fixed = TRUE
        )
        writeLines(src, tmp_qmd)

        message("A renderizar relatório (em ", rdir, ")...")
        quarto::quarto_render(input = tmp_qmd)
        message("Render concluído. A verificar output...")

        if (!file.exists(tmp_html)) {
          html_files <- list.files(rdir, pattern = "\\.html$")
          stop("HTML não encontrado após renderização. Ficheiros HTML: ",
               if (length(html_files)) paste(html_files, collapse = ", ") else "(nenhum)")
        }

        dir.create(dirname(report_dest), showWarnings = FALSE, recursive = TRUE)
        file.copy(tmp_html, report_dest, overwrite = TRUE)
        message("✓ Relatório gerado.")
      },
      args   = list(qmd = qmd_abs, abs_parquet = abs_parquet, report_dest = rep_dest),
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
        { job$get_result(); list(ok = TRUE, err = NULL) },
        error = function(e) list(ok = FALSE, err = conditionMessage(e))
      )
      rep_ok(result$ok)
      rep_err(result$err)
      rep_done(TRUE)
    }
  })

  observeEvent(rep_done(), {
    req(rep_done())
    if (rep_ok()) {
      rep_url(rep_url_path)
      session$sendCustomMessage("open_report", rep_url_path)
      showNotification(
        "Relatório gerado. Se não abrir automaticamente, use o link no painel.",
        type = "message", duration = 10
      )
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
    url  <- rep_url()
    card(card_body(
      if (!done)
        div(class = "d-flex align-items-center gap-2",
          div(class = "spinner-border spinner-border-sm text-success"),
          span("A renderizar relatório..."))
      else if (ok)
        div(class = "alert alert-success mb-0",
          strong(icon("circle-check"), " Relatório gerado. "),
          if (!is.null(url))
            tags$a(href = url, target = "_blank", class = "alert-link",
                   "Abrir numa nova janela."),
          " ",
          downloadButton("dl_report", "Descarregar HTML",
                         class = "btn-outline-secondary btn-sm ms-2")
        )
      else
        div(class = "alert alert-danger mb-0",
          tags$b("Erro: "),
          if (!is.null(err)) err else "Erro desconhecido. Ver log abaixo.")
    ))
  })

  output$dl_report <- downloadHandler(
    filename = "relatorio-execucao-orcamental.html",
    content  = function(file) file.copy(rep_dest, file)
  )
}

shinyApp(ui, server)
