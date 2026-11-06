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
options(shiny.maxRequestSize = 500 * 1024^2)  # 500 MB total upload per request

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

  tags$head(tags$script(HTML("
    Shiny.addCustomMessageHandler('open_report', function(url){ window.open(url, '_blank'); });

    // Cumulative folder picker — each click opens a folder picker and
    // APPENDS the new files to the existing selection via DataTransfer.
    // After each addition the accumulated FileList is assigned to Shiny's
    // hidden input and a change event is dispatched so Shiny uploads them.
    var accFiles = new DataTransfer();

    function pushToShiny() {
      var si = document.getElementById('files');
      if (!si) return;
      si.files = accFiles.files;
      si.dispatchEvent(new Event('change', { bubbles: true }));
      var n = accFiles.files.length;
      document.getElementById('files_count').textContent =
        n > 0 ? n + ' ficheiro(s) seleccionado(s)' : '';
    }

    $(document).on('click', '#btn_add_folder', function() {
      var tmp = document.createElement('input');
      tmp.type = 'file';
      tmp.setAttribute('webkitdirectory', '');
      tmp.setAttribute('multiple', '');
      tmp.style.cssText = 'position:fixed;top:-9999px;left:-9999px;';
      document.body.appendChild(tmp);
      tmp.addEventListener('change', function() {
        for (var i = 0; i < tmp.files.length; i++) accFiles.items.add(tmp.files[i]);
        document.body.removeChild(tmp);
        pushToShiny();
      });
      tmp.addEventListener('cancel', function() { document.body.removeChild(tmp); });
      tmp.click();
    });

    $(document).on('click', '#btn_clear_files', function(e) {
      e.preventDefault();
      accFiles = new DataTransfer();
      pushToShiny();
    });
  "))),

  sidebar = sidebar(
    width = 270,

    tags$b("Ficheiros e-SISTAFE"),
    tags$button(
      id    = "btn_add_folder",
      class = "btn btn-default w-100 mt-1 mb-1",
      tags$i(class = "fa fa-folder-open me-1"), "Adicionar pasta..."
    ),
    tags$small(id = "files_count", class = "text-muted d-block"),
    uiOutput("periods_ui"),
    tags$a(id = "btn_clear_files", href = "#",
           class = "small text-danger d-block mt-1 mb-1",
           "Limpar selecção"),

    # Hidden Shiny binding target — the visible picker above feeds into this.
    tags$div(
      style = "display:none",
      tags$input(id = "files", class = "shiny-input-file", name = "files",
                 type = "file", multiple = "multiple"),
      tags$div(id = "files_progress", class = "progress shiny-file-input-progress",
               tags$div(class = "progress-bar"))
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
        uiOutput("sitrep_out")
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
        suppressWarnings(suppressPackageStartupMessages({
          library(tidyverse)
          library(readxl)
          library(writexl)
          library(janitor)
          library(arrow)
          library(easystafe)
        }))

        # Redirect both stdout and stderr/message connections to null.
        # cli can write to either depending on context, so we must sink both.
        quietly <- function(expr) {
          out_con <- file(nullfile(), "w")
          msg_con <- file(nullfile(), "w")
          sink(out_con, type = "output")
          sink(msg_con, type = "message")
          on.exit({
            sink(type = "output")
            sink(type = "message")
            close(out_con)
            close(msg_con)
          }, add = TRUE)
          force(expr)
        }

        message("[1/4] A carregar lookups...")
        lookups <- quietly(carregar_lookups_esistafe(meta))

        message("[2/4] A processar ", length(paths), " período(s)...")
        df_raw <- map(paths, \(p) {
          message("  → ", basename(p), "...")
          quietly(processar_extracto_esistafe(
            source_path           = p,
            df_ugb_lookup         = lookups$ugb,
            include_percent       = FALSE,
            include_file_metadata = TRUE,
            include_metrica       = TRUE,
            correct_negatives     = TRUE,
            quiet                 = TRUE
          ))
        }) |> list_rbind()
        message("  Total: ", nrow(df_raw), " linhas")

        df <- quietly(
          df_raw |>
            adicionar_lookups_esistafe(lookups) |>
            escolaR::padronizar_geo_dim(col_provincia = "provincia", col_distrito = "distrito")
        )

        message("[3/4] A verificar qualidade (sitrep)...")
        sitrep_dir <- file.path(outdir, "sitrep")
        gaps <- quietly(sitrep_processamento_esistafe(
          df, lookups,
          lookup_path   = meta,
          output_folder = sitrep_dir,
          quiet         = TRUE
        ))
        sitrep_excel <- {
          p <- file.path(sitrep_dir,
                         paste0("sitrep_esistafe_", format(Sys.Date(), "%Y-%m-%d"), ".xlsx"))
          if (file.exists(p)) p else NULL
        }

        message("[4/4] A gravar ficheiros...")
        if (!dir.exists(outdir)) dir.create(outdir, recursive = TRUE)
        quietly(arrow::write_parquet(df, file.path(outdir, "e-SISTAFE_processado.parquet")))
        quietly(writexl::write_xlsx(df, file.path(outdir, "e-SISTAFE_processado.xlsx")))

        message("✓ Processamento concluído.")
        list(gaps = gaps, nrows = nrow(df), sitrep_excel = sitrep_excel)
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
    invalidateLater(400)
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

  output$sitrep_out <- renderUI({
    req(proc_done())
    res <- proc_res()
    if (is.null(res) || !is.null(res$error)) return(NULL)

    gaps <- res$gaps
    n    <- gaps$n_afectados

    fmt_mt <- function(x) {
      if (is.na(x) || x == 0) return("MT 0")
      if (x >= 1e9) paste0("MT ", round(x / 1e9, 1), "B")
      else           paste0("MT ", round(x / 1e6, 0), "M")
    }

    get_stats <- function(tipo) {
      r <- n[n$gap_tipo == tipo, ]
      list(n_gap = r$n_gap, n_linhas = r$n_linhas, dotacao = r$dotacao_total)
    }

    fa   <- get_stats("funcao_nivel")
    pb   <- get_stats("programa_tipo")
    ced  <- get_stats("ced_nome")
    ugb_n <- nrow(gaps$ugb_completude)
    total <- fa$n_gap + pb$n_gap + ced$n_gap + ugb_n

    gap_section <- function(letter, title, n_gap, impact, description, followup,
                            last = FALSE) {
      ok <- n_gap == 0
      icon_cls <- if (ok) "fa fa-circle-check text-success mt-1"
                  else    "fa fa-triangle-exclamation text-warning mt-1"
      div(class = if (last) "mb-2" else "mb-3 pb-3 border-bottom",
        div(class = "d-flex gap-2",
          tags$i(class = icon_cls, style = "min-width:16px"),
          div(class = "w-100",
            tags$b(paste0("[", letter, "] ", title)),
            if (ok)
              p(class = "small text-success mb-0 mt-1", "Sem lacunas detectadas.")
            else
              tagList(
                p(class = "small text-muted mt-1 mb-1", description),
                div(class = "mb-1",
                  tags$span(class = "badge text-bg-warning", impact)),
                div(class = "small mb-0", followup)
              )
          )
        )
      )
    }

    n_fmt <- function(x) format(x, big.mark = " ")

    sec_A <- gap_section(
      "A", "Classificação de Funções",
      fa$n_gap,
      paste0(fa$n_gap,
             if (fa$n_gap == 1) " função" else " funções",
             " em falta → ", n_fmt(fa$n_linhas), " linhas → ", fmt_mt(fa$dotacao)),
      paste0(fa$n_gap, " código(s) de função nos dados e-SISTAFE sem classificação de ",
             "nível no ficheiro de metadados (esistafe_dim_lookup.xlsx → folha 'funcao')."),
      tagList("Adicionar os ", fa$n_gap, " código(s) em falta à folha ",
              tags$b("funcao"), " do ficheiro de metadados.")
    )

    sec_B <- gap_section(
      "B", "Tipo de Programa",
      pb$n_gap,
      paste0(pb$n_gap,
             if (pb$n_gap == 1) " combinação" else " combinações",
             " em falta → ", n_fmt(pb$n_linhas), " linhas → ", fmt_mt(pb$dotacao)),
      paste0(pb$n_gap, " combinação(ões) programa-âmbito-FR nos dados e-SISTAFE sem ",
             "tipo de programa correspondente no ficheiro de metadados ",
             "(esistafe_dim_lookup.xlsx → folha 'programa')."),
      tagList("Adicionar as ", pb$n_gap, " combinação(ões) em falta à folha ",
              tags$b("programa"), " do ficheiro de metadados.")
    )

    sec_C <- gap_section(
      "C", "Nomes de CED (Classificação Económica de Despesa)",
      ced$n_gap,
      paste0(ced$n_gap,
             if (ced$n_gap == 1) " código CED" else " códigos CED",
             " em falta → ", n_fmt(ced$n_linhas), " linhas → ", fmt_mt(ced$dotacao)),
      paste0(ced$n_gap, " código(s) CED nos dados e-SISTAFE sem nome correspondente ",
             "no ficheiro de metadados (esistafe_dim_lookup.xlsx)."),
      tagList(
        "Actualizar ", tags$b("em cascata"), " por cada código em falta:",
        tags$ol(class = "small mb-1 mt-1 ps-3",
          tags$li(tagList("Folha ", tags$b("ced"), " — sempre")),
          tags$li(tagList("Folha ", tags$b("ced_3"),
                          " — se ", tags$code("ced_3_existe = FALSE"))),
          tags$li(tagList("Folha ", tags$b("ced_nivel"),
                          " — se ced_3 é novo ou ", tags$code("ced_nivel_existe = FALSE"))),
          tags$li(tagList("Folha ", tags$b("ced_2"),
                          " — se ", tags$code("ced_2_existe = FALSE")))
        ),
        tags$small(class = "text-muted",
          "O Excel de lacunas indica os passos necessários para cada código.")
      )
    )

    sec_D <- gap_section(
      "D", "Completude de UGBs (Funcionamento)",
      ugb_n,
      paste0(ugb_n,
             if (ugb_n == 1) " combinação período × UGB"
             else             " combinações período × UGB",
             " sem dados de Funcionamento"),
      paste0(ugb_n, " combinação(ões) período × UGB esperada(s) segundo o ficheiro ",
             "de metadados não estão presentes nos dados e-SISTAFE carregados."),
      "Verificar se todos os relatórios mensais de Funcionamento foram incluídos. Se a ausência for esperada (UGB sem actividade nesse período), pode ignorar.",
      last = TRUE
    )

    tagList(
      if (total == 0)
        div(class = "d-flex gap-2 align-items-center text-success mb-3",
          tags$i(class = "fa fa-circle-check"),
          tags$b("Sem lacunas detectadas — todos os lookups estão completos."))
      else
        p(class = "text-muted small mb-3",
          paste0(total, " lacuna(s) detectada(s) — ver detalhes abaixo.")),
      sec_A, sec_B, sec_C, sec_D,
      if (total > 0 && !is.null(res$sitrep_excel) && file.exists(res$sitrep_excel))
        div(class = "mt-3",
          downloadButton("dl_sitrep", "Descarregar Excel de lacunas",
                         class = "btn-outline-warning btn-sm"))
    )
  })

  output$dl_sitrep <- downloadHandler(
    filename = function() paste0("sitrep_esistafe_", format(Sys.Date(), "%Y-%m-%d"), ".xlsx"),
    content  = function(file) {
      path <- proc_res()$sitrep_excel
      if (!is.null(path) && file.exists(path)) file.copy(path, file)
    }
  )

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
    invalidateLater(400)
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
