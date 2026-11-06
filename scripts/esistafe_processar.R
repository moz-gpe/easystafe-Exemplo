# DOCUMENTACAO DO PACOTE 'easystafe'
# https://github.com/moz-gpe/easystafe

# ACTUALIZAR O PACOTE
# pak::pak("moz-gpe/easystafe")

# CORRER CODIGOS ABAIXO LINHA POR LINHA NA JANELA 'TERMINAL' PARA ALINHAR SCRIPT COM GITHUB
# git fetch origin
# git reset --hard origin/main
# git clean -fd

library(tidyverse)
library(readxl)
library(writexl)
library(janitor)
library(glue)
library(arrow)
library(easystafe)

# GLOBAL VARS -------------------------------------------------

metadata_lookup <- "documents/esistafe_dim_lookup.xlsx"

paths_esistafe <- c(
  # "data/demonstrativo_consolidado/202502",
  # "data/demonstrativo_consolidado/202503",
  # "data/demonstrativo_consolidado/202504",
  # "data/demonstrativo_consolidado/202505",
  # "data/demonstrativo_consolidado/202506",
  # "data/demonstrativo_consolidado/202507",
  # "data/demonstrativo_consolidado/202512",
  # "data/demonstrativo_consolidado/202601",
  # "data/demonstrativo_consolidado/202602",
  # "data/demonstrativo_consolidado/202603",
  # "data/demonstrativo_consolidado/202604",
  # "data/demonstrativo_consolidado/202605",
  # "data/demonstrativo_consolidado/202606",
  # "data/demonstrativo_consolidado/202607"
)


# LOAD LOOKUPS ------------------------------------------------------------

lookups <- carregar_lookups_esistafe(metadata_lookup)

# PROCESSAR E-SISTAFE -----------------------------------------------------

df_esistafe <- paths_esistafe |>
  map(\(path) {
    processar_extracto_esistafe(
      source_path = path,
      df_ugb_lookup = lookups$ugb,
      include_percent = FALSE,
      include_file_metadata = TRUE,
      include_metrica = TRUE,
      correct_negatives = TRUE,
      quiet = FALSE
    )
  }) |>
  list_rbind() |>
  adicionar_lookups_esistafe(lookups) |>
  escolaR::padronizar_geo_dim(col_provincia = "provincia", col_distrito = "distrito")

gaps <- sitrep_processamento_esistafe(
  df_esistafe,
  lookups,
  lookup_path = metadata_lookup
)

# GRAVAR DADOS PROCESSADOS -----------------------------------------------

arrow::write_parquet(df_esistafe, "dataout/e-SISTAFE_processado.parquet")
writexl::write_xlsx(df_esistafe, "dataout/e-SISTAFE_processado.xlsx")

# GERAR RELATORIO --------------------------------------------------------

quarto::quarto_render(
  input = "Relatório de Execução Orçamental.qmd"
)

rstudioapi::viewer("Relatório de Execução Orçamental.html")
