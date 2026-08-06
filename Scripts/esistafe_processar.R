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
library(pins)
library(easystafe)

# GLOBAL VARS -------------------------------------------------

metadata_lookup <- "Documents/lookup.xlsx"

paths_esistafe <- c(
  "Data/demonstrativo_consolidado/202502",
  "Data/demonstrativo_consolidado/202503",
  "Data/demonstrativo_consolidado/202504",
  "Data/demonstrativo_consolidado/202505",
  "Data/demonstrativo_consolidado/202512",
  "Data/demonstrativo_consolidado/202601",
  "Data/demonstrativo_consolidado/202602",
  "Data/demonstrativo_consolidado/202603",
  "Data/demonstrativo_consolidado/202604",
  "Data/demonstrativo_consolidado/202605"
)

board <- board_folder(
  "C:/Users/Joe/Documents/R/moz-data-pins/Data",
  versioned = TRUE
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
  adicionar_lookups_esistafe(lookups)

# CONFIGURAR PARA DUCKDB -----------------------------------------------

df_esistafe_db <- df_esistafe |>
  config_para_duckdb()

# GRAVAR DADOS PROCESSADOS -----------------------------------------------

gravar_esistafe(df_esistafe, output_folder = "Dataout/", quiet = TRUE)

pin_write(
  board = board,
  x = df_esistafe,
  name = "demostrativo_consolidado_compilado",
  type = "parquet",
  description = "Compilação do Demonstrativo Consolidado gerado pelo pipeline easystafe."
)

pin_write(
  board = board,
  x = df_esistafe_db,
  name = "esistafe_dm_duckdb",
  type = "parquet",
  description = "Compilação do Demonstrativo Consolidado gerado pelo pipeline easystafe."
)

# GERAR RELATORIO --------------------------------------------------------

quarto::quarto_render(
  input = "Relatório de Execução Orçamental.qmd"
)

rstudioapi::viewer("Relatório de Execução Orçamental.html")
