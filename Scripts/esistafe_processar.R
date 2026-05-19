
# DOCUMENTACAO DO PACOTE 'easystafe' DISPONIVEL ABAIXO
# https://github.com/moz-gpe/easystafe

# CORRER CODIGOS ABAIXO PARA ACTUALIZAR O PACOTE
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

metadata_lookup <- "Documents/lookup.xlsx"

paths_esistafe <- c(
  "Data/202503",
  "Data/202504",
  "Data/202601",
  "Data/202602",
  "Data/202603",
  "Data/202604"
)

#path_razao_contabalistica <- "Data/razao_cont/2026_01/"


# LOAD LOOKUPS ------------------------------------------------------------

lookups <- list(

  ugb = suppressMessages(
    read_excel(metadata_lookup, sheet = "ugb")) %>%
    clean_names() %>%
    select(codigo_ugb,
           provincia,
           distrito,
           ambito,
           starts_with("adm"),
           nivel_da_instituicao,
           descricao) %>%
    filter(!codigo_ugb == "Total"),

  funcao = suppressMessages(
    read_excel(metadata_lookup, sheet = "funcao")) %>%
    clean_names() %>%
    select(funcao,
           funcao_nivel = classificacao_funcional_por_nivel) %>%
    filter(!is.na(funcao)),

  programa = suppressMessages(
    read_excel(metadata_lookup, sheet = "programa")) %>%
    clean_names() %>%
    select(programa,
           programa_tipo) %>%
    filter(!is.na(programa_tipo))
)


# PROCESSAR E GRAVAR E-SISTAFE -----------------------------------------------------

df_esistafe <- paths_esistafe %>%
  map(\(path) processar_extracto_esistafe(
    source_path = path,
    df_ugb_lookup          = lookups$ugb,
    include_percent        = FALSE,
    include_file_metadata  = TRUE,
    include_metrica        = TRUE,
    correct_negatives      = TRUE,
    quiet                  = FALSE
  ) %>%
    mutate(conjunto = basename(path))
  ) %>%
  bind_rows() %>%
  left_join(lookups$ugb, by = join_by(ugb_id == codigo_ugb)) %>%
  left_join(lookups$funcao, by = join_by(funcao == funcao)) %>%
  left_join(lookups$programa, by = join_by(programa == programa)) %>%
  relocate(funcao_nivel, .after = funcao) %>%
  relocate(provincia, distrito, ambito, adm2020_24, adm2025_29,
           nivel_da_instituicao, descricao, programa_tipo,
           .after = ced) |>
  relocate(conjunto, .before = everything())

gravar_esistafe(df_esistafe,
                output_folder = "Dataout/",
                quiet = TRUE)


# PROCESSAR E GRAVAR RAZAO CONTABILISTICA. & ABSA ---------------------------------------------------------

df_razao <- processar_extracto_razao_c(source_path = path_razao_contabalistica)
df_absa <- processar_extracto_absa(source_path = path_razao_contabalistica)

df_razao <- bind_rows(df_razao, df_absa)

gravar_extracto_razao_c(df_razao)

rm(df_absa)

# COMPILAR FICHEIROS ------------------------------------------------------


path_razao <- gravar_compilacao_razao_c()
read_xlsx(path_razao) |>
  write_parquet(sub("\\.xlsx$", ".parquet", path_razao),
                compression = "zstd", compression_level = 9)

