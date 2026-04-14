# pak::pak("moz-gpe/easystafe")
# https://github.com/moz-gpe/easystafe

library(tidyverse)
library(readxl)
library(writexl)
library(janitor)
library(glue)
library(easystafe)

# GLOBAL VARS -------------------------------------------------

metadata_lookup <- "Documents/lookup_ugb.xlsx"


# LOAD LOOKUPS ------------------------------------------------------------


lookup_ugb <- suppressMessages(
  read_excel(metadata_lookup, sheet = "Sheet1")) %>%
  clean_names() %>%
  select(codigo_ugb,
         provincia,
         distrito,
         ambito,
         starts_with("adm"),
         nivel_da_instituicao,
         descricao) %>%
  filter(!codigo_ugb == "Total")


lookup_funcao <- suppressMessages(
  read_excel(metadata_lookup, sheet = "Sheet2")) %>%
  clean_names() %>%
  select(funcao,
         funcao_nivel = classificacao_funcional_por_nivel) %>%
  filter(!is.na(funcao))


lookup_programa <- suppressMessages(
  read_excel(metadata_lookup, sheet = "Sheet2")) %>%
  clean_names() %>%
  select(programa_esistafe = programa_e_sistafe,
         programa_educacao) %>%
  filter(!is.na(programa_esistafe))


# PROCESSAMENTO E-SISTAFE -----------------------------------------------------

df_esistafe <- processar_extracto_esistafe(
  source_path = "Data/2026/",
  df_ugb_lookup  = lookup_ugb,
  include_percent = FALSE,
  include_file_metadata = TRUE,
  include_metrica = TRUE,
  quiet = FALSE) %>%
  left_join(lookup_ugb, by = join_by(ugb_id == codigo_ugb)) %>%
  left_join(lookup_funcao, by = join_by(funcao == funcao)) %>%
  recode_programa_tipo() %>%
  relocate(funcao_nivel, .after = funcao) %>%
  relocate(provincia, distrito, ambito, adm2020_24, adm2025_29,
           nivel_da_instituicao, descricao, programa_tipo,
           .after = ced)


# PROCESSAMENTO RAZAO CONT. & ABSA ---------------------------------------------------------

path_folder_source <- "Data/razao_cont/2026_02/"

df_razao <- processar_extracto_razao_c(source_path = path_folder_source)
df_absa <- processar_extracto_absa(path_folder_source)

df_razao <- bind_rows(df_razao, df_absa)

# GRAVAR FICHEIRO DO PERIODO -----------------------------------------------------------------

gravar_extracto_sistafe(df_esistafe)
gravar_extracto_razao_c(df_razao)


rm(df_absa,
   lookup_funcao,
   lookup_programa,
   lookup_ugb)


# COMPILAR FICHEIROS ------------------------------------------------------

gravar_compilacao_sistafe()
gravar_compilacao_razao_c()
