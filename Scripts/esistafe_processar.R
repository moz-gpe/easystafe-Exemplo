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

stopifnot(
  "UGB file not found" = file.exists(path_ugb_lookup),
  "Data folder not found" = dir.exists("Data/")
)


# LOAD LOOKUPS ------------------------------------------------------------

lookup_ugb <- read_excel(metadata_lookup,
                         sheet = "Sheet1") %>%
  clean_names() %>%
  select(codigo_ugb,
         provincia,
         distrito,
         ambito,
         starts_with("adm"),
         nivel_da_instituicao,
         descricao) %>%
  filter(!codigo_ugb == "Total")


lookup_funcao <- read_excel(metadata_lookup,
                            sheet = "Sheet2") %>%
  clean_names() %>%
  select(funcao,
         funcao_nivel =classificacao_funcional_por_nivel) %>%
  filter(!is.na(funcao))


lookup_programa <- read_excel(metadata_lookup,
                              sheet = "Sheet2") %>%
  clean_names() %>%
  select(programa_esistafe = programa_e_sistafe,
         programa_educacao) %>%
  filter(!is.na(programa_esistafe))


# PROCESSAMENTO E-SISTAFE -----------------------------------------------------

df_esistafe <- processar_extracto_esistafe(
  df_ugb_lookup  = lookup_ugb,
  include_percent = FALSE,
  include_file_metadata = TRUE,
  include_metrica = TRUE,
  quiet = TRUE) %>%
  left_join(lookup_ugb, by = join_by(ugb_id == codigo_ugb)) %>%
  left_join(lookup_funcao, by = join_by(funcao == funcao)) %>%
  left_join(lookup_programa, by = join_by(programa == programa_esistafe)) %>%
  recode_programa_tipo() %>%
  relocate(funcao_nivel, .after = funcao) %>%
  relocate(programa_educacao, .after = programa) %>%
  relocate(provincia, distrito, ambito, adm2020_24, adm2025_29,
           nivel_da_instituicao, descricao, programa_tipo,
           .after = ced)


funcao_na <- df_esistafe %>%
  filter(is.na(funcao_nivel)) %>%
  distinct(funcao, funcao_nivel)


# PROCESSAMENTO RAZAO CONT. & ABSA ---------------------------------------------------------

path_folder_source <- "Data/razao_cont/2026_02/"

df_razao <- processar_extracto_razao_c(source_path = path_folder_source)
df_absa <- processar_extracto_absa(path_folder_source)


# GRAVAR -----------------------------------------------------------------

gravar_extracto_sistafe(df_esistafe)
gravar_extracto_razao_c(df_razao)
gravar_extracto_absa(df_absa)
