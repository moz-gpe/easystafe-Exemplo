# pak::pak("moz-gpe/easystafe")
# https://github.com/moz-gpe/easystafe

library(tidyverse)
library(readxl)
library(writexl)
library(janitor)
library(glue)
library(easystafe)

# GLOBAL VARS -------------------------------------------------------------

path_data_folder <- "Data/"
path_ugb_lookup <- "Documents/OrganicaEducação.xlsx"


stopifnot(
  "UGB file not found — check path_ugb_lookup" = file.exists(path_ugb_lookup),
  "Data folder not found — check path_data_folder" = dir.exists(path_data_folder)
)

path_files <- list.files(
  path = path_data_folder,
  pattern = "\\.xls$",
  full.names = TRUE,
  ignore.case = TRUE
)


# LOAD LOOKUPS ------------------------------------------------------------

ugb_lookup <- read_excel(path_ugb_lookup,
                         sheet = "Sheet1") %>%
  clean_names() %>%
  select(-c(starts_with("nome_"), codigo_provincia)) %>%
  filter(!codigo_ugb == "Total")


# PROCESSAR FICHEIROS -----------------------------------------------------

df <- processar_extracto_esistafe(
  source_path = path_files,
  df_ugb_lookup  = ugb_lookup,
  include_percent = TRUE,
  include_file_metadata = TRUE,
  include_metrica = TRUE,
  quiet = TRUE
)

df_final <- df %>%
  left_join(ugb_lookup, by = join_by(ugb_id == codigo_ugb)) %>%
  recode_programa_tipo()


t <- df_final %>%
  separate(col = programa,
           into = c("programa_id", "programa_nome"),
           sep = " - ")
t %>%
  filter_out(programa_tipo == "Outro") %>%
  distinct(programa_id)

# VERIFICAR RECODIFICACAO ----------------------------------------------------------

df_final %>% distinct(programa_tipo) %>% print(n = Inf)


# ESCREVER A DISCO --------------------------------------------------------

# Default output folder
gravar_extracto_sistafe(df_final)


