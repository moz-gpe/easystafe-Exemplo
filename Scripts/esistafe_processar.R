pak::pak("moz-gpe/easystafe")

library(tidyverse)
library(readxl)
library(writexl)
library(janitor)
library(glue)
library(easystafe)


# GLOBAL VARS -------------------------------------------------------------

path_data_folder <- "Data/"
path_ugb_file <- "Documents/Codigos de UGBs.xlsx"
path_ugb_lookup <- "Documents/ugb_lookup.xlsx"


stopifnot(
  "UGB file not found — check path_ugb_file" = file.exists(path_ugb_file),
  "Data folder not found — check path_data_folder" = dir.exists(path_data_folder)
)

path_files <- list.files(
  path = path_data_folder,
  pattern = "\\.xls$",
  full.names = TRUE,
  ignore.case = TRUE
)


# LOAD LOOKUPS ------------------------------------------------------------

ugb_raw <- read_excel(path_ugb_file, sheet = "UGBS")
ugb_lookup <- read_excel(path_ugb_lookup) %>% clean_names() %>% filter_out(codigo_ugb == "Total") %>% select(codigo_ugb, ambito, provincia, distrito, descricao)


# PROCESSAR FICHEIROS -----------------------------------------------------

df <- processar_extracto_esistafe(
  source_path = path_files,
  ugb_lookup  = ugb_raw
)

df_final <- df %>%
  left_join(ugb_lookup, by = join_by(ugb_id == codigo_ugb)) %>%
  recode_programa_tipo()


# CHECK RECODING ----------------------------------------------------------

df_final %>% distinct(programa_tipo) %>% print(n = Inf)


# ESCREVER A DISCO --------------------------------------------------------

# Default output folder
gravar_extracto_sistafe(df_final)


