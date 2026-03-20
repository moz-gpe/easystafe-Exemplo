library(tidyverse)
library(readxl)
library(writexl)
library(janitor)
library(glue)
library(gt)
#library(easystafe)
source("Scripts/functions.R")



path_ugb_file <- "../../Data/ugb/Codigos de UGBs.xlsx"
path_data_folder <- "Data/debugging/"

stopifnot(
  "UGB file not found — check path_ugb_file" = file.exists(path_ugb_file),
  "Data folder not found — check path_data_folder" = dir.exists(path_data_folder)
)

path_files <- list.files(
  path = path_data_folder,
  pattern = "\\.xlsx$",
  full.names = TRUE,
  ignore.case = TRUE
)


#paths_meta <- extrair_meta_extracto(path_files)

ugb_raw <- read_excel(path_ugb_file, sheet = "UGBS")

# Default — with metadata, with percent columns
df <- processar_extracto_sistafe(
  source_path = path_files,
  ugb_lookup  = ugb_raw
)

# With metadata, without percent columns
df <- processar_extracto_sistafe(
  source_path     = path_files,
  ugb_lookup      = ugb_raw,
  include_percent = FALSE
)

# Without metadata (drops file_name too), with percent columns
df <- processar_extracto_sistafe(
  source_path  = path_files,
  ugb_lookup   = ugb_raw,
  include_meta = FALSE
)

# Without metadata, without percent columns — leanest possible output
df <- processar_extracto_sistafe(
  source_path     = path_files,
  ugb_lookup      = ugb_raw,
  include_percent = FALSE,
  include_meta    = FALSE
)


# Default output folder
gravar_extracto_sistafe(df)

# Custom output folder
gravar_extracto_sistafe(df, output_folder = "Data/processed")

# With verbose output
gravar_extracto_sistafe(df, quiet = FALSE)
