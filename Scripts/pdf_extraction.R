library(tidyverse)
library(pdftools)
library(openxlsx)
library(easystafe)
#source("Scripts/functions.R")


path_folder_source <- "Data/razao_cont/2026_02/"


df_razao <- processar_extracto_razao_c(
  source_path = path_folder_source
)


gravar_extracto_razao_c(df_razao)




# TESTING -----------------------------------------------------------------

path_folder_source_absa <- "Data/razao_cont/2026_02/outro/"

df_absa <- processar_extracto_absa(path_folder_source_absa)


gravar_extracto_absa(df_absa)
