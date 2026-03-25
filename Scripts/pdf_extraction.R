library(tidyverse)
library(pdftools)
library(openxlsx)

library(easystafe)
#source("Scripts/functions.R")


path_folder_source <- "~/Data/esistafe/razao_contabilistico/"


df_razao <- processar_extracto_razao_c(
  source_path = path_folder_source
)


gravar_extracto_razao_c(df_razao)


