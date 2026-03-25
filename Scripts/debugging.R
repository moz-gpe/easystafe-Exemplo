library(tidyverse)
library(readxl)
library(writexl)
library(janitor)
library(glue)
library(gt)
library(easystafe)
#source("Scripts/functions.R")


path_ugb_file <- "../../Data/ugb/Codigos de UGBs.xlsx"
path_data_folder <- "../../Data/esistafe/essistafe_dez25/"
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


#paths_meta <- extrair_meta_extracto(path_files)

ugb_raw <- read_excel(path_ugb_file, sheet = "UGBS")
ugb_lookup <- read_excel(path_ugb_lookup) %>% clean_names() %>% filter_out(codigo_ugb == "Total") %>% select(codigo_ugb, ambito, provincia, distrito, descricao)

# Default — with metadata, with percent columns
df <- processar_extracto_esistafe(
  source_path = path_files,
  ugb_lookup  = ugb_raw
)


df_final <- df %>%
  left_join(ugb_lookup, by = join_by(ugb_id == codigo_ugb)) %>%
  mutate(
    distrito = distrito %>%
      str_to_title() %>%
      str_replace_all("\\b(De|Da|Do|Dos|Das)\\b", ~ str_to_lower(.x))
  ) %>%
  mutate(
    programa_tipo = case_when(

      # ADE - ESG
      str_detect(programa, regex("Apoiar Escolas Secund", ignore_case = TRUE)) |
        str_detect(programa, regex("ADE\\s*-\\s*ESG", ignore_case = TRUE)) |
        str_detect(programa, regex("ADE \\* ESG", ignore_case = TRUE)) |
        str_detect(programa, regex("ESG\\s*-\\s*ADE", ignore_case = TRUE)) |
        str_detect(programa, regex("ESG1\\s*-\\s*1\\s*CICLO", ignore_case = TRUE)) |
        str_detect(programa, regex("ESG\\s*-\\s*I\\s*CICLO", ignore_case = TRUE)) |
        str_detect(programa, regex("Ensino Secundário Geral 1 Ciclo", ignore_case = TRUE)) |
        str_detect(programa, regex("ADE Ensino Secundario", ignore_case = TRUE)) |
        str_detect(programa, regex("APOIO DIRECTO PARA AS ESCOLAS SECUNDARIAS", ignore_case = TRUE)) |
        str_detect(programa, regex("Apoiar as Escolas Secundarias atraves do Fundo de Apoio Directo as Escolas", ignore_case = TRUE)) |
        str_detect(programa, regex("APOIAR AS ESCOLAS SECUNDARIAS, ATRAVES DO FUNDO DE APOIO DIRECTO AS ESCOLAS", ignore_case = TRUE)) |
        str_detect(programa, regex("APOIO DIRECTO AS ESCOLAS ESG", ignore_case = TRUE)) ~ "ADE - ESG",

      # ADE - Basica
      str_detect(programa, regex("Apoiar as escolas basicas, atraves do fundo de apoio directo as escolas", ignore_case = TRUE)) |
        str_detect(programa, regex("Apoiar Escolas Básicas através do Fundo do Apoio Directo as Escolas", ignore_case = TRUE)) |
        str_detect(programa, regex("Apoiar Escolas Básicas através do Fundo  do Apoio Directo as Escolas", ignore_case = TRUE)) |
        str_detect(programa, regex("Apoiar Escolas Básicas através do Fundo de Apoio Directo as Escolas", ignore_case = TRUE)) |
        str_detect(programa, regex("Apoiar Escolas Básicas através de Fundo de Apoio directo as Escolas", ignore_case = TRUE)) |
        str_detect(programa, regex("Apoiar Escolas Basicas atraves de Fundo de Apoio Directo as Escolas", ignore_case = TRUE)) |
        str_detect(programa, regex("Apoiar Escolas Básicas Através de Fundos de Apoio Directo as Escolas", ignore_case = TRUE)) |
        str_detect(programa, regex("APOIAR AS ESCOLAS BASICAS ATRAVES DO FUNDO DE APOIO DIRECTO AS ESCOLAS", ignore_case = TRUE)) ~ "ADE - Basica",

      # ADE Primária
      str_detect(programa, regex("APOIO DIRECTO AS ESCOLA \\(ADE\\)", ignore_case = TRUE)) |
        str_detect(programa, regex("APOIAR AS ESCOLAS PRIMARIAS ATRAVES DO FUNDO DIRECTO AS ESCOLAS", ignore_case = TRUE)) |
        str_detect(programa, regex("APOIO DIRECTO AS ESCOLAS", ignore_case = TRUE)) |
        str_detect(programa, regex("APOIO DIRECTO PARA AS ESCOLAS PRIMARIAS \\(ADE\\)", ignore_case = TRUE)) |
        str_detect(programa, regex("APOIO DIRECTO AS ESCOLAS PRIM]ARIAS \\(ADE\\)", ignore_case = TRUE)) |
        str_detect(programa, regex("APOIO DIRECTO A ESCOLAS", ignore_case = TRUE)) |
        str_detect(programa, regex("Apoiar as Escolas Primarias através do Fundo Directo as Escolas", ignore_case = TRUE)) |
        str_detect(programa, regex("Apoiar as Escolas Primarias através do Fundo de apoio as Escolas", ignore_case = TRUE)) |
        str_detect(programa, regex("APOIO DIRECTO  AS ESCOLAS PRIMARIAS \\(ADE\\)", ignore_case = TRUE)) ~ "ADE Primária",

      # Supervisão Distrital
      str_detect(programa, regex("SUPERVISAO DISTRITAL", ignore_case = TRUE)) |
        str_detect(programa, regex("Supervisão Distrital", ignore_case = TRUE)) |
        str_detect(programa, regex("Supervisao das Escolas", ignore_case = TRUE)) ~ "Supervisão Distrital",

      # Supervisão Provincial
      str_detect(programa, regex("SUPERVISAO PROVINCI", ignore_case = TRUE)) |
        str_detect(programa, regex("Supervisão Provincial", ignore_case = TRUE)) |
        str_detect(programa, regex("PROVINCIA PARA A REALIZACAO DA SUPERVISAO", ignore_case = TRUE)) ~ "Supervisão Provincial",

      # Controlo Interno
      str_detect(programa, regex("Control", ignore_case = TRUE)) ~ "Controlo Interno da Educação",

      # Professores
      str_detect(programa, regex("Profess", ignore_case = TRUE)) ~ "Capacitação e Formação de Professores",

      # Primeira Infância
      str_detect(programa, regex("Inf", ignore_case = TRUE)) |
        str_detect(programa, regex("Facilita", ignore_case = TRUE)) |
        str_detect(programa, regex("Pilot", ignore_case = TRUE)) ~ "Projecto Piloto da Primeira Infância",

      # HIV
      str_detect(programa, regex("HIV", ignore_case = TRUE)) ~ "Prevenção e Combate do HIV/SIDA",

      # Livro
      str_detect(programa, regex("Livro", ignore_case = TRUE)) ~ "Livro Escolar",

      # Material Informático
      str_detect(programa, regex("Adqui", ignore_case = TRUE)) ~ "Adquirir Material Informáticos",

      # Construção ESG
      str_detect(programa, regex("Constr", ignore_case = TRUE)) &
        str_detect(programa, regex("ESG", ignore_case = TRUE)) ~ "Construção ESG",

      # Construção Basica
      str_detect(programa, regex("Constr", ignore_case = TRUE)) &
        str_detect(programa, regex("Basic", ignore_case = TRUE)) ~ "Construção Basica",

      # Construção Primária
      (
        str_detect(programa, regex("Constr", ignore_case = TRUE)) &
          str_detect(programa, regex("ACELER", ignore_case = TRUE))
      ) |
        str_detect(programa, regex("Construir e apetrechar escolinhas comunitarias", ignore_case = TRUE)) |
        str_detect(programa, regex("CONSTRUIR CENTROS DE APOIO A PRENDIZAGEM NO AMBITO DO PROJECTO MOZLEANING", ignore_case = TRUE)) ~ "Construção Primária",

      # Requalificação
      str_detect(programa, regex("Requ", ignore_case = TRUE)) ~ "Requalificação de escolas primárias em basicas",

      # Alimentação
      str_detect(programa, regex("Aliment", ignore_case = TRUE)) ~ "Alimentação",

      # Lares
      str_detect(programa, regex("Lare", ignore_case = TRUE)) ~ "Fundo de Apoio Alimentar para os Centros Internatos e Lares",

      TRUE ~ "Outro"
    )
  )



df_final %>% distinct(programa_tipo) %>% print(n = Inf)


# Default output folder
gravar_extracto_sistafe(df_final)


