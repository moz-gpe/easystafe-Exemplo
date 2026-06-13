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
library(pins)
library(easystafe)

# GLOBAL VARS -------------------------------------------------

paths_razao_contabilistica <- c(
  "Data/razao_cont/202601/",
  "Data/razao_cont/202602/",
  "Data/razao_cont/202603/",
  "Data/razao_cont/202604/"
)

rates_diarias <- obter_conversao_bancomoc(wide = TRUE)


# PROCESSAR E GRAVAR RAZAO CONTABILISTICA. & ABSA ---------------------------------------------------------

df_razao <- paths_razao_contabilistica |>
  map(\(path) {
    df_r <- processar_extracto_razao_c(source_path = path)
    df_a <- processar_extracto_absa(source_path = path)
    bind_rows(df_r, df_a)
  }) |>
  list_rbind() |>
  left_join(lookup_razao, by = "source_file")

# TAXAS DE CAMBIO BANCO DE MOCAMBIQUE ------------------------------------

df_razao <- adicionar_conversao_moeda(df_razao, rates_diarias)

df_razao |> distinct(tipo)

df_razao |>
  ggplot(aes(data, taxa_euro)) +
  geom_line()

# WRITE TO DISK -------------------------------------------------

gravar_extracto_razao_c(df_razao)
