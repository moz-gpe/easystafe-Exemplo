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

lookup_razao <- tribble(
  ~source_file                            , ~descricao                                     , ~provincia         ,
  "CENTRAL EUR.pdf"                       , "Conta 7321000 CUT EUR - MEF"                  , NA                 ,
  "CENTRAL EURO.pdf"                      , "Conta 7321000 CUT EUR - MEF"                  , NA                 ,
  "CENTRAL MT.pdf"                        , "Conta 7321000 CUT MTN - MEF"                  , NA                 ,
  "CENTRAL USD.pdf"                       , "Conta 7321000  CUT USD - MEF"                 , NA                 ,
  "EXTRACTO ABSA BANK MZN.pdf"            , "ABSA BANK MT"                                 , NA                 ,
  "EXTRACTO ABSA BANK MT.pdf"             , "ABSA BANK MT"                                 , NA                 ,
  "EXTRACTO ABSA BANK USD.pdf"            , "ABSA BANK USD"                                , NA                 ,
  "FOREX EUR"                             , "Conta nr.004037601011   EUR  Forex - B.M. "   , NA                 ,
  "FOREX USD"                             , "Conta Especial 004187601014 - USD FOREX B.M." , NA                 ,
  "CABO DELGADO DESCENTRALIZADO.pdf"      , "Cabo Delgado"                                 , "Cabo Delgado"     ,
  "CABO DELGADO.pdf"                      , "Cabo Delgado"                                 , "Cabo Delgado"     ,
  "GAZA DESCENTRALIZADO.pdf"              , "Gaza"                                         , "Gaza"             ,
  "GAZA.pdf"                              , "Gaza"                                         , "Gaza"             ,
  "INHAMBANE DESCENTRALIZADO.pdf"         , "Inhambane"                                    , "Inhambane"        ,
  "INHAMBANE.pdf"                         , "Inhambane"                                    , "Inhambane"        ,
  "MANICA DESCENTRALIZADO.pdf"            , "Manica"                                       , "Manica"           ,
  "MANICA.pdf"                            , "Manica"                                       , "Manica"           ,
  "MAPUTO CIDADE.pdf"                     , "Maputo Cidade"                                , "Maputo Cidade"    ,
  "M. PROVINCIA DESCENTRALIZADO.pdf"      , "Maputo Província"                             , "Maputo Província" ,
  "MAPUTO PROVINCIA DESCENTRALIZADO .pdf" , "Maputo Província"                             , "Maputo Província" ,
  "MAPUTO PROVINCIA DESCETRALIZADO.pdf"   , "Maputo Província"                             , "Maputo Província" ,
  "MAPUTO PROVINCIA.pdf"                  , "Maputo Província"                             , "Maputo Província" ,
  "NAMPULA DESCENTRALIZADO.pdf"           , "Nampula"                                      , "Nampula"          ,
  "NAMPULA.pdf"                           , "Nampula"                                      , "Nampula"          ,
  "NIASSA DESCENTRALIZADO.pdf"            , "Niassa"                                       , "Niassa"           ,
  "NIASSA.pdf"                            , "Niassa"                                       , "Niassa"           ,
  "SOFALA DESCENTRALIZADO.pdf"            , "Sofala"                                       , "Sofala"           ,
  "SOFALA.pdf"                            , "Sofala"                                       , "Sofala"           ,
  "TETE DESCENTRALIZADO.pdf"              , "Tete"                                         , "Tete"             ,
  "TETE.pdf"                              , "Tete"                                         , "Tete"             ,
  "ZAMBEZIA DESCENTRALIZADO.pdf"          , "Zambézia"                                     , "Zambézia"         ,
  "ZAMBEZIA.pdf"                          , "Zambézia"                                     , "Zambézia"
)

# PROCESSAR E GRAVAR RAZAO CONTABILISTICA. & ABSA ---------------------------------------------------------

df_razao <- paths_razao_contabilistica |>
  map(\(path) {
    df_r <- processar_extracto_razao_c(source_path = path)
    df_a <- processar_extracto_absa(source_path = path)
    bind_rows(df_r, df_a)
  }) |>
  list_rbind() |>
  left_join(lookup_razao, by = "source_file") |>
  relocate(descricao, provincia, .after = source_file)

# TAXAS DE CAMBIO BANCO DE MOCAMBIQUE ------------------------------------

rates_diarias <- obter_conversao_bancomoc(wide = TRUE)

df_razao <- adicionar_conversao_moeda(df_razao, rates_diarias)

# WRITE TO DISK -------------------------------------------------

gravar_extracto_razao_c(df_razao)
