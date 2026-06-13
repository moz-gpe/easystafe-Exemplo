library(tidyverse)
library(readxl)
library(writexl)
library(janitor)
library(glue)
library(arrow)
library(pins)
library(easystafe)
library(esquisse)

# GLOBAL VARS -------------------------------------------------

source <- "Documents/temp_compare.xlsx"

old <- read_excel(source, sheet = "old") %>%
  select(-c(conjunto, Integra)) %>%
  mutate(key = paste(reporte_tipo, ambito, programa, ugb, ano, mes, sep = "-"))

new <- read_excel(source, sheet = "new") %>%
  select(-'...23') %>%
  select(-c(ced_nome, ced_2_nome, ced_3_nome)) %>%
  mutate(key = paste(reporte_tipo, ambito, programa, ugb, ano, mes, sep = "-"))

# COMPARE COLUMNS -------------------------------------------------

old_only <- setdiff(names(old), names(new))
new_only <- setdiff(names(new), names(old))
in_both <- intersect(names(old), names(new))

cat("In OLD but not NEW:\n")
print(old_only)
cat("In NEW but not OLD:\n")
print(new_only)

# ROWS IN NEW BUT NOT IN OLD --------------------------------------

new_not_old <- anti_join(new, old, by = in_both)


path <- "Documents/lookup.xlsx"

sheets <- c("ugb", "funcao", "programa", "ced", "ced_2", "ced_3", "ced_nivel")
keys <- c(
  "codigo_ugb",
  "funcao",
  "programa_ambito_fr",
  "ced",
  "ced_2",
  "ced_3",
  "ced_3_nome"
)

purrr::walk2(sheets, keys, function(sheet, key) {
  df <- readxl::read_excel(path, sheet = sheet) |> janitor::clean_names()
  dups <- df |> dplyr::count(.data[[key]]) |> dplyr::filter(n > 1)
  if (nrow(dups) > 0) {
    cat("\n--- DUPLICATES in sheet:", sheet, "(key:", key, ") ---\n")
    print(dups)
  } else {
    cat("\nSheet", sheet, "- OK\n")
  }
})


readxl::read_excel(path, sheet = "programa") |>
  janitor::clean_names() |>
  dplyr::add_count(programa_ambito_fr) |>
  dplyr::filter(n > 1) |>
  dplyr::arrange(programa_ambito_fr)
