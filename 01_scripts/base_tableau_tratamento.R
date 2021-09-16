library(tidyverse)
library(readxl)
library(writexl)
library(stringr)
library(lubridate)
library(openxlsx)

# Carregamento de dados ----

orc_cognos_tbl <- read_excel("00_data/cognos/opex_tidy_2020_2021.xlsx", range = cell_rows(7:373255))
plan_cognos_tbl <- read_excel("00_data/suporte/plan_contas_cognos.xlsx")


# Tratamento de dados

base_cognos_tbl_1 <- orc_cognos_tbl %>%
    filter(Valor != 0) %>% 
    left_join(y = plan_cognos_tbl, by = c("Conta contabil - Cognos" = "Conta-1")) %>% 
    #filter(Valor>0) %>% 
    #separate(col = `Centro de custo`, into = c("centro_custo_cognos", "centro_custo_cognos_nome"), sep = "-") %>%
    #separate(col = `Conta contabil - Cognos`, into = c("DRE", "conta_contabil_cognos", "conta_contabil_cognos_nome", "complemento"), sep = "-") %>% 
    #unite(orc_cognos_tbl, conta_contabil_cognos_nome, complemento, sep = "-", na.rm = TRUE) %>% 
    mutate(Empresa = str_replace(Empresa,"Aegea Saneamento - Matriz", "CAA"),
           centro_custo_cognos = str_trim(`Centro de custo`)) %>% 
    set_names(c("Empresa", "CC", "Conta-1", "Tipo", "Ano", "Mes",
                "Valor", "Grupo Conta Nivel 1", "Grupo Conta Nivel 2", 
                "Grupo Conta Nivel 3", "Grupo Conta Nivel 4", "CC_2")) %>% 
    select(-CC_2)

# Export

base_cognos_tbl_1 %>% 
    write_xlsx("07_outputs/base_cognos_tbl_1.xlsx")