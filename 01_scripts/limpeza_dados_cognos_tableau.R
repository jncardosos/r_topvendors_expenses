

library(tidyverse)
library(readxl)
library(writexl)
library(stringr)
library(lubridate)
library(openxlsx)

orc_cognos_tbl <- read_excel("00_data/cognos/dB_orcamento.xlsx", range = cell_rows(7:36000))
plan_cognos_tbl <- read_excel("00_data/suporte/plano_conta_cognos.xlsx", sheet = "plano_conta_cognos")


# 1.0 Arquivo Cognos Opex

orc_cognos_limpo_tbl <- orc_cognos_tbl %>%
    left_join(y = plan_cognos_tbl, by = c("Conta contabil - Cognos" = "Conta-1")) %>% 
    select(Empresa, `Centro de custo`, `Conta contabil - Cognos`, Tipo, Ano, Mes, Valor, `Grupo Conta Nivel 1`, `Grupo Conta Nivel 2`, `Grupo Conta Nivel 3`, `Grupo Conta Nivel 4`) %>% 
    #filter(Valor>0) %>% 
    #separate(col = `Centro de custo`, into = c("centro_custo_cognos", "centro_custo_cognos_nome"), sep = "-") %>%
    #separate(col = `Conta contabil - Cognos`, into = c("DRE", "conta_contabil_cognos", "conta_contabil_cognos_nome", "complemento"), sep = "-") %>% 
    #unite(orc_cognos_tbl, conta_contabil_cognos_nome, complemento, sep = "-", na.rm = TRUE) %>% 
    #mutate(Empresa = str_replace(Empresa,"Aegea Saneamento - Matriz", "CAA"),
    #centro_custo_cognos = str_trim(centro_custo_cognos)) %>% 
    set_names(c("Empresa", "CC", "Conta-1", "Tipo", "Ano", "Mes", "Valor", "Grupo Conta Nivel 1", "Grupo Conta Nivel 2", "Grupo Conta Nivel 3", "Grupo Conta Nivel 4"))
    

# orc_cognos_limpo_tbl %>%
#     #filter(Tipo == "Realizado") %>%
#     #filter(Empresa == "Aegea Saneamento - Matriz") %>%
#     filter(Ano == "2020") %>% 
#     group_by(centro_custo_cognos_nome, Tipo, Empresa) %>%
#     summarise(total_ano = sum(Valor)/1000000) %>% 
#     arrange(total_ano) %>% 
#     spread(key = Tipo, value = total_ano) %>% 
#     ungroup() %>% 
#     filter(Realizado != 0 | `Rolling Forecast 3T20_Ajustado CA_JAN21` != 0) %>% set_names(c("Centro de Custos", "Empresa", "Realizado", "RFCST3")) %>% 
#     select(Empresa, `Centro de Custos`, RFCST3, Realizado) %>% 
#     group_by(Empresa) %>% 
#     mutate(`Proporcção Orçamento Total` = RFCST3/sum(RFCST3),
#            Desvio = Realizado - RFCST3,
#            `Desvio Pct` = Realizado/RFCST3 - 1) %>% 
#     arrange(Empresa) 
    
   
# 2.0 Outputs

# 2.1 Cognos limpo - Tableau
orc_cognos_limpo_tbl %>% 
    write_xlsx("07_outputs/dados_opex_cognos_tableau.xlsx")
