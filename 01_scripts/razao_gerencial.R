library(tidyverse)
library(readxl)
library(writexl)
library(stringr)
library(lubridate)
library(openxlsx)

# Import ----
razao_sap_tbl <- read_excel("00_data/sap/GLV_GL_ACCOUNT_LINE_ITEMSSet (104).xlsx")
de_para_cognos_conta_tbl <- read_excel("00_data/suporte/de_para_conta_cognos.xlsx")
de_para_cognos_centro_custo_tbl <- read_excel("00_data/suporte/de_para_centro_custo_cognos.xlsx")
financeiro_sap_tbl <- read_excel("00_data/sap/Items (70).xlsx")
plan_cognos_tbl <- read_excel("00_data/suporte/plan_contas_cognos.xlsx")
de_para_categoria_tbl <- read_excel("00_data/suporte/de_para_categoria.xlsx")
categ_conta_tbl <- read_excel("07_outputs/PlanoConta.xlsx")
categ_conta_tbl_2 <- categ_conta_tbl %>% select(`Conta Cognos`, `Grupo de despesa`)

# 0.0 PreWrang ----

de_para_cognos_conta_tbl <- de_para_cognos_conta_tbl %>% 
    mutate(`N conta contábil` = `N conta contábil` %>% as.character())

# 1.0 Wrangling ----

razao_cognos_1_tbl <- razao_sap_tbl %>% 
    filter(!is.na(`Centro de custo`)) %>% 
    mutate(cc_cognos = str_replace_all(`Centro de custo`,c("LV" = '', "GS" = '', "AS" = '', "AE" = '', "FP" = ''))%>% as.numeric()) %>%
    mutate(`Conta do Razão` = `Conta do Razão` %>% as.character()) %>% 
    mutate(Dt.lançamento = ymd(Dt.lançamento)) %>% 
    
    # Join           
    left_join(y = de_para_cognos_conta_tbl, by = c("Conta do Razão" = "N conta contábil"), keep = TRUE) %>% 
    #"DenomLonga conta RZ" = "Descrição conta contábil"), keep = TRUE) %>%  %>%
    left_join(y = de_para_cognos_centro_custo_tbl, by = c("cc_cognos" = "Ccusto SAP"), keep = TRUE) %>% 
    left_join(y = financeiro_sap_tbl, by = c("Partida individual...15" = "Partida individual", "Empresa" = "Empresa"), keep = TRUE) #%>% view()


razao_gerencial_tbl <- razao_cognos_1_tbl%>%
    mutate(Desc_Ordem = unite(razao_cognos_1_tbl, Empresa, Ordem, sep = "_", remove = FALSE)) %>%
    select(Empresa.x, `Nome da empresa`, Dt.lançamento, `Conta do Razão`, 
           `DenomLonga conta RZ`, Conta, Descrição, `Centro de custo`, 
           cc_cognos, `Desc. SAP`, Ordem, Desc_Ordem, `Elemento PEP`, 
           `Partida individual...15`, `Texto de item`, `Mont.moeda empresa`, 
           Referência, Fornecedor, `Nome do fornecedor`, YyEbeln) %>% set_names(c("Empresa", "Empresa Cognos", "Dt.lançamento", "Conta do Razão",
                                                                                  "DenomLonga conta RZ", "Conta Cognos", "Desc. Conta Cognos", 
                                                                                  "Centro de custo", "C.Custo Cognos", "Desc. C.Custo Cognos",
                                                                                  "Ordem", "Desc.Ordem", "Elemento PEP", "Partida individual",
                                                                                  "Texto de item", "Valor", "Nota Fiscal", "Código Fornecedor",
                                                                                  "Nome Fornecedor", "Pedido")) %>% 
    select(Empresa, `Conta do Razão`, `DenomLonga conta RZ`, `Centro de custo`, `C.Custo Cognos`, 
           Ordem, Desc.Ordem, `Elemento PEP`, Pedido, `Empresa Cognos`, Dt.lançamento, `Conta Cognos`,
           `Desc. Conta Cognos`, `Desc. C.Custo Cognos`, `Partida individual`, `Texto de item`, Valor,
           `Nota Fiscal`, `Código Fornecedor`, `Nome Fornecedor`)

# 1.1 Razao_wider

razao_gerencial_pivotado_tbl <- razao_gerencial_tbl %>% mutate(mes_n = Dt.lançamento %>% month(label = TRUE, abbr = FALSE)) %>%
    select(mes_n, Empresa, `Conta do Razão`, `DenomLonga conta RZ`, 
                               `Conta Cognos`, `Desc. Conta Cognos`, `C.Custo Cognos` , `Desc. C.Custo Cognos`, `Partida individual`, 
                               `Código Fornecedor`, `Nome Fornecedor`, `Texto de item`, Valor)%>%  left_join(y = categ_conta_tbl_2, by = c("Conta Cognos" = "Conta Cognos")) %>% distinct() %>% 
    group_by(mes_n, Empresa, `Grupo de despesa`, `Conta do Razão`, `DenomLonga conta RZ`, 
             `Conta Cognos`, `Desc. Conta Cognos`, `C.Custo Cognos` , `Desc. C.Custo Cognos`, 
             `Código Fornecedor`, `Nome Fornecedor`, `Texto de item`) %>% summarise(total_mes = -sum(Valor)) %>% ungroup() %>% 
    pivot_wider(names_from = mes_n, values_from = total_mes)




# 2.0 Export ----

# Aegea Matriz ----

ano  <- razao_gerencial_tbl %>% select(Dt.lançamento) %>% pull() %>% min() %>% year()
mes_i <- razao_gerencial_tbl %>% select(Dt.lançamento) %>% pull() %>% min() %>% month()
mes_f <- razao_gerencial_tbl %>% select(Dt.lançamento) %>% pull() %>% max() %>% month()
dia_i <- razao_gerencial_tbl %>% select(Dt.lançamento) %>% pull() %>% min() %>% day()
dia_f <- razao_gerencial_tbl %>% select(Dt.lançamento) %>% pull() %>% max() %>% day()
today <- today()

razao_gerencial_tbl %>%
    #filter(Dt.lançamento >= as_date("2021-07-01")) %>% 
    #filter(str_detect(`Conta Cognos`, "^4")) %>%
    #filter(Empresa == "AS00") %>%
    filter(`Desc. Conta Cognos` != "Amortização" & `Desc. Conta Cognos` != "Depreciações") %>% 
    arrange(desc(`Conta Cognos`)) %>% 
    write_xlsx(str_glue("07_outputs/Razao Gerencial - CAA_Holding_LVE - 0{dia_i}.0{mes_i}.{ano} à {dia_f}.0{mes_f}.{ano} extraído em {today}.xlsx"))

# LVE ----
razao_gerencial_tbl %>%
    #filter(Dt.lançamento >= as_date("2021-07-01")) %>% 
    #filter(str_detect(`Conta Cognos`, "^4")) %>%
    filter(Empresa == "LV00") %>% 
    #filter(`Desc. Conta Cognos` != "Amortização" & `Desc. Conta Cognos` != "Depreciações") %>% 
    arrange(desc(`Conta Cognos`)) %>% 
    write_xlsx(str_glue("07_outputs/Razao Gerencial - LVE - 0{dia_i}.0{mes_i}.{ano} à {dia_f}.0{mes_f}.{ano} extraído em {today}.xlsx"))

# Razão Gerencial Trainee ----
razao_gerencial_tbl %>%
    filter(`Desc. C.Custo Cognos` == "Trainee") %>% 
    #filter(str_detect(`Conta Cognos`, "^4")) %>%
    filter(Empresa == "AS00") %>% 
    #filter(`Desc. Conta Cognos` != "Amortização" & `Desc. Conta Cognos` != "Depreciações") %>% 
    arrange(desc(`Conta Cognos`)) %>% 
    write_xlsx("07_outputs/AS00_razao_gerencial_02082021_Trainee.xlsx")

# Razão Gerencial Pivotado
razao_gerencial_pivotado_tbl %>% 
    write_xlsx(str_glue("07_outputs/Razao Gerencial Pivotado - AS00_AE00_GS00_LV00 - {dia_i}.0{mes_i}.{ano} à {dia_f}.0{mes_f}.{ano} extraído em {today}.xlsx"))


