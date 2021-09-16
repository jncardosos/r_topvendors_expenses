

library(tidyverse)
library(readxl)
library(writexl)
library(stringr)
library(lubridate)
library(openxlsx)

orc_cognos_tbl <- read_excel("00_data/cognos/opex_tidy_2020_2021.xlsx", range = cell_rows(7:373255))

# Razão
razao_sap_tbl <- read_excel("00_data/sap/GLV_GL_ACCOUNT_LINE_ITEMSSet (98).xlsx")
de_para_cognos_conta_tbl <- read_excel("00_data/suporte/de_para_conta_cognos.xlsx")
de_para_cognos_centro_custo_tbl <- read_excel("00_data/suporte/de_para_centro_custo_cognos.xlsx")
financeiro_sap_tbl <- read_excel("00_data/sap/Items (69).xlsx")
plan_cognos_tbl <- read_excel("00_data/suporte/plan_contas_cognos.xlsx")
de_para_categoria_tbl <- read_excel("00_data/suporte/de_para_categoria.xlsx")


indicadores_fin_orc_tbl <- read_excel("00_data/suporte/Indicadores consolidados (001)_sv (008).xlsx", sheet = "Indicadores_t")
indicadores_fin_dfs_tbl <- read_excel("00_data/suporte/Indicadores consolidados (001)_sv (008).xlsx", sheet = "Indicadores2_t")
real_custo_tbl <- read_excel("00_data/suporte/HC - alocação do físico e financeira.xlsx", sheet = "BaseJan21")


# 0.0 PreTrat

de_para_cognos_conta_tbl <- de_para_cognos_conta_tbl %>% 
    mutate(`N conta contábil` = `N conta contábil` %>% as.character())

real_custo_tab_tbl <- real_custo_tbl %>% 
    group_by(CC_n, Mês, Emp., Estd.) %>% 
    summarise(total_val = sum(`Total Geral`),
              count = n()) %>% 
    #spread(key = Estd., value = n) %>% 
    mutate(Ano = "2021") %>% 
    select(CC_n, Ano, Mês, Emp., everything()) %>% 
    ungroup()

# 1.0 Razão gerencial

razao_cognos_1_tbl <- razao_sap_tbl %>% 
    filter(!is.na(`Centro de custo`)) %>% 
    mutate(cc_cognos = str_replace_all(`Centro de custo`,c("LV" = '', "GS" = '', "AS" = ''))%>% as.numeric()) %>%
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
                                                                      "Nome Fornecedor", "Pedido"))
# 1.1 Arquivo Cognos ----

orc_cognos_limpo_tbl <- orc_cognos_tbl %>%
    left_join(y = plan_cognos_tbl, by = c("Conta contabil - Cognos" = "conta_contabil")) %>% 
    #filter(Valor>0) %>% 
    separate(col = `Centro de custo`, into = c("centro_custo_cognos", "centro_custo_cognos_nome"), sep = "-") %>%
    separate(col = `Conta contabil - Cognos`, into = c("DRE", "conta_contabil_cognos", "conta_contabil_cognos_nome", "complemento"), sep = "-") %>% 
    unite(orc_cognos_tbl, conta_contabil_cognos_nome, complemento, sep = "-", na.rm = TRUE) %>% 
    mutate(Empresa = str_replace(Empresa,"Aegea Saneamento - Matriz", "CAA"),
           centro_custo_cognos = str_trim(centro_custo_cognos))

orc_cognos_limpo_2_tbl <- orc_cognos_limpo_tbl %>% 
    full_join(y = indic_fin_tab_tbl, by = c("centro_custo_cognos_nome" = "Àrea", "Empresa" = "Empresa", "Ano" = "Ano")) %>%
    full_join(y = real_custo_tab_tbl, by = c("centro_custo_cognos" = "CC_n", "Mes" = "Mês", "Empresa" = "Emp.", "Ano" = "Ano"))
    

orc_cognos_limpo_tbl %>%
    #filter(Tipo == "Realizado") %>%
    #filter(Empresa == "Aegea Saneamento - Matriz") %>%
    filter(Ano == "2020") %>% 
    group_by(centro_custo_cognos_nome, Tipo, Empresa) %>%
    summarise(total_ano = sum(Valor)/1000000) %>% 
    arrange(total_ano) %>% 
    spread(key = Tipo, value = total_ano) %>% 
    ungroup() %>% 
    filter(Realizado != 0 | `Rolling Forecast 3T20_Ajustado CA_JAN21` != 0) %>% set_names(c("Centro de Custos", "Empresa", "Realizado", "RFCST3")) %>% 
    select(Empresa, `Centro de Custos`, RFCST3, Realizado) %>% 
    group_by(Empresa) %>% 
    mutate(`Proporcção Orçamento Total` = RFCST3/sum(RFCST3),
           Desvio = Realizado - RFCST3,
           `Desvio Pct` = Realizado/RFCST3 - 1) %>% 
    arrange(Empresa) 
    
    # Format

# 1.1 Cognos x SAP ----

razao_cognos_2_tbl <- razao_gerencial_tbl %>% 
    filter(str_detect(`Conta Cognos`, "^4")) %>%
    filter(`Desc. Conta Cognos` != "Amortização" & `Desc. Conta Cognos` != "Depreciações") %>%
    mutate(`Empresa Cognos` = str_replace(`Empresa Cognos`, "Aegea San.", "Aegea Saneamento - Matriz")) %>% 
    mutate(`C.Custo Cognos` = `C.Custo Cognos` %>% as.character(),
           Mes = Dt.lançamento %>% month(label = TRUE, abbr = FALSE, locale = Sys.getlocale("LC_TIME")) %>% as.character() %>% str_to_title()) %>% 
    group_by(`Empresa Cognos`, `Conta Cognos`, `C.Custo Cognos`, `Desc. C.Custo Cognos`, `Desc. Conta Cognos`, Mes) %>% 
    summarise(total_razao = sum(Valor))


cognos_sap_tbl <- orc_cognos_limpo_tbl %>%
    #filter(Tipo == "Realizado") %>%
    #filter(Empresa == "Aegea Saneamento - Matriz") %>% 
    filter(Empresa == "Aegea Saneamento - Matriz") %>%
    filter(Ano == "2021") %>% 
    spread(key = Tipo, value = Valor) %>%
    group_by(Empresa, centro_custo_cognos, centro_custo_cognos_nome, conta_contabil_cognos, orc_cognos_tbl, Mes, categoria_custo_1, categoria_custo_2, categoria_custo_3) %>%
    summarise(Realizado = sum(-Realizado),
              `Orçamento vigente` = sum(-`Rolling Forecast 3T20_ Encaixe_Real_2020`)) %>% select(Realizado) %>% pull %>% sum()
    mutate(conta_contabil_cognos = conta_contabil_cognos %>% str_trim(),
           centro_custo_cognos = centro_custo_cognos %>% str_trim()) %>% 
    left_join(y = razao_cognos_2_tbl, by = c("conta_contabil_cognos" = "Conta Cognos", "centro_custo_cognos" = "C.Custo Cognos", "Mes" = "Mes", "Empresa" = "Empresa Cognos"), keep = TRUE) %>% 
    select(Empresa, Mes.x, Mes.y, centro_custo_cognos, centro_custo_cognos_nome, categoria_custo_1, categoria_custo_2, categoria_custo_3,conta_contabil_cognos, orc_cognos_tbl, `Orçamento vigente`, Realizado, total_razao) %>%
    set_names(c("Empresa","Mes", "Mes SAP", "Centro de Custos (Cognos)", "Centro de custos descrição","Custo tier 1", "Custo tier 2", "Custo tier 3", "Conta contábil", "Conta contábil descrição", "Orçamento vigente", "Realizado (cognos)", "Realizado (razão SAP)")) %>% 
    mutate(`Desvio Orçado x Realizado` = `Orçamento vigente` - `Realizado (cognos)`,
           `Desvio Cognos x Razão` = `Realizado (cognos)` - `Realizado (razão SAP)`) %>% 
    select(Empresa, Mes, `Mes SAP`, `Centro de Custos (Cognos)`, `Centro de custos descrição`,
           `Custo tier 1`, `Custo tier 2`, `Custo tier 3`, `Conta contábil`, `Conta contábil descrição`,
           `Orçamento vigente`, `Realizado (cognos)`, `Desvio Orçado x Realizado`, `Realizado (razão SAP)`, 
           `Desvio Cognos x Razão`) %>% 
    left_join(y = de_para_categoria_tbl, by = c("Centro de custos descrição" = "Centro de custos descrição", "Empresa" = "Empresa")) %>%
    select(-...4, -`Centro de Custos (Cognos).y`)
    
   
# 2.0 Outputs

# Comparativo razão x cognos ----

wb <- createWorkbook()

addWorksheet(wb, sheetName = "Realizado (Cognos x SAP)", gridLines = FALSE, zoom = 80)

headers <- createStyle(
    fontSize = 12,
    fgFill = "#2c3e50",
    fontColour = "#ffffff",
    border = "TopBottom",
    borderColour = "#000000",
    textDecoration = "bold",
    halign = "center"
)

addStyle(wb, sheet = "Realizado (Cognos x SAP)", headers, cols = 1:16, rows = 1)

setColWidths(wb, "Realizado (Cognos x SAP)", cols = 1:8, widths = 25)

#writeDataTable(wb, sheet = "Realizado (Cognos x SAP)", x = cognos_sap_tbl)

writeData(wb, sheet = "Realizado (Cognos x SAP)", x = cognos_sap_tbl, rowNames = FALSE)

saveWorkbook(wb, "07_outputs/Realizado cog x Realizado sap.xlsx", overwrite = TRUE)


# Razão Gerencial AEGEA ----
razao_gerencial_tbl %>%
    
    #filter(Dt.lançamento >= as_date("2021-07-01")) %>% 
    #filter(str_detect(`Conta Cognos`, "^4")) %>%
    filter(Empresa == "AS00") %>%
    filter(`Desc. Conta Cognos` != "Amortização" & `Desc. Conta Cognos` != "Depreciações") %>% 
    arrange(desc(`Conta Cognos`)) %>% 
    write_xlsx("07_outputs/AS00_razao_gerencial_0121a0721_10082021.xlsx")

# Razão Gerencial LVE ----
razao_gerencial_tbl %>%
    filter(Dt.lançamento >= as_date("2021-07-01")) %>% 
    #filter(str_detect(`Conta Cognos`, "^4")) %>%
    filter(Empresa == "LV00") %>% 
    #filter(`Desc. Conta Cognos` != "Amortização" & `Desc. Conta Cognos` != "Depreciações") %>% 
    arrange(desc(`Conta Cognos`)) %>% 
    write_xlsx("07_outputs/LV00_razao_gerencial_31082021.xlsx")

# Razão Gerencial Trainee ----
razao_gerencial_tbl %>%
    filter(`Desc. C.Custo Cognos` == "Trainee") %>% 
    #filter(str_detect(`Conta Cognos`, "^4")) %>%
    filter(Empresa == "AS00") %>% 
    #filter(`Desc. Conta Cognos` != "Amortização" & `Desc. Conta Cognos` != "Depreciações") %>% 
    arrange(desc(`Conta Cognos`)) %>% 
    write_xlsx("07_outputs/AS00_razao_gerencial_02082021_Trainee.xlsx")

# Razão Gerencial GSS ----
razao_gerencial_tbl %>%
    filter(str_detect(`Conta Cognos`, "^4")) %>%
    filter(Empresa == "GS00") %>%
    filter(`Desc. Conta Cognos` != "Amortização" & `Desc. Conta Cognos` != "Depreciações") %>% 
    arrange(desc(`Conta Cognos`)) %>% 
    write_xlsx("07_outputs/GS00_razao_gerencial_04032021.xlsx")

# Cognos limpo - Tableau
orc_cognos_limpo_2_tbl %>% 
    mutate(Ebtida = Ebtida %>% as.double(),
           Valor.y = Valor.y %>% as.double(),
           `Receita Líquida` = `Receita Líquida` %>% as.double(),
           `Receita Bruta` = `Receita Bruta` %>% as.double(),
           Economia = Economia %>% as.double()) %>% 
    #filter(Valor!=0) %>%
    #left_join(y = de_para_categoria_tbl, by = c("centro_custo_cognos_nome" = "Centro de custos descrição", "Empresa" = "Empresa")) %>% view()
    mutate(Valor = Valor.x*-1) %>% 
    spread(key = Tipo, value = Valor.x) %>% 
    write_xlsx("07_outputs/base_cognos_tratada_2.xlsx")


# Indicadores financeiros - Tableau ----

indicadores_fin_dfs_tidy_tbl <- indicadores_fin_dfs_tbl %>% 
    gather(key = Ano, value = Valor, "2015", "2016", "2017", "2018", "2019", "2020", "2021e") %>% 
    spread(key = ...1, value = Valor)

indic_fin_tab_tbl <- indicadores_fin_orc_tbl %>% 
    gather(key = Ano, value = Valor, "2015", "2016", "2017", "2018", "2019", "2020", "2021e") %>% 
    left_join(y = indicadores_fin_dfs_tidy_tbl, by = c("Ano" = "Ano")) %>% 
    mutate(Valor = Valor*1000000) %>% 
    mutate(Empresa = case_when(
        `Àrea` == "LVE" ~ "LVE",
        `Àrea` == "GSS" ~ "GSS",
        TRUE ~ "CAA"
    )) #%>% 
    #write_xlsx("00_data/suporte/indicadores_tidy.xlsx")
    