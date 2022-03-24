
options(scipen = 100000000000000000)

library(dplyr)
library(tidyr)
library(broom)
library(readxl)
library(openxlsx)
#########################################################################################################################
# load dataset
clsi <- read_excel("C:/Users/~mypath~/CLSI_HTM_vs_MHF_Media.xlsx")
# check column names
names(clsi)

# select drug, DSI columns
df1 <-clsi%>%select(Drug,HTM_DSI,MHF_BBL_DSI,MHF_Difco_DSI)

utils::View(df1)

# check if there blanks or missing in the dataframe col by col
colSums(is.na(df1))

# Omit rows with NAs
df1<- na.omit(df1)

# check if data is from a normal distribution 
shapiro.test(df1$MHF_Difco_DSI)

shapiro.test(df1$MHF_BBL_DSI)

shapiro.test(df1$HTM_DSI)
#################################################################################################################################
# Bonferroni test

bonferroni_result <- df1 %>% 
  pivot_longer(cols = -Drug) %>% 
  group_by(Drug) %>%group_modify(~ .x %>%summarise(result = list(pairwise.t.test(value, name,p.adjust.method = "bonferroni") %>% tidy))) 
  %>% ungroup %>% unnest(result)
  
################################################################################################################################
#Export as excel worksheet

wb <- createWorkbook()

addWorksheet(wb, "clsi")

writeData(wb, "clsi", bonferroni_result)

highlight <- createStyle(fontColour = rgb(1, 0, 0),fontSize = 11, fontName = "Arial",#fgFill = "#1A33CC"
                         #border = "TopBottomLeftRight",borderColour = "black",
                         #bgFill = "#FFFF05"
                         )

conditionalFormatting(wb,
                      "clsi",
                      cols = 1:length(bonferroni_result),
                      rows = 1:nrow(bonferroni_result),
                      rule = "$D1<0.05",
                      style = highlight)

saveWorkbook(wb, "CLSI_Analysis.xlsx", overwrite = TRUE)


