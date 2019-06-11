rm (list = ls())
library(readxl)
library(tidyverse)
library(plyr)
library(tibble)
x <- getwd()

#Lectura de excel
llamadas <- "cobraza_contactados.xlsx"
tab_names <- excel_sheets(path = llamadas)
list_all <- lapply(tab_names, function(x) read_excel(path = llamadas, sheet = x))
llamadas_file <- rbind.fill(list_all)

#Conteo de llamadas
llamadas_gestor <- llamadas_file %>% 
  replace(is.na(.), 0) %>% 
  group_by(gestor, mes, fecha) %>%
  select(10:27) %>%
  summarise_each(funs(sum)) 

#Cambio de carpetas
currenTime <- Sys.time()
newdir <- paste("Llamadas cobranza",currenTime, sep = "_")
dir.create(newdir)
setwd(newdir)

png("Total de llamadas de cobranza por gestor.png", width=1280,height=800)
  ggplot(llamadas_gestor, aes(fecha,`total de contactos`, color = gestor, shape = gestor)) +
  geom_point() +
    facet_wrap(~mes) +
    ggtitle("Total de llamadas de cobranza por gestor") +
    geom_line() +
    scale_fill_hue(l=40, c=80) +
    theme(text = element_text(size = 12))
dev.off()

png("Respuesta de llamadas de cobranza por gestor.png", width=1280,height=800)
llamadas_gestor %>%
  gather(5:21,  key = "Respuesta", value = "Total") %>%
  ggplot(aes(gestor,Total, fill = Respuesta)) +
  geom_bar(stat="identity") +
  facet_wrap(~mes) +
  scale_fill_hue(l=40, c=80) +
  ggtitle("Total de llamadas de cobranza por gestor") +
  theme(text = element_text(size = 12))
dev.off()

