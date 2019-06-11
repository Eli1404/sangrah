rm (list = ls())
library(xlsx)
library(tidyverse)
library(psych)
x <- getwd()

#Lectura de datos
base <- read.xlsx("BASE GENERAL.xlsx", sheetName = "Hoja1",rowIndex = 2:1000, header = T)
attach(base)
base$IMPORTE.VENCIDO <- round(IMPORTE.VENCIDO,1)
 
#Creación de columna MORA y selección de conceptos de intéres
base.filtro <- base %>%
  mutate(ESTADO_CUENTA = ifelse(NÚMERO.DE.RETRASOS %in% c(2,3), "Mora 1", ifelse (NÚMERO.DE.RETRASOS %in% c(4,5), "Mora 2", 
                               ifelse (NÚMERO.DE.RETRASOS %in% c(6,7,8), "Mora 3", ifelse (NÚMERO.DE.RETRASOS %in% c(9,10,11), "Mora 4",
                                               ifelse (NÚMERO.DE.RETRASOS %in% 12:14, "Mora 5", ifelse (NÚMERO.DE.RETRASOS >= 15, "Mora 6",
                                                               ifelse (NÚMERO.DE.RETRASOS < 1, 0,NA)))))))) %>%
  select(c(2,1,5,16,21,36,39,52,37,22,40,34,35,41,43,45,49,50)) %>%
  arrange(ESTADO_CUENTA, desc(IMPORTE.VENCIDO))

#Frecuencia, porcentaje de cliente por mora 
estado.general <- base.filtro %>%
  group_by(POLÍTICA.DE.PAGO,ESTADO.1, ESPECIFICO, ESTADO_CUENTA) %>%
  arrange(ESTADO_CUENTA) %>%
  summarise(Monto.Vencido = sum(IMPORTE.VENCIDO),
            Saldo.Actual = sum(SALDO.ACTUAL),
            no.clientes = n()) %>%
  arrange(ESTADO_CUENTA, desc(Monto.Vencido))
estado.general$porcentaje <- round((estado.general$no.clientes / sum(estado.general$no.clientes)) * 100,1)

#Frecuencia, porcentaje de cliente por mora y paquete
paquete <- base.filtro %>%
  group_by(CLAVE.DE.PAQUETE,POLÍTICA.DE.PAGO,ESTADO.1, ESPECIFICO, ESTADO_CUENTA) %>%
  arrange(ESTADO_CUENTA, IMPORTE.VENCIDO) %>%
  summarise(Monto = sum(IMPORTE.VENCIDO),
            no.clientes = n()) %>%
  arrange(CLAVE.DE.PAQUETE, ESTADO_CUENTA, desc(Monto))
paquete$porcentaje <- round((paquete$no.clientes / sum(paquete$no.clientes)) * 100,1)

#Frecuencia, porcentaje de cliente por mora y asesor
asesor <- base.filtro %>%
  group_by(POLÍTICA.DE.PAGO,ESTADO.1, ESPECIFICO, ESTADO_CUENTA, ASESOR, COORDINADOR) %>%
  arrange(ESTADO_CUENTA, IMPORTE.VENCIDO) %>%
  summarise(Monto = sum(IMPORTE.VENCIDO),
            no.clientes = n()) %>%
  arrange(ESTADO_CUENTA, desc(Monto))
asesor$porcentaje <- round((asesor$no.clientes / sum(asesor$no.clientes)) * 100,1)

#Frecuencia, porcentaje de cliente por mora y coordinador
coordinador <- base.filtro %>%
  group_by(POLÍTICA.DE.PAGO,ESTADO.1, ESPECIFICO, ESTADO_CUENTA, COORDINADOR) %>%
  arrange(ESTADO_CUENTA, IMPORTE.VENCIDO) %>%
  summarise(Monto = sum(IMPORTE.VENCIDO),
            no.clientes = n()) %>%
  arrange(ESTADO_CUENTA, desc(Monto))
coordinador$porcentaje <- round((coordinador$no.clientes / sum(coordinador$no.clientes)) * 100,1)

#Cambio de carpetas
currenTime <- Sys.time()
newdir <- paste("Resumen",currenTime, sep = "_")
dir.create(newdir)
setwd(newdir)

#Escritura de datos en excel
write.xlsx(as.data.frame(base.filtro), file = paste0("Resumen", currenTime, sep = "_",".xlsx"), 
           sheetName="BASE_GENERAL_FILTRO", col.names = T, row.names=F,  append = F)
write.xlsx(as.data.frame(estado.general), file = paste0("Resumen", currenTime, sep = "_",".xlsx"), 
           sheetName="ESTADO_GENERAL", col.names = T, row.names=F,  append = T)
write.xlsx(as.data.frame(paquete), file = paste0("Resumen", currenTime, sep = "_",".xlsx"), 
           sheetName="PAQUETE", col.names = T, row.names=F,  append = T)
write.xlsx(as.data.frame(asesor), file = paste0("Resumen", currenTime, sep = "_",".xlsx"), 
           sheetName="ASESOR", col.names = T, row.names=F,  append = T)
write.xlsx(as.data.frame(coordinador), file = paste0("Resumen", currenTime, sep = "_",".xlsx"), 
           sheetName="COORDINADOR", col.names = T, row.names=F,  append = T)

#Gráficas
png("Clientes por tipo de mora.png", width=1280,height=800)
estado.general %>%
  na.omit() %>%
  ggplot(aes(ESTADO_CUENTA,porcentaje, fill = ESTADO_CUENTA)) +
  geom_bar(stat = "identity", alpha =0.8) + 
  scale_fill_hue(l=41, c=90) +
  ggtitle("Clientes por tipo de mora") +
  ggrepel::geom_text_repel(aes(label = no.clientes), color = "black", size = 2.5, segment.color = "grey") +
  scale_y_continuous(breaks = c(.5,1,1.5,2,2.5,
                                3,3.5,4,4.5,5,5.5,6,6.5)) +
  xlab(element_blank()) +  
  theme(text = element_text(size =14))+
  guides(fill = F)
dev.off()

png("Clientes en mora por tipo de paquete.png", width=1280,height=800)
paquete %>%
  na.omit() %>%
  ggplot(aes(CLAVE.DE.PAQUETE,porcentaje, fill = CLAVE.DE.PAQUETE)) +
  geom_bar(stat = "identity", alpha =0.8) + 
  ggrepel::geom_text_repel(aes(label = no.clientes), alpha = 1,size = 2.5, segment.color = "grey") +
  scale_fill_hue(l=40, c=80) +
  facet_wrap(~ESTADO_CUENTA) +
  ggtitle("Clientes en mora por tipo de paquete") +
  theme(text = element_text(size =12), axis.text.x = element_blank())
dev.off()

png("Clientes en mora por asesor.png", width=1280,height=800)
asesor%>%
  na.omit() %>%
  ggplot(aes(ASESOR, no.clientes, fill = ESTADO_CUENTA)) +
  geom_bar(stat = "identity", alpha =0.8) + 
  scale_fill_hue(l=41, c=90) +
  ggtitle("Clientes en mora por asesor") +
  theme(axis.text.x =element_blank()) +
  theme(text = element_text(size =14)) 
dev.off()

png("Clientes en mora por coordinador.png", width=1280,height=800)
coordinador %>%
  na.omit() %>%
  ggplot(aes(COORDINADOR,porcentaje, fill = COORDINADOR)) +
  geom_bar(stat = "identity", alpha =0.8) + 
  facet_wrap(~ESTADO_CUENTA) +
  scale_fill_hue(l=41, c=90) +
  ggtitle("Clientes en mora por coordinador") +
  ggrepel::geom_text_repel(aes(label = no.clientes), alpha = 1,size = 2.5, segment.color = "grey") +
  scale_y_continuous(breaks = c(.1,.2,.3,.4,.5,.6,.7,.8,.9,1,
                                1.1,1.2,1.3)) +
  xlab(element_blank()) +  
  theme(text = element_text(size =14), axis.text.x = element_blank())
dev.off()

#correlacion
variables <- base.filtro %>%
  select(4,6,8,9,10,11,14,15) %>%
  na.omit()
png("Correlacion entre conceptos - estatus del cliente.png", width=1280,height=800)
pairs.panels(variables)
dev.off()

#Cambio de carpeta a la inicial
setwd(x) 