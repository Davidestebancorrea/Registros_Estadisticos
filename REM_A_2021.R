
# Librerias ---------------------------------------------------------------
library(tidyverse)
library(readxl)
library(lubridate)
library(janitor)
library(dplyr)
library(openxlsx)

anio <- "2022"
meses <- c("01")
# ("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11","12")
for (i in meses) {


fecha_mes <- paste0(anio,"-",i,"-01")
ruta_base <- "C:/Users/control.gestion3/OneDrive/"
archivo <- paste0(ruta_base,"BBDD Produccion/REM/Serie A/",anio,"/",anio,"-",i," REM serie A.xlsx")

# Representan las BBDD donde esta guardada la información -----------------
A07BBDD <- paste0(ruta_base,"BBDD Produccion/Ambulatorio/A07 BBDD.xlsx")
A07CNBBDD <- paste0(ruta_base,"BBDD Produccion/Ambulatorio/A07 BBDD CN.xlsx")
A30BBDD <- paste0(ruta_base,"BBDD Produccion/Ambulatorio/A30 BBDD.xlsx")
A32BBDD <- paste0(ruta_base,"BBDD Produccion/Ambulatorio/A32 BBDD.xlsx")
A28BBDD <- paste0(ruta_base,"BBDD Produccion/Ambulatorio/A28 BBDD.xlsx")
A08BBDD01 <- paste0(ruta_base,"BBDD Produccion/Urgencia/A08 BBDD_01.xlsx")
A08BBDD02 <- paste0(ruta_base,"BBDD Produccion/Urgencia/A08 BBDD_02.xlsx")
A08BBDD03 <- paste0(ruta_base,"BBDD Produccion/Urgencia/A08 BBDD_03.xlsx")
A08BBDD04 <- paste0(ruta_base,"BBDD Produccion/Urgencia/A08 BBDD_04.xlsx")
A08BBDD05 <- paste0(ruta_base,"BBDD Produccion/Urgencia/A08 BBDD_05.xlsx")
A08BBDD06 <- paste0(ruta_base,"BBDD Produccion/Urgencia/A08 BBDD_06.xlsx")
A09BBDD01 <- paste0(ruta_base,"BBDD Produccion/Ambulatorio/A09 BBDD_01.xlsx")
A09BBDD02 <- paste0(ruta_base,"BBDD Produccion/Ambulatorio/A09 BBDD_02.xlsx")
A09BBDD03 <- paste0(ruta_base,"BBDD Produccion/Ambulatorio/A09 BBDD_03.xlsx")
A06BBDD <- paste0(ruta_base,"BBDD Produccion/Ambulatorio/A06 BBDD.xlsx")
A19bBBDD <- paste0(ruta_base,"BBDD Produccion/OIRS/A19b BBDD.xlsx")
A211BBDD <- paste0(ruta_base,"BBDD Produccion/Quirurgico/A21_1 BBDD.xlsx")
A212BBDD <- paste0(ruta_base,"BBDD Produccion/Quirurgico/A21_2 BBDD.xlsx")
A213BBDD <- paste0(ruta_base,"BBDD Produccion/Quirurgico/A21_3 BBDD.xlsx")
A214BBDD <- paste0(ruta_base,"BBDD Produccion/Quirurgico/A21_4 BBDD.xlsx")
A041BBDD <- paste0(ruta_base,"BBDD Produccion/ADyT/A04_1 BBDD.xlsx")
A043BBDD <- paste0(ruta_base,"BBDD Produccion/ADyT/A04_3 BBDD.xlsx")
B_IMGBBDD <- paste0(ruta_base,"BBDD Produccion/ADyT/B_IMG BBDD.xlsx")


# A32 ----------------------------------------------------------------------
A32B <- read_excel(A32BBDD)

A32T <- read_excel(archivo, sheet = "A32", range = "B30", col_names = FALSE)
A32N <- read_excel(archivo, sheet = "A32", range = "Z30", col_names = FALSE) + read_excel(archivo, sheet = "A32", range = "AE30", col_names = FALSE)
A32A <- read_excel(archivo, sheet = "A32", range = "X45", col_names = FALSE) + read_excel(archivo, sheet = "A32", range = "Y45", col_names = FALSE)
A32NSP <- read_excel(archivo, sheet = "A32", range = "V45", col_names = FALSE) + read_excel(archivo, sheet = "A32", range = "W45", col_names = FALSE)

A32 <- data.frame("Fecha"= as.Date(fecha_mes), "Atenciones_Remotas" = A32T[,1], "Atenciones_Remotas_Nuevas" = A32N[,1], "Altas" = A32A[,1], "NSP"= A32NSP[,1])
colnames(A32)[2] <- "Atenciones_Remotas"

A32 <- rbind(A32,A32B)


# A07 ---------------------------------------------------------------------
A07M <- read_xlsx(archivo,
                  na = " ",
                  col_names = TRUE,
                  range = "A07!A10:AX70")
A07M <- mutate_all(A07M, ~replace(., is.na(.), 0))
A07M$Fecha <- fecha_mes
colnames(A07M)[1] <- "Especialidad"
colnames(A07M)[2] <- "Total"
colnames(A07M)[36] <- "IC_Menor_15"
colnames(A07M)[37] <- "IC_Mayor_15"
colnames(A07M)[38] <- "Alta_Menor_15"
colnames(A07M)[39] <- "Alta_Mayor_15"
colnames(A07M)[35] <- "C_Abreviada"
colnames(A07M)[3] <- "0 a 4 años"
colnames(A07M)[4] <- "5 a 9 años"
colnames(A07M)[5] <- "10 a 14 años"
colnames(A07M)[6] <- "15 a 19 años"
A07M$"19 años o más" <- A07M$`20 - 24`+A07M$`25 - 29`+A07M$`30 - 34`+A07M$`35 - 39`+A07M$`40 - 44`+A07M$`45 - 49`+
  A07M$`50 - 54`+A07M$`55 - 59`+A07M$`60 - 64`+A07M$`65 - 69`+A07M$`70 - 74`+A07M$`75 - 79`+A07M$`80 y mas`
A07M$"Consulta_Nueva" <- A07M$TOTAL...23+A07M$TOTAL...27
A07M$"CN_Origen_APS" <- A07M$APS...24+A07M$APS...28
A07M$"CN_Origen_Hospital" <- A07M$`CAE/CDT/CRS/HOSPITALIZACIÓN...25`+A07M$`CAE/CDT/CRS/HOSPITALIZACIÓN...29`
A07M$"CN_Origen_Urgencia" <- A07M$URGENCIA...26+A07M$URGENCIA...30
A07M$"IC" <- A07M$IC_Menor_15+A07M$IC_Mayor_15
A07M$"Alta" <- A07M$Alta_Menor_15+A07M$Alta_Mayor_15
A07M$"NSP" <- A07M$NUEVAS + A07M$CONTROLES
A07M$`0 - 4` <- NULL
A07M$`5 - 9` <- NULL
A07M$`10 - 14` <- NULL
A07M$`15 - 19` <- NULL
A07M$`20 - 24` <- NULL
A07M$`25 - 29` <- NULL
A07M$`30 - 34` <- NULL
A07M$`35 - 39` <- NULL
A07M$`40 - 44` <- NULL
A07M$`45 - 49` <- NULL
A07M$`50 - 54` <- NULL
A07M$`55 - 59` <- NULL
A07M$`60 - 64` <- NULL
A07M$`65 - 69` <- NULL
A07M$`70 - 74` <- NULL
A07M$`75 - 79` <- NULL
A07M$`80 y mas` <- NULL
A07M$...2 <- NULL
A07M$TOTAL...23 <- NULL
A07M$TOTAL...27 <- NULL
A07M$APS...24 <- NULL
A07M$APS...28 <- NULL
A07M$`CAE/CDT/CRS/HOSPITALIZACIÓN...25` <- NULL
A07M$`CAE/CDT/CRS/HOSPITALIZACIÓN...29` <- NULL
A07M$URGENCIA...26 <- NULL
A07M$URGENCIA...30 <- NULL
A07M$...20 <- NULL
colnames(A07M)[11] <- "NSP_Nuevas"
colnames(A07M)[12] <- "NSP_Controles"
A07M$...42 <- NULL
A07M$...43 <- NULL
colnames(A07M)[9] <- "Pertinencia_Protocolo_Referencia"
colnames(A07M)[10] <- "Pertinencia_Tiempo_Establecido"

#Uso select para indicar que columnas me interesan
A07M <- A07M %>% select(Fecha, Especialidad, Total,  Consulta_Nueva, CN_Origen_APS,CN_Origen_Hospital, CN_Origen_Urgencia,
                        NSP, NSP_Nuevas, NSP_Controles, IC, Alta, Pertinencia_Protocolo_Referencia, Pertinencia_Tiempo_Establecido,
                        Hombres, Mujeres, C_Abreviada, '0 a 4 años', '5 a 9 años', '10 a 14 años', '15 a 19 años', '19 años o más')


A07CNM <- A07M %>% select(Fecha, Especialidad, CN_Origen_APS,CN_Origen_Hospital, CN_Origen_Urgencia,
                        NSP_Nuevas, NSP_Controles, Hombres, Mujeres,
                        '0 a 4 años', '5 a 9 años', '10 a 14 años', '15 a 19 años', '19 años o más')

# A07 C -----------------------------------------------------------------
A07C <- read_xlsx(archivo, na = " ",col_names = FALSE,
                   range = "A07!A94:AQ103") %>% fill(...1)
A07C <- mutate_all(A07C, ~replace(., is.na(.), 0))
colnames(A07C)[1] <- "Profesional"
colnames(A07C)[3] <- "Total"
colnames(A07C)[4] <- "Hombres"
colnames(A07C)[5] <- "Mujeres"
A07C$Fecha <- fecha_mes
A07C$Actividad <- "CAE"
A07C$REM <- "A07"
A07C$"Migrantes" <- A07C$...43
A07C$"Pueblos Originarios" <- A07C$...42
A07C$"0 a 4 años" <- A07C$...6+A07C$...7
A07C$"5 a 9 años" <- A07C$...8+A07C$...9
A07C$"10 a 14 años" <- A07C$...10+A07C$...11
A07C$"15 años o más" <-A07C$...12+A07C$...13+A07C$...14+A07C$...15+A07C$...16+A07C$...17+A07C$...18+A07C$...19+
  A07C$...20+A07C$...21+A07C$...22+A07C$...23+A07C$...24+A07C$...25+A07C$...26+A07C$...27+A07C$...28+A07C$...29+
  A07C$...30+A07C$...31+A07C$...32+A07C$...33+A07C$...34+A07C$...35+A07C$...36+A07C$...37+A07C$...38+A07C$...39
A07C <- A07C %>%select(Fecha,REM, Actividad, Profesional, Total, Hombres, Mujeres, `Pueblos Originarios`, Migrantes, '0 a 4 años',
                       '5 a 9 años', '10 a 14 años', '15 años o más')
# A06 ---------------------------------------------------------------------
A061 <- read_xlsx(archivo, na = " ",col_names = FALSE,
                  range = "A06!A13:AR22") %>% fill(...1)
A061 <- mutate_all(A061, ~replace(., is.na(.), 0))
A062 <- read_xlsx(archivo, na = " ",col_names = FALSE,
                  range = "A06!A25:AR25")
A062 <- mutate_all(A062, ~replace(., is.na(.), 0))
A063 <- read_xlsx(archivo, na = " ",col_names = FALSE,
                  range = "A06!A26:AR27") %>% fill(...1)
A063 <- mutate_all(A063, ~replace(., is.na(.), 0))
A06M <- rbind(A061, A062, A063)

colnames(A06M)[1] <- "Actividad"
colnames(A06M)[2] <- "Profesional"
colnames(A06M)[3] <- "Total"
colnames(A06M)[4] <- "Hombres"
colnames(A06M)[5] <- "Mujeres"
A06M$Fecha <- fecha_mes
A06M$REM <- "A06"
A06M$"Migrantes" <- A06M$...43
A06M$"Pueblos Originarios" <- A06M$...42
A06M$"0 a 4 años" <- A06M$...6+A06M$...7
A06M$"5 a 9 años" <- A06M$...8+A06M$...9
A06M$"10 a 14 años" <- A06M$...10+A06M$...11
A06M$"15 años o más" <-A06M$...12+A06M$...13+A06M$...14+A06M$...15+A06M$...16+A06M$...17+A06M$...18+A06M$...19+
  A06M$...20+A06M$...21+A06M$...22+A06M$...23+A06M$...24+A06M$...25+A06M$...26+A06M$...27+A06M$...28+A06M$...29+
  A06M$...30+A06M$...31+A06M$...32+A06M$...33+A06M$...34+A06M$...35+A06M$...36+A06M$...37+A06M$...38+A06M$...39
A06M <- A06M %>%select(Fecha, REM, Actividad, Profesional, Total, Hombres, Mujeres, `Pueblos Originarios`, Migrantes, '0 a 4 años',
                     '5 a 9 años', '10 a 14 años', '15 años o más')

A06M <- rbind(A06M, A07C)

# A30 ---------------------------------------------------------------------
A30M <- read_xlsx(archivo,
                  na = " ",
                  col_names = TRUE,
                  range = "A30!A17:AA77")
A30M$Fecha <- fecha_mes
A30M <- mutate_all(A30M, ~replace(., is.na(.), 0))
colnames(A30M)[1] <- "Especialidad"
colnames(A30M)[2] <- "F2"
colnames(A30M)[3] <- "F3"
colnames(A30M)[4] <- "F4"
colnames(A30M)[5] <- "F5"
A30M$"Telemedicina_Nueva" <- A30M$F2+A30M$F3+A30M$F4+A30M$F5
colnames(A30M)[6] <- "F6"
colnames(A30M)[7] <- "F7"
colnames(A30M)[8] <- "F8"
colnames(A30M)[9] <- "F9"
A30M$"Telemedicina_Control" <- A30M$F6+A30M$F7+A30M$F8+A30M$F9
colnames(A30M)[17] <- "F17"
colnames(A30M)[18] <- "F18"
colnames(A30M)[19] <- "F19"
colnames(A30M)[20] <- "F20"
A30M$"Telemedicina_Hospitalizados" <- A30M$F17+A30M$F18+A30M$F19+A30M$F20
A30M <- A30M %>% select(Fecha,Especialidad, Telemedicina_Nueva,Telemedicina_Control,Telemedicina_Hospitalizados)

# A28 ---------------------------------------------------------------------
A28M <- read_xlsx(archivo, na = " ",col_names = FALSE,
                  range = "A28!A168:B192")
A28M <- mutate_all(A28M, ~replace(., is.na(.), 0))
A28M$Seccion <- "B1"
A28M$Nombre_Seccion <- "INGRESOS Y EGRESOS  AL PROGRAMA DE REHABILITACIÓN INTEGRAL"
A28M$Variable <- "Programa"
colnames(A28M)[1] <- "Item"
colnames(A28M)[2] <- "Total"

A28K <- read_xlsx(archivo, na = " ",col_names = FALSE,
                  range = "A28!A194:B198")
A28K <- mutate_all(A28K, ~replace(., is.na(.), 0))
A28K$Seccion <- "B1"
A28K$Nombre_Seccion <- "INGRESOS Y EGRESOS  AL PROGRAMA DE REHABILITACIÓN INTEGRAL"
A28K$Variable <- "Egresos"
colnames(A28K)[1] <- "Item"
colnames(A28K)[2] <- "Total"
A28M <- rbind(A28M, A28K)

A28K <- read_xlsx(archivo, na = " ",col_names = FALSE,
                  range = "A28!A203:B208")
A28K <- mutate_all(A28K, ~replace(., is.na(.), 0))
A28K$Seccion <- "B2"
A28K$Nombre_Seccion <- "Evaluación Inicial"
A28K$Variable <- "Profesional"
colnames(A28K)[1] <- "Item"
colnames(A28K)[2] <- "Total"
A28M <- rbind(A28M, A28K)

A28K <- read_xlsx(archivo, na = " ",col_names = FALSE,
                  range = "A28!A212:B216")
A28K <- mutate_all(A28K, ~replace(., is.na(.), 0))
A28K$Seccion <- "B3"
A28K$Nombre_Seccion <- "Evaluación Intermedia"
A28K$Variable <- "Profesional"
colnames(A28K)[1] <- "Item"
colnames(A28K)[2] <- "Total"
A28M <- rbind(A28M, A28K)

A28K <- read_xlsx(archivo, na = " ",col_names = FALSE,
                  range = "A28!A220:B223")
A28K <- mutate_all(A28K, ~replace(., is.na(.), 0))
A28K$Seccion <- "B4"
A28K$Nombre_Seccion <- "Sesiones de Rehabilitación"
A28K$Variable <- "Profesional"
colnames(A28K)[1] <- "Item"
colnames(A28K)[2] <- "Total"
A28M <- rbind(A28M, A28K)

A28K <- read_xlsx(archivo, na = " ",col_names = FALSE,
                  range = "A28!A227:B229")
A28K <- mutate_all(A28K, ~replace(., is.na(.), 0))
A28K$Seccion <- "B5"
A28K$Nombre_Seccion <- "Derivaciones"
A28K$Variable <- "Establecimiento"
colnames(A28K)[1] <- "Item"
colnames(A28K)[2] <- "Total"
A28M <- rbind(A28M, A28K)

A28K <- read_xlsx(archivo, na = " ",col_names = FALSE,
                  range = "A28!A232:B265")
A28K <- mutate_all(A28K, ~replace(., is.na(.), 0))
A28K$Seccion <- "B6"
A28K$Nombre_Seccion <- "Sesiones de Rehabilitación"
A28K$Variable <- "Tipo"
colnames(A28K)[1] <- "Item"
colnames(A28K)[2] <- "Total"
A28M <- rbind(A28M, A28K)

A28K <- read_xlsx(archivo, na = " ",col_names = FALSE,
                  range = "A28!A272:B272")
A28K <- mutate_all(A28K, ~replace(., is.na(.), 0))
A28K$Seccion <- "C1"
A28K$Nombre_Seccion <- "N° Personas con ayudas técnicas"
A28K$Variable <- "Ayudas técnicas"
colnames(A28K)[1] <- "Item"
colnames(A28K)[2] <- "Total"
A28M <- rbind(A28M, A28K)

A28K <- read_xlsx(archivo, na = " ",col_names = FALSE,
                  range = "A28!A276:B297")
A28K <- mutate_all(A28K, ~replace(., is.na(.), 0))
A28K$Seccion <- "C2"
A28K$Nombre_Seccion <- "Ayudas técnicas por tipo"
A28K$Variable <- "Ayudas técnicas"
colnames(A28K)[1] <- "Item"
colnames(A28K)[2] <- "Total"
A28M <- rbind(A28M, A28K)

A28K <- read_xlsx(archivo, na = " ",col_names = FALSE,
                  range = "A28!A301:B323")
A28K <- mutate_all(A28K, ~replace(., is.na(.), 0))
A28K$Seccion <- "C3"
A28K$Nombre_Seccion <- "Ayudas técnicas por condición de salud"
A28K$Variable <- "Ayudas técnicas"
colnames(A28K)[1] <- "Item"
colnames(A28K)[2] <- "Total"
A28M <- rbind(A28M, A28K)

A28K <- read_xlsx(archivo, na = " ",col_names = FALSE,
                  range = "A28!B349:C352")
A28K <- mutate_all(A28K, ~replace(., is.na(.), 0))
A28K$Seccion <- "D2"
A28K$Nombre_Seccion <- "Evaluación intermedia remota"
A28K$Variable <- "Profesional"
colnames(A28K)[1] <- "Item"
colnames(A28K)[2] <- "Total"
A28M <- rbind(A28M, A28K)

A28K <- read_xlsx(archivo, na = " ",col_names = FALSE,
                  range = "A28!B345:C348")
A28K <- mutate_all(A28K, ~replace(., is.na(.), 0))
A28K$Seccion <- "D2"
A28K$Nombre_Seccion <- "Evaluación intermedia remota"
A28K$Variable <- "Profesional"
colnames(A28K)[1] <- "Item"
colnames(A28K)[2] <- "Total"
A28M <- rbind(A28M, A28K)

A28K <- read_xlsx(archivo, na = " ",col_names = FALSE,
                  range = "A28!B349:C352")
A28K <- mutate_all(A28K, ~replace(., is.na(.), 0))
A28K$Seccion <- "D2"
A28K$Nombre_Seccion <- "Sesión de rehabilitación vía remota (telerehabilitación)"
A28K$Variable <- "Profesional"
colnames(A28K)[1] <- "Item"
colnames(A28K)[2] <- "Total"
A28M <- rbind(A28M, A28K)

A28K <- read_xlsx(archivo, na = " ",col_names = FALSE,
                  range = "A28!B353:C356")
A28K <- mutate_all(A28K, ~replace(., is.na(.), 0))
A28K$Seccion <- "D2"
A28K$Nombre_Seccion <- "Educación remota a usuario y/o cuidador"
A28K$Variable <- "Profesional"
colnames(A28K)[1] <- "Item"
colnames(A28K)[2] <- "Total"
A28M <- rbind(A28M, A28K)
A28M$Fecha <- fecha_mes

A28M <- A28M %>% select(Fecha,Seccion, Nombre_Seccion, Variable, Item, Total)

# A08 ---------------------------------------------------------------------
A08M <- read_xlsx(archivo, na = " ",col_names = FALSE,
                  range = "A08!A12:AS15")
A08M <- mutate_all(A08M, ~replace(., is.na(.), 0))
colnames(A08M)[1] <- "Tipo de Atención"
colnames(A08M)[2] <- "Total"
colnames(A08M)[3] <- "Hombres"
colnames(A08M)[4] <- "Mujeres"
A08M$"0 a 4 años" <- A08M$...5+A08M$...6
A08M$"5 a 9 años" <- A08M$...7+A08M$...8
A08M$"10 a 14 años" <- A08M$...9+A08M$...10
A08M$"15 años o más" <- A08M$...11+A08M$...12+A08M$...13+A08M$...14+A08M$...15+A08M$...16+A08M$...17+A08M$...18+A08M$...19+A08M$...20+
  A08M$...21+A08M$...22+A08M$...23+A08M$...24+A08M$...25+A08M$...26+A08M$...27+A08M$...28+A08M$...29+A08M$...30+A08M$...31+A08M$...32+
  A08M$...33+A08M$...34+A08M$...35+A08M$...36+A08M$...37+A08M$...38
A08M$"SAPU_SAR_SUR" <- A08M$...40
A08M$"Hospital_Baja_Complejidad" <- A08M$...41
A08M$"Hospital_Mediana-Alta Complejidad" <- A08M$...42
A08M$"Otros_Establecimientos_de_la_Red" <- A08M$...43
A08M$"Establecimientos de otra Red" <- A08M$...44
A08M$Fecha <- fecha_mes
A08M <- A08M %>% select(Fecha, `Tipo de Atención`, Total,  Hombres, Mujeres, SAPU_SAR_SUR,Hospital_Baja_Complejidad,
                        `Hospital_Mediana-Alta Complejidad`, Otros_Establecimientos_de_la_Red, `Establecimientos de otra Red`,
                         '0 a 4 años', '5 a 9 años', '10 a 14 años', '15 años o más')

A08M2 <- read_xlsx(archivo, na = " ",col_names = FALSE,
                  range = "A08!A58:B63")
A08M2 <- mutate_all(A08M2, ~replace(., is.na(.), 0))
colnames(A08M2)[1] <- "Categorización"
colnames(A08M2)[2] <- "Total"
A08M2$Fecha <- fecha_mes
A08M2 <- A08M2 %>%select(Fecha, Categorización, Total)

A08M3 <- read_xlsx(archivo, na = " ",col_names = FALSE,
                   range = "A08!A67:B85")
A08M3 <- mutate_all(A08M3, ~replace(., is.na(.), 0))
colnames(A08M3)[1] <- "Especialidades"
colnames(A08M3)[2] <- "Total"
A08M3$Fecha <- fecha_mes
A08M3 <- A08M3 %>%select(Fecha, Especialidades, Total)

A08M4 <- read_xlsx(archivo, na = " ",col_names = FALSE,
                   range = "A08!A92:c98")
A08M4 <- mutate_all(A08M4, ~replace(., is.na(.), 0))
colnames(A08M4)[1] <- "Item"
colnames(A08M4)[2] <- 'Categoria pacientes'
colnames(A08M4)[3] <- "Total"
A08M4$Fecha <- fecha_mes
A08M4 <- A08M4 %>%select(Fecha, Item,'Categoria pacientes', Total)
A08M4$Item <- ifelse( A08M4$Item == 0, 'PACIENTES QUE INGRESAN A CAMA HOSPITALARIA SEGÚN TIEMPO DE DEMORA AL INGRESO', A08M4$Item)
A08M4$`Categoria pacientes` <- ifelse( A08M4$`Categoria pacientes` == 0, A08M4$Item, A08M4$`Categoria pacientes`)

A08M5 <- read_xlsx(archivo, na = " ",col_names = FALSE,
                   range = "A08!A108:B110")
A08M5 <- mutate_all(A08M5, ~replace(., is.na(.), 0))
colnames(A08M5)[1] <- "Tipo de pacientes"
colnames(A08M5)[2] <- "Total"
A08M5$Fecha <- fecha_mes
A08M5 <- A08M5 %>%select(Fecha, 'Tipo de pacientes', Total)

A08M6 <- read_xlsx(archivo, na = " ",col_names = FALSE,
                   range = "A08!A191:B195")
A08M6 <- mutate_all(A08M6, ~replace(., is.na(.), 0))
colnames(A08M6)[1] <- "Animal mordedor"
colnames(A08M6)[2] <- "Total"
A08M6$Fecha <- fecha_mes
A08M6 <- A08M6 %>%select(Fecha, 'Animal mordedor', Total)

# A09 ---------------------------------------------------------------------
A09M1 <- read_xlsx(archivo, na = " ",col_names = FALSE,
                   range = "A09!A19:D35")
A09M1 <- mutate_all(A09M1, ~replace(., is.na(.), 0))
colnames(A09M1)[1] <- "Tipo de actividad"
colnames(A09M1)[4] <- "Total"
A09M1$Fecha <- fecha_mes
A09M1 <- A09M1 %>%select(Fecha, 'Tipo de actividad', Total)

A09M2 <- read_xlsx(archivo, na = " ",col_names = FALSE,
                   range = "A09!A91:D144")
A09M2 <- mutate_all(A09M2, ~replace(., is.na(.), 0))
colnames(A09M2)[1] <- "ACTIVIDADES DE ESPECIALIDADES"
colnames(A09M2)[4] <- "Total"
A09M2$Fecha <- fecha_mes
A09M2 <- A09M2 %>%select(Fecha, 'ACTIVIDADES DE ESPECIALIDADES', Total)

A09M3 <- read_xlsx(archivo, na = " ",col_names = FALSE,
                   range = "A09!A221:D306") %>% fill(...1) #el fill lo uso para repetir el valor en las celdas combinadas
A09M3 <- mutate_all(A09M3, ~replace(., is.na(.), 0))
colnames(A09M3)[1] <- "ESPECIALIDAD"
colnames(A09M3)[2] <- "TIPO DE INGRESO O EGRESO"
colnames(A09M3)[4] <- "Total"
A09M3$Fecha <- fecha_mes
A09M3 <- A09M3 %>% select(Fecha, ESPECIALIDAD, 'TIPO DE INGRESO O EGRESO', Total)

# A19b --------------------------------------------------------------------
A019bM <- read_xlsx(archivo, na = " ",col_names = FALSE,
                  range = "A19b!A12:I24")
A019bM <- mutate_all(A019bM, ~replace(., is.na(.), 0))
A019bM$"Tipo consulta" <- "RECLAMO"
A019bM1 <- read_xlsx(archivo, na = " ",col_names = FALSE,
                    range = "A19b!A25:I29")
A019bM1 <- mutate_all(A019bM1, ~replace(., is.na(.), 0))
A019bM1$"Tipo consulta" <- A019bM1$...1

A019bM <- rbind(A019bM, A019bM1)

colnames(A019bM)[1] <- "Tipo atención"
colnames(A019bM)[2] <- "Total"
colnames(A019bM)[3] <- "Hombres"
colnames(A019bM)[4] <- "Mujeres"
colnames(A019bM)[5] <- "RESPUESTAS DEL MES DENTRO DE PLAZOS LEGALES ( 15 DIAS HÁBILES) Reclamos generados en el mes"
colnames(A019bM)[6] <- "RESPUESTAS DEL MES DENTRO DE PLAZOS LEGALES ( 15 DIAS HÁBILES) Reclamos generados en el mes anterior RECLAMOS RESPONDIDOS FUERA DE PLAZOS LEGALES"
colnames(A019bM)[7] <- "RECLAMOS RESPONDIDOS FUERA DE PLAZOS LEGALES"
colnames(A019bM)[8] <- "RECLAMOS PENDIENTES Respuestas pendientes dentro del plazo legal"
colnames(A019bM)[9] <- "RECLAMOS PENDIENTES Respuestas pendientes fuera del plazo legal"
A019bM$Fecha <- fecha_mes
A019bM <- A019bM %>%select(Fecha, `Tipo consulta`, `Tipo atención`, Total, Hombres, Mujeres,
                           `RESPUESTAS DEL MES DENTRO DE PLAZOS LEGALES ( 15 DIAS HÁBILES) Reclamos generados en el mes`,
                           `RESPUESTAS DEL MES DENTRO DE PLAZOS LEGALES ( 15 DIAS HÁBILES) Reclamos generados en el mes anterior RECLAMOS RESPONDIDOS FUERA DE PLAZOS LEGALES`,
                           `RECLAMOS RESPONDIDOS FUERA DE PLAZOS LEGALES`,
                           `RECLAMOS PENDIENTES Respuestas pendientes dentro del plazo legal` ,
                           `RECLAMOS PENDIENTES Respuestas pendientes fuera del plazo legal`)

# A21 ---------------------------------------------------------------------
#Ocupación
A211M <- read_xlsx(archivo, na = " ",col_names = FALSE,
                  range = "A21!A13:AB16")
A211M <- mutate_all(A211M, ~replace(., is.na(.), 0))
A211M$Fecha <- fecha_mes
colnames(A211M)[1] <- "TIPO DE QUIRÓFANOS"
colnames(A211M)[2] <- "NÙMERO DE QUIRÓFANOS EN DOTACIÓN"
colnames(A211M)[3] <- "PROMEDIO MENSUAL DE QUIRÓFANOS HABILITADOS"
colnames(A211M)[4] <- "PROMEDIO MENSUAL  DE QUIRÓFANOS EN TRABAJO"
colnames(A211M)[5] <- "TOTAL  DE HORAS MENSUALES DE QUIRÓFANOS HABILITADOS (HORARIO HÁBIL)"
colnames(A211M)[6] <- "TOTAL DE HORAS MENSUALES DE QUIRÓFANOS EN TRABAJO"
colnames(A211M)[7] <- "HORAS MENSUALES PROGRAMADAS DE TABLA QUIRÚRGICA DE QUIRÓFANOS EN TRABAJO"
colnames(A211M)[11] <- "HORAS MENSUALES OCUPADAS DE QUIRÓFANOS EN TRABAJO HORARIO HÁBIL"
colnames(A211M)[16] <- "HORAS MENSUALES OCUPADAS DE QUIRÓFANOS EN TRABAJO HORARIO INHÁBIL DE LUNES A VIERNES"
colnames(A211M)[21] <- "HORAS MENSUALES OCUPADAS DE QUIRÓFANOS EN TRABAJO HORARIO SÁBADO, DOMINGO Y FESTIVO"
colnames(A211M)[27] <- "TOTAL HORAS MENSUALES OCUPADAS DE QUIRÓFANOS EN TRABAJO (Cirugia Menor)"
colnames(A211M)[28] <- "TOTAL HORAS MENSUALES OCUPADAS DE QUIRÓFANOS EN TRABAJO (Procedimientos)"
A211M <- A211M %>%select(Fecha, 'TIPO DE QUIRÓFANOS', 'NÙMERO DE QUIRÓFANOS EN DOTACIÓN', 'PROMEDIO MENSUAL DE QUIRÓFANOS HABILITADOS','PROMEDIO MENSUAL  DE QUIRÓFANOS EN TRABAJO','TOTAL  DE HORAS MENSUALES DE QUIRÓFANOS HABILITADOS (HORARIO HÁBIL)',
                       'TOTAL DE HORAS MENSUALES DE QUIRÓFANOS EN TRABAJO','HORAS MENSUALES PROGRAMADAS DE TABLA QUIRÚRGICA DE QUIRÓFANOS EN TRABAJO', 'HORAS MENSUALES OCUPADAS DE QUIRÓFANOS EN TRABAJO HORARIO HÁBIL', 'HORAS MENSUALES OCUPADAS DE QUIRÓFANOS EN TRABAJO HORARIO INHÁBIL DE LUNES A VIERNES','HORAS MENSUALES OCUPADAS DE QUIRÓFANOS EN TRABAJO HORARIO SÁBADO, DOMINGO Y FESTIVO','TOTAL HORAS MENSUALES OCUPADAS DE QUIRÓFANOS EN TRABAJO (Cirugia Menor)','TOTAL HORAS MENSUALES OCUPADAS DE QUIRÓFANOS EN TRABAJO (Procedimientos)')

#Procedimientos
A212M <- read_xlsx(archivo, na = " ",col_names = FALSE,
                   range = "A21!A19:G23")
A212M <- mutate_all(A212M, ~replace(., is.na(.), 0))
A212M$Fecha <- fecha_mes
colnames(A212M)[1] <- "Componentes"
colnames(A212M)[2] <- "Total"
colnames(A212M)[3] <- "Quimioterapia"
colnames(A212M)[4] <- "Hemodialisis"
colnames(A212M)[5] <- "CMA"
colnames(A212M)[6] <- "Coronariografia"
colnames(A212M)[7] <- "Otras"

#Suspensiones
A213M <- read_xlsx(archivo, na = " ",col_names = FALSE,
                   range = "A21!A76:I87")
A213M <- mutate_all(A213M, ~replace(., is.na(.), 0))
A213M$Fecha <- fecha_mes
colnames(A213M)[1] <- "Especialidad"
A213M$'Dias de estada prequirurgicos' <- A213M$...2+A213M$...3
A213M$'Pacientes Intervenidos' <- A213M$...4+A213M$...5
A213M$'Pacientes Programados' <- A213M$...6+A213M$...7
A213M$'Pacientes Suspendidos' <- A213M$...8+A213M$...9
A213M <- A213M %>%select(Fecha,Especialidad, `Dias de estada prequirurgicos`,
                         `Pacientes Intervenidos`,`Pacientes Programados`, `Pacientes Suspendidos`)
#Causas Suspensiones
A214M <- read_xlsx(archivo, na = " ",col_names = FALSE,
                   range = "A21!A92:B98")
A214M <- mutate_all(A214M, ~replace(., is.na(.), 0))
A214M$Fecha <- fecha_mes
colnames(A214M)[1] <- "Causa atribuible a"
colnames(A214M)[2] <- "Cantidad"
A214M <- A214M %>%select(Fecha, `Causa atribuible a`, Cantidad)



# A04 ---------------------------------------------------------------------
A041M <- read_xlsx(archivo, na = " ",col_names = FALSE,
                   range = "A04!A90:F94")
A041M <- mutate_all(A041M, ~replace(., is.na(.), 0))
A041M$Servicio_Farmaceutico <- "Atención Farmaceútica"
A042M <- read_xlsx(archivo, na = " ",col_names = FALSE,
                   range = "A04!A96:F98")
A042M <- mutate_all(A042M, ~replace(., is.na(.), 0))
A042M$Servicio_Farmaceutico <- "Farmacovigilancia"
A041M <- rbind(A041M, A042M)
A041M$Fecha <- fecha_mes
colnames(A041M)[1] <- "Componente"
colnames(A041M)[2] <- "Total"
colnames(A041M)[3] <- "Atención Abierta"
colnames(A041M)[4] <- "Atención Cerrada"
colnames(A041M)[5] <- "Servicios Farmaceúticos realizados en establecimientos de Salud"
colnames(A041M)[6] <- "Servicios Farmaceúticos realizados en Domicilio"
A041M <- A041M %>%select(Fecha, Servicio_Farmaceutico,Componente,Total,`Atención Abierta`,`Atención Cerrada`,
                        `Servicios Farmaceúticos realizados en establecimientos de Salud`,
                        `Servicios Farmaceúticos realizados en Domicilio`)

A043M <- read_xlsx(archivo, na = " ",col_names = FALSE,
                   range = "A04!A103:J105")
A043M <- mutate_all(A043M, ~replace(., is.na(.), 0))
A043M$Fecha <- fecha_mes
colnames(A043M)[1] <- "Tipo de receta"
colnames(A043M)[2] <- "Despacho Total"
colnames(A043M)[3] <- "Despacho Parcial"
colnames(A043M)[4] <- "Despachos en Centros de Salud"
colnames(A043M)[5] <- "Despacho en Domicilio"
colnames(A043M)[6] <- "Prescripciones Emitidas"
colnames(A043M)[7] <- "Prescripciones Rechazadas"
colnames(A043M)[8] <- "Prescripciones en Centros de Salud"
colnames(A043M)[9] <- "Prescripciones en Domicilio"
colnames(A043M)[10] <- "Recetas despachadas con Oportunidad"


# Lee las BBDD ------------------------------------------------------------
A07 <- read_excel(A07BBDD)
A07$Fecha=as.character(A07$Fecha)
A07CN <- read_excel(A07CNBBDD)
A07CN$Fecha=as.character(A07CN$Fecha)
A30 <- read_excel(A30BBDD)
A30$Fecha=as.character(A30$Fecha)
A28 <- read_excel(A28BBDD)
A28 <- A28 %>% select(Fecha, Seccion, Nombre_Seccion, Variable, Item, Total)
A08 <- read_excel(A08BBDD01)
A082 <- read_excel(A08BBDD02)
A083 <- read_excel(A08BBDD03)
A084 <- read_excel(A08BBDD04)
A085 <- read_excel(A08BBDD05)
A086 <- read_excel(A08BBDD06)
A091 <- read_excel(A09BBDD01)
A092 <- read_excel(A09BBDD02)
A093 <- read_excel(A09BBDD03)
A06 <- read_excel(A06BBDD)
A19b <- read_excel(A19bBBDD)
A211 <- read_excel(A211BBDD)
A212 <- read_excel(A212BBDD)
A213 <- read_excel(A213BBDD)
A214 <- read_excel(A214BBDD)
A041 <- read_excel(A041BBDD)
A043 <- read_excel(A043BBDD)


# junta la informacion con el REM mensual   -------------------------------
A07 <- rbind(A07, A07M)
A07CN <- rbind(A07CN, A07CNM)
A30 <- rbind(A30, A30M)
A28 <- rbind(A28, A28M)
A08 <- rbind(A08, A08M)
A082 <- rbind(A082, A08M2)
A083 <- rbind(A083, A08M3)
A084 <- rbind(A084, A08M4)
A085 <- rbind(A085, A08M5)
A086 <- rbind(A086, A08M6)
A091 <- rbind(A091, A09M1)
A092 <- rbind(A092, A09M2)
A093 <- rbind(A093, A09M3)
A06 <- rbind(A06, A06M)
A19b <- rbind(A19b, A019bM)
A211 <- rbind(A211, A211M)
A212 <- rbind(A212, A212M)
A213 <- rbind(A213, A213M)
A214 <- rbind(A214, A214M)
A041 <- rbind(A041, A041M)
A043 <- rbind(A043, A043M)


# da formato fecha a la variable fecha ------------------------------------
A07$Fecha=as.Date(A07$Fecha)
A07CN$Fecha=as.Date(A07CN$Fecha)
A30$Fecha=as.Date(A30$Fecha)
A28$Fecha=as.Date(A28$Fecha)
A091$Fecha=as.Date(A091$Fecha)
A092$Fecha=as.Date(A092$Fecha)
A093$Fecha=as.Date(A093$Fecha)
A08$Fecha=as.Date(A08$Fecha)
A082$Fecha=as.Date(A082$Fecha)
A083$Fecha=as.Date(A083$Fecha)
A084$Fecha=as.Date(A084$Fecha)
A085$Fecha=as.Date(A085$Fecha)
A086$Fecha=as.Date(A086$Fecha)
A06$Fecha=as.Date(A06$Fecha)
A19b$Fecha=as.Date(A19b$Fecha)
A211$Fecha=as.Date(A211$Fecha)
A212$Fecha=as.Date(A212$Fecha)
A213$Fecha=as.Date(A213$Fecha)
A214$Fecha=as.Date(A214$Fecha)
A041$Fecha=as.Date(A041$Fecha)
A043$Fecha=as.Date(A043$Fecha)
#
#
# Graba las BBDD en el archivo excel --------------------------------------
openxlsx::write.xlsx(A07, A07BBDD, colNames = TRUE, sheetName = "A07", overwrite = T)
openxlsx::write.xlsx(A07CN, A07CNBBDD, colNames = TRUE, sheetName = "A07", overwrite = T)
openxlsx::write.xlsx(A30, A30BBDD, colNames = TRUE, sheetName = "A30", overwrite = T)
openxlsx::write.xlsx(A32, A32BBDD, colNames = TRUE, sheetName = "A32", overwrite = T)
openxlsx::write.xlsx(A28, A28BBDD, colNames = TRUE, sheetName = "A28", overwrite = T)
openxlsx::write.xlsx(A08, A08BBDD01, colNames = TRUE, sheetName = "A08_1", overwrite = T)
openxlsx::write.xlsx(A082, A08BBDD02, colNames = TRUE, sheetName = "A08_2", overwrite = T)
openxlsx::write.xlsx(A083, A08BBDD03, colNames = TRUE, sheetName = "A08_3", overwrite = T)
openxlsx::write.xlsx(A084, A08BBDD04, colNames = TRUE, sheetName = "A08_4", overwrite = T)
openxlsx::write.xlsx(A085, A08BBDD05, colNames = TRUE, sheetName = "A08_5", overwrite = T)
openxlsx::write.xlsx(A086, A08BBDD06, colNames = TRUE, sheetName = "A08_6", overwrite = T)
openxlsx::write.xlsx(A091, A09BBDD01, colNames = TRUE, sheetName = "A09_1", overwrite = T)
openxlsx::write.xlsx(A092, A09BBDD02, colNames = TRUE, sheetName = "A09_2", overwrite = T)
openxlsx::write.xlsx(A093, A09BBDD03, colNames = TRUE, sheetName = "A09_3", overwrite = T)
openxlsx::write.xlsx(A06, A06BBDD, colNames = TRUE, sheetName = "A06", overwrite = T)
openxlsx::write.xlsx(A19b, A19bBBDD, colNames = TRUE, sheetName = "A19b", overwrite = T)
openxlsx::write.xlsx(A211, A211BBDD, colNames = TRUE, sheetName = "A21_1", overwrite = T)
openxlsx::write.xlsx(A212, A212BBDD, colNames = TRUE, sheetName = "A21_2", overwrite = T)
openxlsx::write.xlsx(A213, A213BBDD, colNames = TRUE, sheetName = "A21_3", overwrite = T)
openxlsx::write.xlsx(A214, A214BBDD, colNames = TRUE, sheetName = "A21_4", overwrite = T)
openxlsx::write.xlsx(A041, A041BBDD, colNames = TRUE, sheetName = "A04_1", overwrite = T)
openxlsx::write.xlsx(A043, A043BBDD, colNames = TRUE, sheetName = "A04_3", overwrite = T)

}

