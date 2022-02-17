
# Librerias ---------------------------------------------------------------
library(tidyverse)
library(readxl)
library(lubridate)
library(janitor)
library(dplyr)
library(openxlsx)
library(xlsx)
#Puedo crear un proyecto en la carpeta MEGA para incluir todo lo de adentro


# Archivos mensuales ------------------------------------------------------
#___________________Cada mes debo cambiar las variables #fecha_mes y archivo
#___________________Estas variables sirven a todas las BBDD del Script


meses <- c("01")
# meses <- c("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11","12")
for (i in meses) {
  fecha_mes <- paste0("2022-",i,"-01")
  ruta_base <- "C:/Users/control.gestion3/OneDrive/"
  archivo <- paste0(ruta_base,"BBDD Produccion/REM/Serie A/2022/2022-",i," REM serie A.xlsx")
  


# EAR ---------------------------------------------------------------------

#A.4_1.3
  A413N <- 
    read_excel(archivo, sheet = "A07", range = "B71", col_names = FALSE) +
    read_excel(archivo, sheet = "A30", range = "J17", col_names = FALSE) +
    read_excel(archivo, sheet = "A30", range = "K17", col_names = FALSE) +
    read_excel(archivo, sheet = "A30", range = "L17", col_names = FALSE) +
    read_excel(archivo, sheet = "A30", range = "M17", col_names = FALSE) +
    read_excel(archivo, sheet = "A30", range = "Q17", col_names = FALSE) +
    read_excel(archivo, sheet = "A30", range = "R17", col_names = FALSE) +
    read_excel(archivo, sheet = "A30", range = "S17", col_names = FALSE) +
    read_excel(archivo, sheet = "A30", range = "T17", col_names = FALSE) +
    read_excel(archivo, sheet = "A32", range = "B45", col_names = FALSE)
  
  colnames(A413N)[1] <- fecha_mes
  if (i == "01"){
    A413N2 <- A413N
  }
  else{
    A413N2 <- cbind(A413N2,A413N)}
  
  #B.3.1.5
  
  B315N <- read_excel(archivo, sheet = "A21", range = "K13", col_names = FALSE)
  
  colnames(B315N)[1] <- fecha_mes
  if (i == "01"){
    B315N2 <- B315N
  }
  else{
    B315N2 <- cbind(B315N2,B315N)}
  
  B315D <- read_excel(archivo, sheet = "A21", range = "E13", col_names = FALSE)
  
  colnames(B315D)[1] <- fecha_mes
  if (i == "01"){
    B315D2 <- B315D
  }
  else{
    B315D2 <- cbind(B315D2,B315D)}
  
  #B.4.1.3
  B413N <- read_excel(archivo, sheet = "A08", range = "AS12", col_names = FALSE) - 
    read_excel(archivo, sheet = "A08", range = "B12", col_names = FALSE)
  
  colnames(B413N)[1] <- fecha_mes
  if (i == "01"){
    B413N2 <- B413N
  }
  else{
    B413N2 <- cbind(B413N2,B413N)}
  
  B413D <- read_excel(archivo, sheet = "A08", range = "AS12", col_names = FALSE)
  
  colnames(B413D)[1] <- fecha_mes
  if (i == "01"){
    B413D2 <- B413D
  }
  else{
    B413D2 <- cbind(B413D2,B413D)}
  
  #B.4.1.4
  B414N <- read_excel(archivo, sheet = "A21", range = "H88", col_names = FALSE) + 
    read_excel(archivo, sheet = "A21", range = "I88", col_names = FALSE)
  
  colnames(B414N)[1] <- fecha_mes
  if (i == "01"){
    B414N2 <- B414N
  }
  else{
    B414N2 <- cbind(B414N2,B414N)}
  
  B414D <- read_excel(archivo, sheet = "A21", range = "F88", col_names = FALSE) + 
    read_excel(archivo, sheet = "A21", range = "G88", col_names = FALSE)
  
  colnames(B414D)[1] <- fecha_mes
  if (i == "01"){
    B414D2 <- B414D
  }
  else{
    B414D2 <- cbind(B414D2,B414D)}
  
  #B.4.1.5
  
  B415N <- read_excel(archivo, sheet = "A08", range = "C92", col_names = FALSE)
  
  colnames(B415N)[1] <- fecha_mes
  if (i == "01"){
    B415N2 <- B415N
  }
  else{
    B415N2 <- cbind(B415N2,B415N)}
  
  B415D <- 
    read_excel(archivo, sheet = "A08", range = "C92", col_names = FALSE) + 
    read_excel(archivo, sheet = "A08", range = "C93", col_names = FALSE) +
    read_excel(archivo, sheet = "A08", range = "C94", col_names = FALSE) +
    read_excel(archivo, sheet = "A08", range = "C97", col_names = FALSE)
  
  colnames(B415D)[1] <- fecha_mes
  if (i == "01"){
    B415D2 <- B415D
  }
  else{
    B415D2 <- cbind(B415D2,B415D)}
  
  #C.4.3.1
  
  C431N <- 
    read_excel(archivo, sheet = "A07", range = "W71", col_names = FALSE) + 
    read_excel(archivo, sheet = "A07", range = "AA71", col_names = FALSE) +
    read_excel(archivo, sheet = "A32", range = "Z45", col_names = FALSE) +
    read_excel(archivo, sheet = "A32", range = "AE45", col_names = FALSE)
  
  
  colnames(C431N)[1] <- fecha_mes
  if (i == "01"){
    C431N2 <- C431N
  }
  else{
    C431N2 <- cbind(C431N2,C431N)}
  
  C431D <- 
    read_excel(archivo, sheet = "A07", range = "B71", col_names = FALSE) + 
    read_excel(archivo, sheet = "A32", range = "B45", col_names = FALSE) 
  
  colnames(C431D)[1] <- fecha_mes
  if (i == "01"){
    C431D2 <- C431D
  }
  else{
    C431D2 <- cbind(C431D2,C431D)}
  
  #C.4.3.4
  
  C434N <-
    read_excel(archivo, sheet = "A07", range = "AL71", col_names = FALSE) + 
    read_excel(archivo, sheet = "A07", range = "AM71", col_names = FALSE) +
    read_excel(archivo, sheet = "A32", range = "X45", col_names = FALSE) +
    read_excel(archivo, sheet = "A32", range = "Y45", col_names = FALSE)
  
  colnames(C434N)[1] <- fecha_mes
  if (i == "01"){
    C434N2 <- C434N
  }
  else{
    C434N2 <- cbind(C434N2,C434N)}
  
# D412 ------------------------------------------------------
  A04N <- read_excel(archivo, sheet = "A04", range = "E106", col_names = FALSE)+ 
    read_excel(archivo, sheet = "A04", range = "J106", col_names = FALSE)

  
  A04D <- read_excel(archivo, sheet = "A04", range = "B106", col_names = FALSE) +
    read_excel(archivo, sheet = "A04", range = "C106", col_names = FALSE)

  colnames(A04N)[1] <- fecha_mes
  if (i == "01"){
    Numerador <- A04N
  }
  else{
    Numerador <- cbind(Numerador,A04N)}
  
  colnames(A04D)[1] <- fecha_mes
  if (i == "01"){
    Denominador <- A04D
  }
  else{
    Denominador <- cbind(Denominador,A04D)}
  

# D417 ------------------------------------------------------
  D417N <- 
    read_excel(archivo, sheet = "A19b", range = "E11", col_names = FALSE) + 
    read_excel(archivo, sheet = "A19b", range = "F11", col_names = FALSE) 
  
  
  colnames(D417N)[1] <- fecha_mes
  if (i == "01"){
    D417N2 <- D417N
  }
  else{
    D417N2 <- cbind(D417N2,D417N)}
  
  D417D <- 
    read_excel(archivo, sheet = "A19b", range = "E11", col_names = FALSE) + 
    read_excel(archivo, sheet = "A19b", range = "F11", col_names = FALSE) +
    read_excel(archivo, sheet = "A19b", range = "G11", col_names = FALSE)
  
  colnames(D417D)[1] <- fecha_mes
  if (i == "01"){
    D417D2 <- D417D
  }
  else{
    D417D2 <- cbind(D417D2,D417D)}
  
  # D435 ------------------------------------------------------
  C435N <- 
    read_excel(archivo, sheet = "A07", range = "AL71", col_names = FALSE) + 
    read_excel(archivo, sheet = "A07", range = "AM71", col_names = FALSE) +
    read_excel(archivo, sheet = "A32", range = "X45", col_names = FALSE) +
    read_excel(archivo, sheet = "A32", range = "Y45", col_names = FALSE)
  
  
  colnames(C435N)[1] <- fecha_mes
  if (i == "01"){
    C435N2 <- C435N
  }
  else{
    C435N2 <- cbind(C435N2,C435N)}
  
  C435D <- 
    read_excel(archivo, sheet = "A07", range = "B71", col_names = FALSE) + 
    read_excel(archivo, sheet = "A32", range = "B45", col_names = FALSE) +
    read_excel(archivo, sheet = "A30", range = "J17", col_names = FALSE) +
    read_excel(archivo, sheet = "A30", range = "K17", col_names = FALSE) +
    read_excel(archivo, sheet = "A30", range = "L17", col_names = FALSE) +
    read_excel(archivo, sheet = "A30", range = "M17", col_names = FALSE) 
  
  colnames(C435D)[1] <- fecha_mes
  if (i == "01"){
    C435D2 <- C435D
  }
  else{
    C435D2 <- cbind(C435D2,C435D)}

}

A413 <- A413N2
A413$indicador <- "A.4.1.3"
A413$nombre <- "Porcentaje de cumplimiento de la programación anual de consultas médicas realizadas por especialista presencial, remota y por telemedicina"
A413$variable <- c("P1")


B315 <- rbind(B315N2, B315D2)
B315$indicador <- "B.3.1.5"
B315$nombre <- "Porcentaje de Horas Ocupadas de Quirófanos Habilitados"
B315$variable <- c("P1", "P2")

B413 <- rbind(B413N2, B413D2)
B413$indicador <- "B.4.1.3"
B413$nombre <- "Porcentaje de Abandono de Pacientes del Proceso de Atención Médica en Unidades de Emergencia Hospitalaria"
B413$variable <- c("P1", "P2")


B414 <- rbind(B414N2, B414D2)
B414$indicador <- "B.4.1.4"
B414$nombre <- "Porcentaje de Intervenciones Quirúrgicas Suspendidas"
B414$variable <- c("P1", "P2")


B415 <- rbind(B415N2, B415D2)
B415$indicador <- "B.4.1.5"
B415$nombre <- "Porcentaje de pacientes con indicación de hospitalización desde UEH, que acceden a cama de dotación en menos de 12 horas"
B415$variable <- c("P1", "P2")

C431 <- rbind(C431N2, C431D2)
C431$indicador <- "C.4.3.1"
C431$nombre <- "Porcentaje de consultas nuevas de especialidad médica en atención secundaria"
C431$variable <- c("P1", "P2")

C434 <- C434N2
C434$indicador <- "C.4.3.4"
C434$nombre <- "Porcentaje de Cumplimiento del envío de Contrarreferencia al alta de especialidad médica"
C434$variable <- c("P1")


D412 <- rbind(Numerador, Denominador)
D412$indicador <- "D.4.1.2"
D412$nombre <- "Porcentaje de Despacho de Receta Total y Oportuno"
D412$variable <- c("P1", "P2")

D417 <- rbind(D417N2, D417D2)
D417$indicador <- "D.4.1.7"
D417$nombre <- "Oportunidad, Calidad y Gestión de la Respuesta a los Reclamos"
D417$variable <- c("P1", "P2")

C435 <- rbind(C435N2, C435D2)
C435$indicador <- "C.4.3.5"
C435$nombre <- "Porcentaje de altas médicas de consulta de especialidad en atención secundaria"
C435$variable <- c("P1", "P2")


rutas_rem_ear <- rbind(A413, B315, B413, B414, B415, C431, C434, C435, D412, D417)
rutas_rem_ear2 <- rutas_rem_ear %>% select(indicador, nombre, variable)
rutas_rem_ear <-  round(rutas_rem_ear %>% select(-indicador, -nombre, -variable))
rutas_rem_ear <- cbind(rutas_rem_ear2, rutas_rem_ear)

openxlsx::write.xlsx(rutas_rem_ear, "C:/Users/control.gestion3/OneDrive/BBDD Produccion/Indicadores/Rutas REM/rutas_rem_ear_2022.xlsx", 
                     colNames = TRUE, sheetName = "rutas", overwrite = TRUE)

rm(A04D, A04N, Denominador, Numerador, rutas_rem_ear, rutas_rem_ear2,
   i, archivo, fecha_mes, meses, ruta_base, D412,
   A411, A411N, A411N2, 
   A412, A412N, A412N2,
   C434, C434N, C434N2,
   B315, B315D, B315D2, B315N, B315N2,  
   B413, B413D, B413D2, B413N, B413N2,
   B414, B414D, B414D2, B414N, B414N2,
   B415, B415D, B415D2, B415N, B415N2,
   C431, C431D, C431D2, C431N, C431N2,
   C435, C435D, C435D2, C435N, C435N2,
   D417, D417D, D417D2, D417N, D417N2)
