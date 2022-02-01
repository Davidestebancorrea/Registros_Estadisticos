
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


meses <- c("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11","12")
for (i in meses) {
  fecha_mes <- paste0("2021-",i,"-01")
  ruta_base <- "C:/Users/control.gestion3/OneDrive/"
  archivo <- paste0(ruta_base,"BBDD Produccion/REM/Serie A/2021/2021-",i," REM serie A.xlsx")
  

# Indicador Farmacia ------------------------------------------------------
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
  

# Indicador Urgencia <12H -------------------------------------------------
  A08N <- read_excel(archivo, sheet = "A08", range = "C92", col_names = FALSE)
  
  A08D <- read_excel(archivo, sheet = "A08", range = "C92", col_names = FALSE) +
    read_excel(archivo, sheet = "A08", range = "C93", col_names = FALSE)+
    read_excel(archivo, sheet = "A08", range = "C94", col_names = FALSE)+
    read_excel(archivo, sheet = "A08", range = "C97", col_names = FALSE)
  
  colnames(A08N)[1] <- fecha_mes
  if (i == "01"){
    Numerador2 <- A08N
  }
  else{
    Numerador2 <- cbind(Numerador2,A08N)}
  
  colnames(A08D)[1] <- fecha_mes
  if (i == "01"){
    Denominador2 <- A08D
  }
  else{
    Denominador2 <- cbind(Denominador2,A08D)}
  
  
}

A04 <- rbind(Numerador, Denominador)
A08 <- rbind(Numerador2, Denominador2)

openxlsx::write.xlsx(A04, "C:/Users/control.gestion3/OneDrive/BBDD Produccion/Indicadores/Rutas REM/A04.xlsx", colNames = TRUE, sheetName = "A04", overwrite = TRUE)

openxlsx::write.xlsx(A08, "C:/Users/control.gestion3/OneDrive/BBDD Produccion/Indicadores/Rutas REM/A08.xlsx", colNames = TRUE, sheetName = "A08", overwrite = TRUE)

rm(A04, A04D, A04N, Numerador, Numerador2, Denominador, Denominador2, A08, A08D, A08N)
