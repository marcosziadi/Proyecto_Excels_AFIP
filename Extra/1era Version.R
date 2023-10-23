#Librerias
library(readxl)
library(openxlsx)
library(dplyr)
library(writexl)



#Carga de archivos
file_paths <- list.files(path = "C:/Users/Marcos/Desktop/Meses",pattern = ".xlsx")
all_data_emitidos <- list()
all_data_recibidos <- list()
for (file_path in file_paths){
  if(grepl('Emitidos', file_path)){
    data_emitidos <- read_excel(paste0("C:/Users/Marcos/Desktop/Meses/",file_path), skip = 1)  # Skip the first row
    all_data_emitidos[[file_path]] <- arrange(data_emitidos, desc(row_number()))
  }
  else{
    data_recibidos <- read_excel(paste0("C:/Users/Marcos/Desktop/Meses/",file_path), skip = 1)  # Skip the first row
    all_data_recibidos[[file_path]] <- arrange(data_recibidos, desc(row_number()))
  }
}


  
#Filtración de columnas y cambio de nombre
data_final <- arrange(subset(do.call(rbind, all_data_emitidos), select = c('Fecha', 'Tipo', 'Número Desde', 'Nro. Doc. Receptor', 'Imp. Total')), desc(row_number()))
colnames(data_final) <- c("fecha", "tipo", "nro_desde", "nro_doc", "imp_total")



#Formateo de Fecha y ordenamiento del dataset
data_final$fecha <- as.Date(data_final$fecha, format = "%d/%m/%Y")
#data_final <- data_final[order(data_final$fecha), ]
data_final$fecha <- format(data_final$fecha, format = "%d/%m/%Y")



file_path <- "C:/Users/Marcos/Desktop/Oficina/FINAL-VENTAS.xlsx"
write.xlsx(data_final,file_path)
wb <- loadWorkbook(file_path)
addWorksheet(wb, "Ventas")



#Cantidad de Meses y Mes inicial
meses_totales<-length(all_data_emitidos)
mes_inicial <- as.integer(substring(data_final[1,1],4,5))



# Notas de Credito - > Importe a negativo
for (i in 1:nrow(data_final)){
  header_style <- createStyle(
    borderStyle = "thin",
    border ="Right"
  )
  addStyle(wb, sheet = "Ventas", rows = (i+1), cols = 1:ncol(data_final), style = header_style)
  if(grepl('Nota de Crédito', data_final[i,2])){
    data_final[i,5] = data_final[i,5]*-1
    header_style <- createStyle(
      fgFill = "#FABBB8",
      borderStyle = "thin",
      border ="Right"
    )
    addStyle(wb, sheet = "Ventas", rows = (i+1), cols = 1:ncol(data_final), style = header_style)
  }
}



# Listas de todas las Facturas y de todas las Notas de Crédito
list_nota <- list()
list_factura <- list()
nota <- 1
factura <- 1
for (i in 1:nrow(data_final)){
  if(grepl('Nota de Crédito', data_final[i,2])){
    list_nota[nota] = data_final[i,3]
    nota = nota+1
  }
  else if(grepl('Factura', data_final[i,2])){
    list_factura[factura] = data_final[i,3]
    factura = factura+1
  }
}



# Corroboración de si falta alguna nota de crédito o Factura
if(length(list_nota)>1){
  for(i in 1:(length(list_nota)-1)){
    if(list_nota[i] != as.integer(list_nota[i+1])-1){
      print(paste("Falta la Nota de Crédito numero:", (list_nota[i]+1)))
      # break
    }
  }
}
if(length(list_factura)>1){
  for (i in 1:(length(list_factura) - 1)){
    if (as.integer(list_factura[i]) != (as.integer(list_factura[i+1])-1)){
      print(paste("Falta la Factura número:", (as.integer(list_factura[i]) + 1)))
      # break
    }
  }
}



#Creacion de un excel y una nueva Hoja
colnames(data_final) <- c('Dia', 'Comprobante', 'Nro Factura', 'CUIT', 'Importe')
writeData(wb, sheet = "Ventas", x = data_final)
colnames(data_final) <- c("fecha", "tipo", "nro_desde", "nro_doc", "imp_total")
removeWorksheet(wb, sheet = "Sheet 1")

header_style <- createStyle(
  fgFill = "#EEFAFF",
  halign = "center",
  textDecoration = "bold",
  fontSize = 14,
  borderStyle = "thin",  # Set the border thickness to "thin"
  border ="TopBottomLeftRight"
)
addStyle(wb, sheet = "Ventas", rows = 1, cols = 1:ncol(data_final), style = header_style)

setColWidths(wb, sheet = "Ventas", cols = 1:ncol(data_final), widths = 15)
setColWidths(wb, sheet = "Ventas", cols = 2, widths = 20)


addWorksheet(wb, sheetName = "Ventas Escalera")



#Creacion tabla Ventas Escalera
acomo_data_final <- matrix(NA, nrow=nrow(data_final), ncol= meses_totales)
acomo_data_final <- as.data.frame(acomo_data_final)
a <- 1
for(i in 1:(nrow(data_final))){
  if((substring(data_final[i,1],4,5) == substring(data_final[i+1,1],4,5)) || (nrow(data_final)==i)){
    acomo_data_final[i,a]=data_final[i,5]
    for(c in 1:meses_totales){
      # if(acomo_data_final[i,c]!=data_final[i,5]){
      #   acomo_data_final[i,c] = NA
      # }
      if(c!=a){
        acomo_data_final[i,c] = NA
      }
    }
  }
  else{
    acomo_data_final[i,a]=data_final[i,5]
    a = a+1
  }
}
acomo_data_final <- as.data.frame(sapply(acomo_data_final, as.integer))



#Creación Lista nombre de meses
c<-1
a<-mes_inicial
mes_nombre<-list()
for (i in mes_inicial:(mes_inicial+meses_totales)){
  if(a==13){
      a=1
  }
  mes_nombre[c] <- month.name[a]
  c=c+1
  a=a+1
}



#Asignacion nombre meses por columna
a<-mes_inicial
año<-as.integer(substring(data_final[1,1],7,10))
for(i in 1:length(mes_nombre)){
  if(a==13){
    año=año+1
    a=1
  }
  mes_nombre[i]<-paste(mes_nombre[i],año)
  a=a+1
}
for (i in 1:length(acomo_data_final)) {
  colnames(acomo_data_final)[i] <- mes_nombre[i]
}



#Suma total ultimos 12 meses
suma_mes <- colSums(acomo_data_final, na.rm = TRUE)
sumas_12_meses <- list()
sumas_12_meses[1] <- NA
suma_p <- suma_mes[1]
c<-1
for(i in 2:length(acomo_data_final)){
  suma_p = suma_p + suma_mes[i]
  if(i>12){
    # suma_p = suma_p + suma_mes[i]
    # sumas_12_meses[i] = suma_p
    suma_p = suma_p - suma_mes[i-12]
    sumas_12_meses[i] = suma_p
  }else{
    sumas_12_meses[i] = NA
  }
}
sumas_12_meses <- as.data.frame(sapply(sumas_12_meses, as.integer))



#Asignacion Suma total 12 meses a la hoja Ventas Escalera
num_columns <- length(sumas_12_meses)
sum_row <- data.frame(t(sumas_12_meses))
suma_mes <- t(as.data.frame(as.vector(as.matrix(suma_mes))))
colnames(suma_mes) <- colnames(acomo_data_final)
acomo_data_final <- rbind(suma_mes, acomo_data_final)
writeData(wb, sheet = "Ventas Escalera", x = acomo_data_final, startRow = 2, startCol = 1, rowNames = FALSE)
writeData(wb, sheet = "Ventas Escalera", x = sum_row, startRow = 1, startCol = 1, colNames = FALSE)
setColWidths(wb, sheet = "Ventas Escalera", cols = 1:ncol(acomo_data_final), widths = 17)



addStyle(wb, sheet = "Ventas Escalera", rows = 3, cols = 1:ncol(acomo_data_final), style = createStyle(
  fgFill = "#EEFAFF",
  halign = "center",
  textDecoration = "bold",
  borderStyle = "thin",
  border ="BottomLeftRight"
))
header_style <- createStyle(
  fgFill = "#CCFFCC",
  halign = "center",
  textDecoration = "bold",
  fontSize = 14,
  borderStyle = "thin",
  border ="TopLeftRight"
)
addStyle(wb, sheet = "Ventas Escalera", rows = 2, cols = 1:ncol(acomo_data_final), style = header_style)
header_style <- createStyle(
  fgFill = "#CCFFCC",
)
for(i in 4:(nrow(acomo_data_final)+2)){
  addStyle(wb, sheet = "Ventas Escalera", rows = i, cols = 1:ncol(acomo_data_final), style = header_style)
}

saveWorkbook(wb, file_path, overwrite = TRUE)






addWorksheet(wb, sheetName = "Compras")

data_final_compras <- arrange(subset(do.call(rbind, all_data_recibidos), select = c('Fecha', 'Tipo', 'Número Desde', 'Denominación Emisor', 'Imp. Total')), desc(row_number()))
colnames(data_final_compras) <- c("fecha", "tipo", "nro_desde", "den_emis", "imp_total")



#Formateo de Fecha y ordenamiento del dataset
data_final_compras$fecha <- as.Date(data_final_compras$fecha, format = "%d/%m/%Y")
data_final_compras <- data_final_compras[order(data_final_compras$fecha), ]
data_final_compras$fecha <- format(data_final_compras$fecha, format = "%d/%m/%Y")



#Cantidad de Meses y Mes inicial
meses_totales_compras<-length(all_data_recibidos)
mes_inicial_compras <- as.integer(substring(data_final_compras[1,1],4,5))



# Notas de Credito - > Importe a negativo
for (i in 1:nrow(data_final_compras)){
  header_style <- createStyle(
    borderStyle = "thin",
    border ="Right"
  )
  addStyle(wb, sheet = "Compras", rows = (i+1), cols = 1:ncol(data_final_compras), style = header_style)
    if(grepl('Nota de Crédito', data_final_compras[i,2])){
    data_final_compras[i,5] = data_final_compras[i,5]*-1
    header_style <- createStyle(
      fgFill = "#FABBB8",
      borderStyle = "thin",
      border ="Right"
    )
    addStyle(wb, sheet = "Compras", rows = (i+1), cols = 1:ncol(data_final_compras), style = header_style)
  }
}



# Listas de todas las Facturas y de todas las Notas de Crédito
# list_nota <- list()
# list_factura <- list()
# nota <- 1
# factura <- 1
# for (i in 1:nrow(data_final_compras)){
#   if(grepl('Nota de Crédito', data_final_compras[i,2])){
#     list_nota[nota] = data_final_compras[i,3]
#     nota = nota+1
#   }
#   else if(grepl('Factura', data_final_compras[i,2])){
#     list_factura[factura] = data_final_compras[i,3]
#     factura = factura+1
#   }
# }



# Corroboración de si falta alguna nota de crédito o Factura
# if(length(list_nota)>1){
#   for(i in 1:(length(list_nota)-1)){
#     if(list_nota[i] != as.integer(list_nota[i+1])-1){
#       print(paste("Falta la Nota de Crédito numero:", (list_nota[i]+1)))
#       # break
#     }
#   }
# }
# if(length(list_factura)>1){
#   for (i in 1:(length(list_factura) - 1)){
#     if (as.integer(list_factura[i]) != (as.integer(list_factura[i+1])-1)){
#       print(paste("Falta la Factura número:", (as.integer(list_factura[i]) + 1)))
#       # break
#     }
#   }
# }



header_style <- createStyle(
  fgFill = "#EEFAFF",
  halign = "center",
  textDecoration = "bold",
  fontSize = 14,
  borderStyle = "thin",  # Set the border thickness to "thin"
  border ="TopBottomLeftRight"
)
addStyle(wb, sheet = "Compras", rows = 1, cols = 1:ncol(data_final_compras), style = header_style)

setColWidths(wb, sheet = "Compras", cols = 1:ncol(data_final), widths = 15)
setColWidths(wb, sheet = "Compras", cols = 2, widths = 20)
setColWidths(wb, sheet = "Compras", cols = 3, widths = 19)
setColWidths(wb, sheet = "Compras", cols = 4, widths = 30)
#setColWidths(wb, sheet = "Compras", cols = 2, widths = 19)
colnames(data_final_compras) <- c('Fecha', 'Tipo', 'Número Desde', 'Denom. Emisor', 'Imp. Total')
writeData(wb, sheet = "Compras", x = data_final_compras)
colnames(data_final_compras) <- c("fecha", "tipo", "nro_desde", "den_emis", "imp_total")
addWorksheet(wb, sheetName = "Compras Escalera")
saveWorkbook(wb, file_path, overwrite = TRUE)







#Creacion tabla Ventas Escalera
acomo_data_final_compras <- matrix(NA, nrow=nrow(data_final_compras), ncol= meses_totales_compras)
acomo_data_final_compras <- as.data.frame(acomo_data_final_compras)
a <- 1
for(i in 1:(nrow(data_final_compras))){
  if((substring(data_final_compras[i,1],4,5) == substring(data_final_compras[i+1,1],4,5)) || (nrow(data_final_compras)==i)){
    acomo_data_final_compras[i,a]=data_final_compras[i,5]
    for(c in 1:meses_totales_compras){
      # if(acomo_data_final[i,c]!=data_final[i,5]){
      #   acomo_data_final[i,c] = NA
      # }
      if(c!=a){
        acomo_data_final_compras[i,c] = NA
      }
    }
  }
  else{
    acomo_data_final_compras[i,a]=data_final_compras[i,5]
    a = a+1
  }
}
acomo_data_final_compras <- as.data.frame(sapply(acomo_data_final_compras, as.integer))



#Creación Lista nombre de meses
c<-1
a<-mes_inicial_compras
mes_nombre<-list()
for (i in mes_inicial:(mes_inicial_compras+meses_totales_compras)){
  if(a==13){
    a=1
  }
  mes_nombre[c] <- month.name[a]
  c=c+1
  a=a+1
}



#Asignacion nombre meses por columna
a<-mes_inicial_compras
año<-as.integer(substring(data_final_compras[1,1],7,10))
for(i in 1:length(mes_nombre)){
  if(a==13){
    año=año+1
    a=1
  }
  mes_nombre[i]<-paste(mes_nombre[i],año)
  a=a+1
}
for (i in 1:length(acomo_data_final_compras)) {
  colnames(acomo_data_final_compras)[i] <- mes_nombre[i]
}



#Suma total ultimos 12 meses
suma_mes <- colSums(acomo_data_final_compras, na.rm = TRUE)
sumas_12_meses <- list()
sumas_12_meses[1] <- NA
suma_p <- suma_mes[1]
c<-1
for(i in 2:length(acomo_data_final_compras)){
  suma_p = suma_p + suma_mes[i]
  if(i>12){
    # suma_p = suma_p + suma_mes[i]
    # sumas_12_meses[i] = suma_p
    suma_p = suma_p - suma_mes[i-12]
    sumas_12_meses[i] = suma_p
  }else{
    sumas_12_meses[i] = NA
  }
}
sumas_12_meses <- as.data.frame(sapply(sumas_12_meses, as.integer))



#Asignacion Suma total 12 meses a la hoja Ventas Escalera
num_columns <- length(sumas_12_meses)
sum_row <- data.frame(t(sumas_12_meses))
suma_mes <- t(as.data.frame(as.vector(as.matrix(suma_mes))))
colnames(suma_mes) <- colnames(acomo_data_final_compras)
acomo_data_final_compras <- rbind(suma_mes, acomo_data_final_compras)
writeData(wb, sheet = "Compras Escalera", x = acomo_data_final_compras, startRow = 2, startCol = 1, rowNames = FALSE)
writeData(wb, sheet = "Compras Escalera", x = sum_row, startRow = 1, startCol = 1, colNames = FALSE)
setColWidths(wb, sheet = "Compras Escalera", cols = 1:ncol(acomo_data_final_compras), widths = 17)


addStyle(wb, sheet = "Compras Escalera", rows = 3, cols = 1:ncol(acomo_data_final_compras), style = createStyle(
  fgFill = "#EEFAFF",
  halign = "center",
  textDecoration = "bold",
  borderStyle = "thin",
  border ="BottomLeftRight"
))
header_style <- createStyle(
  fgFill = "#CCFFCC",
  halign = "center",
  textDecoration = "bold",
  fontSize = 14,
  borderStyle = "thin",  # Set the border thickness to "thin"
  border ="TopLeftRight"
)
addStyle(wb, sheet = "Compras Escalera", rows = 2, cols = 1:ncol(acomo_data_final_compras), style = header_style)
header_style <- createStyle(
  fgFill = "#CCFFCC",
)
for(i in 4:(nrow(acomo_data_final_compras)+2)){
  addStyle(wb, sheet = "Compras Escalera", rows = i, cols = 1:ncol(acomo_data_final_compras), style = header_style)
}

saveWorkbook(wb, file_path, overwrite = TRUE)



