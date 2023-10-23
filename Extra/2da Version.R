#Librerias
library(readxl)
library(openxlsx)
library(dplyr)
library(writexl)



clientes_cuit <- read_excel("C:/Users/Marcos/Desktop/CUITS.xlsx")
file_paths <- list.files(path = "C:/Users/Marcos/Desktop/Meses",pattern = ".xlsx")


for(c in 1:nrow(clientes_cuit)){
  all_data_emitidos <- data.frame()
  all_data_recibidos <- data.frame()
  
  for(file_path in file_paths){
    if(grepl(clientes_cuit[c,2], file_path)){
      if(grepl('Emitidos', file_path)){
        data_emitidos <- read_excel(paste0("C:/Users/Marcos/Desktop/Meses/",file_path), skip = 1)
        all_data_emitidos <- rbind(all_data_emitidos,(arrange(data_emitidos, desc(row_number()))))
      }
      else{
        data_recibidos <- read_excel(paste0("C:/Users/Marcos/Desktop/Meses/",file_path), skip = 1)
        all_data_recibidos <- rbind(all_data_recibidos,(arrange(data_recibidos, desc(row_number()))))
      }
    }
  }
  if(length(all_data_emitidos) > 0){
    data_final <- arrange(subset(all_data_emitidos, select = c('Fecha', 'Tipo', 'Número Desde', 'Denominación Receptor', 'Imp. Total')), desc(row_number()))
    colnames(data_final) <- c("Fecha", "Tipo", "Nro Factura", "Denom. Receptor", "Imp. Total")
    data_final$Fecha <- as.Date(data_final$Fecha, format = "%d/%m/%Y")
    data_final <- data_final[order(data_final$Fecha), ]
    data_final$Fecha <- format(data_final$Fecha, format = "%d/%m/%Y")
    ``
    file_path_a <- paste0("C:/Users/Marcos/Desktop/Oficina/Monotributo/",clientes_cuit[c,1],".xlsx")
    wb <- loadWorkbook(file_path_a)
    excel_cliente <- read_xlsx(file_path_a, sheet = "VENTAS NUEVO", range = cell_cols(1:5))
    size <- nrow(excel_cliente)
    excel_cliente$Fecha <- format(excel_cliente$Fecha, format = "%d/%m/%Y")
    excel_cliente <- rbind(excel_cliente,data_final)
    writeData(wb, sheet = "VENTAS NUEVO", x = excel_cliente, startRow = 6, startCol = 1)
    
    for(i in 1:nrow(data_final)){
      b=size+6+i
      addStyle(wb, sheet = "VENTAS NUEVO", rows = b, cols = 1:5, style = createStyle(fontName = "Calibri", fontSize = 11))
    }
    addStyle(wb, sheet = "VENTAS NUEVO", rows = 7:(nrow(excel_cliente)+7), cols = 3, style =  createStyle(halign = "center", fontName = "Calibri", fontSize = 11))
    saveWorkbook(wb, file_path_a, overwrite = TRUE)
  }
  
  if(length(all_data_recibidos) > 0){
    data_final <- arrange(subset(all_data_recibidos, select = c('Fecha', 'Tipo', 'Número Desde', 'Denominación Emisor', 'Imp. Total')), desc(row_number()))
    colnames(data_final) <- c("Fecha", "Tipo", "Nro Desde", "Denom. Emisor", "Imp. Total")
    data_final$Fecha <- as.Date(data_final$Fecha, format = "%d/%m/%Y")
    data_final <- data_final[order(data_final$Fecha), ]
    data_final$Fecha <- format(data_final$Fecha, format = "%d/%m/%Y")
    
    file_path_a <- paste0("C:/Users/Marcos/Desktop/Oficina/Monotributo/",clientes_cuit[c,1],".xlsx")
    wb <- loadWorkbook(file_path_a)
    excel_cliente <- read_xlsx(file_path_a, sheet = "COMPRAS NUEVO", range = cell_cols(1:5))
    size <- nrow(excel_cliente)
    excel_cliente$Fecha <- format(excel_cliente$Fecha, format = "%d/%m/%Y")
    excel_cliente <- rbind(excel_cliente,data_final)
    writeData(wb, sheet = "COMPRAS NUEVO", x = excel_cliente, startRow = 6, startCol = 1)
    for(i in 1:nrow(data_final)){
      b=size+6+i
      addStyle(wb, sheet = "COMPRAS NUEVO", rows = b, cols = 1:5, style = createStyle(fontName = "Calibri", fontSize = 11))
    }
    addStyle(wb, sheet = "COMPRAS NUEVO", rows = 7:(nrow(excel_cliente)+7), cols = 3, style =  createStyle(halign = "center", fontName = "Calibri", fontSize = 11))
    saveWorkbook(wb, file_path_a, overwrite = TRUE)
  }
}
