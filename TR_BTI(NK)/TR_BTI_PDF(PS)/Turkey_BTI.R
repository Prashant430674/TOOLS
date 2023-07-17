library(openxlsx)
library(magick)
library(imager)

dir <- dirname(rstudioapi::getSourceEditorContext()$path)

path_file <- paste(dir,"/IN_FILES/TR_BTI.xlsx",sep = "")
path_file1 <- paste(dir,"/IN_FILES/req_file.xlsx",sep = "")

TR_BTI <- read.xlsx(path_file,sep.names = " ")
header <- read.xlsx(path_file1,colNames = FALSE)

#(nrow(TR_BTI)-1)
#######################################################################
for (var_pdf in 1:nrow(TR_BTI)) {
  
  df <- data.frame("col1"="","col2"="",stringsAsFactors = FALSE)[-1,]
  
  df[1,1] <- header[1,1]
  df[1,2] <- ""
  
  df[2,2] <- "BTB Numarasi:"
  df[2,3] <- TR_BTI[var_pdf,1]
  
  df[3,2] <- "Gtip No:"
  df[3,3] <- TR_BTI[var_pdf,2]
  
  df[4,2] <- "Ge?erlilik Bsg.Tarihi:"
  df[4,3] <- TR_BTI[var_pdf,3]
  
  df[5,2] <- "Siniflandirmanin Gerek?esi:"
  df[6,2] <- TR_BTI[var_pdf,4]
  
  df[7,2] <- "Esyanin Tanimi:"
  df[8,2] <- TR_BTI[var_pdf,5]
  

  
  excel_name <- paste(dir,"/excel_files/",TR_BTI[var_pdf,1],".xlsx",sep = "")

  df[is.na(df)] <- " "
  
  library(xlsx)
  
  write.xlsx(df,excel_name,row.names = FALSE,col.names = FALSE)
  
  
  



############################################################################
################################################################
####### excel formatting #######################################  

wb <- loadWorkbook(excel_name)  # load workbook

sheets <- getSheets(wb)               # get all sheets
sheet <- sheets[[1]] 



#### TO SET COLUMN WIDTH FOR COLUMN "c"##########

setColumnWidth(sheet,colIndex =1, colWidth = 38)
setColumnWidth(sheet,colIndex =2, colWidth = 40)
setColumnWidth(sheet,colIndex =3, colWidth = 20)
###### TO SET ROW HEIGHT FOR HEADERS ############

row_height <- getRows(sheet,rowIndex = 1)
setRowHeight(row_height,50)

sheets <- getSheets(wb)               # get all sheets
sheet <- sheets[[1]] 



############# Adding  alignmeent ############## 
rows_border <- getRows(sheet, rowIndex= 1:9 )   # get rows
cells_border <- getCells(rows_border, colIndex = 1:2 )



#border1 <-  Border( position=c("TOP", "BOTTOM","LEFT","RIGHT"))  
###### Allbordders #########


F2 <- CellStyle(wb,alignment = Alignment(horizontal = "ALIGN_LEFT",vertical = "VERTICAL_TOP"))  

for (var1 in 3:10) {
  
  setCellStyle(cells_border[[var1]], F2)
  
}

rows <- getRows(sheet, rowIndex= 1:8 )   # get rows
cells_wrap <- getCells(rows, colIndex = 2 )



###### wrap text ###########
F1 <- CellStyle(wb,alignment = Alignment(wrapText = TRUE,horizontal = "ALIGN_LEFT",vertical = "VERTICAL_TOP") )

setCellStyle(cells_wrap[[6]], F1)
setCellStyle(cells_wrap[[7]], F1)
setCellStyle(cells_wrap[[8]], F1)

#################################################

rows <- getRows(sheet, rowIndex= 2:8 )   # get rows
cells_wrap <- getCells(rows, colIndex = 2:3 )



###### headers bold###########
F4 <- CellStyle(wb,alignment = Alignment(wrapText = TRUE,horizontal = "ALIGN_LEFT",vertical = "VERTICAL_TOP"),font = Font(wb,heightInPoints=12,color = "#16365C",isBold=TRUE) )
F5 <- CellStyle(wb,alignment = Alignment(wrapText = TRUE,horizontal = "ALIGN_LEFT",vertical = "VERTICAL_TOP"),font = Font(wb,heightInPoints=12,color = "#16365C") )


setCellStyle(cells_wrap[[1]], F4)
setCellStyle(cells_wrap[[3]], F4)
setCellStyle(cells_wrap[[5]], F4)
setCellStyle(cells_wrap[[7]], F4)
setCellStyle(cells_wrap[[11]], F4)

setCellStyle(cells_wrap[[2]], F5)
setCellStyle(cells_wrap[[4]], F5)
setCellStyle(cells_wrap[[6]], F5)
setCellStyle(cells_wrap[[9]], F5)
setCellStyle(cells_wrap[[13]], F5)
###### TO merge cells ############# 
addMergedRegion(sheet, 1, 1, 1, 3)

F3 <- CellStyle(wb,alignment = Alignment(h="ALIGN_CENTER",vertical = "VERTICAL_TOP"),font = Font(wb,name ="Cambria",heightInPoints=16, isBold=TRUE,color = "red"))

rows <- getRows(sheet, rowIndex= 1 )   # get rows
cells_font <- getCells(rows, colIndex = 1 )

setCellStyle(cells_font[[1]], F3)


saveWorkbook(wb, excel_name)



############ ADDING IMAGE #####################################  

  #IMG <- image_read(TR_BTI[var_pdf,6])

img_name <- paste(dir,"/images/",TR_BTI[var_pdf,1],".png",sep = "")

download.file(TR_BTI[var_pdf,6],img_name, mode = 'wb')


  
  #image_write(IMG, path = img_name, format = "png")
  
 
  img <- load.image(img_name)

  ##### Dimention reduction with respect to height ##########  
 
    
  img <- resize(img,200,200)

  save.image(img,img_name)

  
  detach("package:xlsx", unload = TRUE)
  
  library(openxlsx)
  
  wb <- openxlsx::loadWorkbook(excel_name)
  wb %>% 
    insertImage("Sheet1", img_name, width = 7, height = 6
                , startRow = 2, startCol = 1, units = "cm", dpi = 300) 
  wb %>% 
    saveWorkbook(excel_name, overwrite = TRUE)
  
  

  Sys.sleep(1)
  
  
}   


sessionInfo()
