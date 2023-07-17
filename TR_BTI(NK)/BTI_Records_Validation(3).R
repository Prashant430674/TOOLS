# set working directory
setwd(dirname(rstudioapi::getSourceEditorContext()$path))

# load library
library(xlsx)
detach("package:xlsx", unload = TRUE)
library(openxlsx)
library(tidyverse)
library(stringr)
library(lookup)

# read the template
templateWb <- loadWorkbook('TR_BTI_Template.xlsx')

# read the DB records
DB_DF <- read.xlsx(list.files(getwd(), pattern = 'DB_Data'))

# take only first 35 columns
DB_DF <- DB_DF[, c(1:35)]

# add a column at the beginning as VALIDATION
DB_DF <- DB_DF %>%
  mutate(
    VALIDATION = '',
    .before = 1
  )


# read online records
Online_DF <- read.xlsx(list.files(getwd(), pattern = 'TR_BTI Records'))
Online_DF <- Online_DF %>%
  mutate(
    Validation = '',
    .before = 1
  )


# do lookup and update Tags in Online (either Addition or Present)
Online_DF$Validation <- lookup(Online_DF$DETAY, DB_DF$REF_NUMBER, DB_DF$REC_STATUS)

for(vi in 1:ncol(Online_DF)) {
  Online_DF[, vi] <- as.character(Online_DF[, vi])
}

Online_DF[is.na(Online_DF)] <- ""
Online_DF$Validation <- ifelse(Online_DF$Validation == 'A', paste0('Present'), paste0('Addition'))
# do lookup and update Tags in DB (either Deletion or Present)
DB_DF$VALIDATION <- lookup(DB_DF$REF_NUMBER,Online_DF$DETAY, Online_DF$DETAY)

for(vi in 1:ncol(DB_DF)) {
  DB_DF[, vi] <- as.character(DB_DF[, vi])
}

DB_DF[is.na(DB_DF)] <- ""
DB_DF$VALIDATION <- ifelse(DB_DF$VALIDATION == '', paste0('Deletion'), paste0('Present'))

# add the DB data in the respective sheets in template
writeData(templateWb, 4, DB_DF) # DB Data
writeData(templateWb, 3, Online_DF) # Online Data

# add the deletion records 
del_DF <- as.data.frame(DB_DF[which(DB_DF$VALIDATION == 'Deletion'), "REF_NUMBER"])
names(del_DF) <- "REF_NUMBER"
writeData(templateWb, 2, del_DF)

# save the addition BTI numbers from Online DF as a text file
add_DF <- Online_DF[which(Online_DF$Validation == 'Addition'), "DETAY"]
write_lines(add_DF, 'bti.txt')


# save the workbook
saveWorkbook(templateWb, paste(getwd(),"/TR_BTI_Update_", format(Sys.Date(),'%d%b%Y'), ".xlsx",sep = ''))

