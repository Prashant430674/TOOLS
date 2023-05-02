############################################################
##### TURKEY(TARIFF) #####
##### IRANNA BADANIKAYI
##### EXTRACTION TIME 30 hours
##### TOOL MIGHT STOP IF THE WEBSITE LOGIN SESSION GOT EXPIRED(LOGIN AGAIN AND RUN THE TOOL FROM WHERE IT GOT STOPPED)
############################################################

##### CUSTOM SCRIPT START

# This code is to get the current Chrome version in the system 
# and load it instead of using the old method which opens the 
# browser in another users login

# load library
library(chromote) # to get the chrome version
library(netstat) # to dynamically change ports
library(RSelenium) # to run Selenium

# get latest chrome browser version
b <- ChromoteSession$new()
browser_version <- b$Browser$getVersion()$product

# Remove all text before "/"
browser_version<-gsub(".*/","",browser_version)

# we need only the major release number
realVersion <- substr(browser_version,1,3)

# get the list of available chromedriver versions
chromeVersions <- binman::list_versions("chromedriver")
chromeVersions <- chromeVersions$win32

# get the matching chromedriver version as per current Chrome version
realChromeVer <- chromeVersions[which(substr(chromeVersions,1,3)==realVersion)]

# launch chrome
driver <- rsDriver(browser = 'chrome', port = free_port(), chromever = realChromeVer[1], verbose = F)
remdr <- driver[['client']]
remdr$maxWindowSize()

##### CUSTOM SCRIPT END

##The value 8000 is the number of megabytes of RAM to allow for the Java heap
options(java.parameters = "-Xmx8000m")

remdr$navigate(url ="http://www.tariff-tr.com/")

library(rvest)
library(stringi)
library(dplyr)
library(openxlsx)
library(lubridate)
library(tidyverse)
setwd(dirname(rstudioapi::getSourceEditorContext()$path))

Header_List <- read.table(paste(getwd(),"/headers.txt",sep = ''), header = TRUE, quote = "", colClasses = "character")
headers <- as.list(Header_List)
headers <- headers[[1]]

#login option button
webElem <- remdr$findElement(using = 'xpath',value = '//*[@id="navbar"]/ul/li[8]/a' )
webElem$clickElement()
Sys.sleep(4)
##username
webElem <- remdr$findElement(using = 'xpath',value = '//*[@id="txtUserName"]' )
webElem$sendKeysToElement(list('pragatishree'))
Sys.sleep(4)
##password
webElem <- remdr$findElement(using = 'xpath',value = '//*[@id="txtPassword"]' )
webElem$sendKeysToElement(list("1098"))
Sys.sleep(4)
#login enter button
webElem <- remdr$findElement(using = 'xpath',value = '//*[@id="btnLogin"]' )
webElem$clickElement()
Sys.sleep(4)
#Tariff Search button
#webElem <- remdr$findElement(using = 'xpath',value = '/html/body/div[3]/table/tbody/tr/td[2]/div/table/tbody/tr/td/table/tbody/tr/td[2]/div/a[1]' )
#webElem$clickElement()
#Sys.sleep(4)


Treeview<-data.frame()
final_tab<-data.frame()

#length(headers)
for(i in 1:length(headers)){
  
  var1 <- headers[i]
  
  #remdr$navigate(url = paste("https://www.tariff-tr.com/gtipgenel/gtipvergimevzuat.aspx?pos=",var1,"&ulke=TUM",sep = ""))
  remdr$navigate(url = paste("http://www.tariff-tr.com/Gtipgenel/gtipvergimevzuat.aspx?pos=",var1,"&ulke=TUM",sep = ""))
  
  Sys.sleep(3)
  
  m <- remdr$getPageSource()
  rawdata <- read_html(m[[1]])
  
  #Treeview extraction
  temp_treeview<- rawdata %>%
    html_nodes(xpath = '//*[@id="DivListeYeni"]/table') %>% 
    html_table(fill = TRUE,convert = FALSE)
  
  ##PAGE WAIT
  if(length(temp_treeview)<1){
    Sys.sleep(3)
    m <- remdr$getPageSource()
    rawdata <- read_html(m[[1]])
    
    #Treeview extraction
    temp_treeview<- rawdata %>%
      html_nodes(xpath = '//*[@id="DivListeYeni"]/table') %>% 
      html_table(fill = TRUE)
  }
  
  ## go to next headers if treeview not found ##
  if(length(temp_treeview)<1){
    next
  }
  temp_treeview<-temp_treeview[[1]]
  temp_treeview<-as.data.frame(temp_treeview)
  
  ##bind treeview dataframes
  Treeview<-bind_rows(Treeview,temp_treeview)
  
  
  #Duty Rates extraction
  duty_tables<- rawdata %>%
    html_nodes(xpath = '//*[@id="DivTab0"]/table') %>% 
    html_table(fill = TRUE,convert = FALSE)
  if(length(duty_tables)<1){
    next
  }
  
  ###===========================================================================###
  
  ##extract duty rates tables(refer 1901 header)
  
  temp_tab<-data.frame()
  for(l in 1:length(duty_tables)){
    ###
    temp_duty<-duty_tables[[l]]
    temp_duty<-as.data.frame(temp_duty)
    
    ####merging 1st, 2nd and 3rd rows(to make them as column names)
    if( (temp_duty[1,1]=="HS Code" & temp_duty[2,1]=="HS Code" & temp_duty[3,1]=="HS Code") | (temp_duty[1,2]=="HS Code" & temp_duty[2,2]=="HS Code" & temp_duty[3,2]=="HS Code") ){
      for(j in 1:ncol(temp_duty)){
        for(k in 2:1){
          if(temp_duty[k,j]!=temp_duty[k+1,j]){
            temp_duty[k,j]<-paste(temp_duty[k,j],temp_duty[k+1,j],sep = "_")
          }
        }
      }
      names(temp_duty)<-temp_duty[1,]
      temp_duty<-temp_duty[-c(1,2,3),]
      
    }else if( (temp_duty[1,1]=="HS Code" & temp_duty[2,1]=="HS Code") | (temp_duty[1,2]=="HS Code" & temp_duty[2,2]=="HS Code") ){
      for(j in 1:ncol(temp_duty)){
        for(k in 1:1){
          if(temp_duty[k,j]!=temp_duty[k+1,j]){
            temp_duty[k,j]<-paste(temp_duty[k,j],temp_duty[k+1,j],sep = "_")
          }
        }
      }
      names(temp_duty)<-temp_duty[1,]
      temp_duty<-temp_duty[-c(1,2),]
    }
    
    ##check if 1st column is empty then remove 1st column
    if(all(temp_duty[,1]=="")){
      temp_duty<-temp_duty[,-1]
    }
    
    ####merging 1st, 2nd and 3rd rows(to make them as column names)(refer 0304 header)
    if( (temp_duty[1,1]=="HS Code/Product Description"  & temp_duty[2,1]=="HS Code/Product Description" & temp_duty[3,1]=="HS Code/Product Description") | (temp_duty[1,1]=="HS Code / Product Description"  & temp_duty[2,1]=="HS Code / Product Description" & temp_duty[3,1]=="HS Code / Product Description") | ( temp_duty[1,2]=="HS Code/Product Description" & temp_duty[2,2]=="HS Code/Product Description" & temp_duty[3,2]=="HS Code/Product Description" ) ){
      temp_duty[,1]<-trimws(temp_duty[,1],which = "both")
      temp_duty[,2]<-trimws(temp_duty[,2],which = "both")
      
      for(j in 1:ncol(temp_duty)){
        for(k in 2:1){
          if(temp_duty[k,j]!=temp_duty[k+1,j]){
            temp_duty[k,j]<-paste(temp_duty[k,j],temp_duty[k+1,j],sep = "_")
          }
        }
      }
      names(temp_duty)<-temp_duty[1,]
      temp_duty<-temp_duty[-c(1,2,3),]
      
    }else if( ( temp_duty[1,1]=="HS Code/Product Description" & temp_duty[2,1]=="HS Code/Product Description") | (temp_duty[1,1]=="HS Code / Product Description"  & temp_duty[2,1]=="HS Code / Product Description") | ( temp_duty[1,2]=="HS Code/Product Description" & temp_duty[2,2]=="HS Code/Product Description") ){
      temp_duty[,1]<-trimws(temp_duty[,1],which = "both")
      for(j in 1:ncol(temp_duty)){
        for(k in 1:1){
          if(temp_duty[k,j]!=temp_duty[k+1,j]){
            temp_duty[k,j]<-paste(temp_duty[k,j],temp_duty[k+1,j],sep = "_")
          }
        }
      }
      names(temp_duty)<-temp_duty[1,]
      temp_duty<-temp_duty[-c(1,2),]
    }
    
    ##check if 1st column is empty then remove 1st column
    if(all(temp_duty[,1]=="")){
      temp_duty<-temp_duty[,-1]
    }
    
    ##split column "HS Code / Product Description"(refer 0304 header)
    if(colnames(temp_duty)[1]=="HS Code / Product Description"){
      
      new_df<-str_split_fixed(temp_duty$`HS Code / Product Description`, " ", 2)
      new_df<-as.data.frame(new_df)
      names(new_df)<-c("HS Code","Product Description")
      
      ##remove column based on column name
      temp_duty <- temp_duty[!names(temp_duty) %in% "HS Code / Product Description"]
      
      temp_duty<-cbind(new_df,temp_duty)
      
    }else if(colnames(temp_duty)[1]=="HS Code/Product Description"){
      
      new_df<-str_split_fixed(temp_duty$`HS Code/Product Description`, " ", 2)
      new_df<-as.data.frame(new_df)
      names(new_df)<-c("HS Code","Product Description")
      
      ##remove column based on column name
      temp_duty <- temp_duty[!names(temp_duty) %in% "HS Code/Product Description"]
      
      temp_duty<-cbind(new_df,temp_duty)
      
    }
    
    ##-----adding text ("The amount to be paid to the Fund Additional") the Footnote column if table contains added text
    temp_cols<-names(temp_duty)
    temp_cols<-as.data.frame(temp_cols)
    temp_col1<-temp_cols[( str_detect(temp_cols[,1],"The amount to be paid to the Fund Additional")),]
    temp_col2<-temp_cols[( str_detect(temp_cols[,1],"Footnote")),]
    
    if(length(temp_col1)>0 & length(temp_col2)>0){
      for(m in 1:nrow(temp_cols)){
        if (str_detect(temp_cols[m,1],"Footnote")){
          temp_cols[m,1]<-paste("The amount to be paid to the Fund Additional",temp_cols[m,1],sep = "_")
          names(temp_duty)<-t(temp_cols)
          break
        }
      }
    }
    
    #####--------------------------
    
    ##-----adding text ("Additional Obligation Charge(% of CIF)") the Footnote column if table contains added text
    temp_cols<-names(temp_duty)
    temp_cols<-as.data.frame(temp_cols)
    temp_col1<-temp_cols[( str_detect(temp_cols[,1],"Additional Obligation Charge")),]
    temp_col2<-temp_cols[( str_detect(temp_cols[,1],"Footnote")),]
    
    if(length(temp_col1)>0 & length(temp_col2)>0){
      for(m in 1:nrow(temp_cols)){
        if (str_detect(temp_cols[m,1],"Footnote")){
          temp_cols[m,1]<-paste("Additional Obligation Charge(% of CIF)",temp_cols[m,1],sep = "_")
          names(temp_duty)<-t(temp_cols)
          break
        }
      }
    }
    
    #####--------------------------
    
    ###
    temp_tab<-bind_rows(temp_tab,temp_duty)
  }
  
  
  ###===========================================================================###
  ##binding dataframes
  final_tab<-bind_rows(final_tab,temp_tab)
  
  print(headers[i])
}

##remove NA
final_tab[is.na(final_tab)] <- ""

fin<-final_tab

#To remove both (NAs and empty)
Treeview <- Treeview[!apply(is.na(Treeview) | Treeview == "", 1, all),]
Treeview[is.na(Treeview)] <- ""

##Remove "Empty Columns"
if (!require("tidyverse")) install.packages("tidyverse")
Treeview<- Treeview %>% discard(~all(is.na(.) | . ==""))

names(Treeview)<-c("HS Code","Product Description")


##write rawdata to excel file
tstamp <- Sys.time()
tstamp <- format(as.POSIXct(tstamp),"%d-%b-%Y_%H-%M")
list_of_datasets <- list("Tree_View" = Treeview, "Tariff_View" = final_tab)
write.xlsx(list_of_datasets, file = paste(getwd(), "/Turkey_Rawdata",tstamp,".xlsx", sep = ''))

########## FORMATTING PART ###############

tariff_data<-final_tab
tree_data<-Treeview

names(tree_data) <- gsub("\\.", " ", names(tree_data))
names(tariff_data) <- gsub("\\.", " ", names(tariff_data))

#Removing NA's from the df table
tariff_data[is.na(tariff_data)] <- ""
tree_data[is.na(tree_data)] <- ""

##replace Dot with Nothing
tree_data[,1]<-gsub(".","",tree_data[,1],fixed = TRUE)
tariff_data[,1]<-gsub(".","",tariff_data[,1],fixed = TRUE)


list <- colnames(tariff_data)
list <- data.frame(list)

list1<-list

##remove "Duty Rates(%)" from column names
for (i in 1:nrow(list1)) {
  if(startsWith(list1[i,1],"Customs Duty Rate") & !str_detect(list1[i,1],"Additional Obligation Charge") & !str_detect(list1[i,1],"Tarim Payi") & !str_detect(list1[i,1],"The amount to be paid to the Fund Additional")){
    list1[i,1]<-gsub(".*_","",list1[i,1])
  }
  if(str_detect(list1[i,1], "Olarak") & str_detect(list1[i,1], "Ek Mali")){
    #remove everything before "_"
    list1[i,1]<-gsub(".*_","",list1[i,1])
    list1[i,1] <- paste("The amount to be paid to the Fund in Euro (equivalent in TL) as agricultural component", list1[i,1], sep = "_")
    
  }
  if(str_detect(list1[i,1], "Additional Obligation Charge")){
    #remove everything before "_"
    list1[i,1]<-gsub(".*_","",list1[i,1])
    list1[i,1] <- paste("Additional Obligation Charge(% of CIF)", list1[i,1], sep = "_")
    
  }
}

names(tariff_data)<-list1[,1]

### MERGE COMMON COLUMNS IF DUPLICATE COLUMNS FOUND
lis <- names(tariff_data)

for (i in 1:ncol(tariff_data)) {
  
  for (j in i+1:ncol(tariff_data)) {
    
    print("first")
    
    if(j<=ncol(tariff_data)){
      
      if(lis[i]==lis[j]){
        
        tariff_data[,i] <- paste(tariff_data[,i],tariff_data[,j],sep = "||")
        
        tariff_data[,j] <- NULL
        
        lis <- names(tariff_data)
        
        print("second")
        
      }
      
    }
  }
}

###########################################



#splitting merged columns based on comma
for (i in 1:nrow(list1)) {
  if(str_detect(list1[i,1],",")){
    if(str_detect(list1[i,1], "Additional Obligation Charge")){
      ##
      reviseddf <-  strsplit(list1[i,1],",")
      reviseddf1 <- reviseddf[[1]]
      reviseddf1 <- trimws(reviseddf1,which="both")
      
      #remove everything before "_"
      reviseddf1 <- gsub(".*_","",reviseddf1)
      reviseddf1 <- paste("Additional Obligation Charge(% of CIF)", reviseddf1, sep = "_")
      ##
    }else if(str_detect(list1[i,1], "The amount to be paid to the Fund")){
      ##
      reviseddf <-  strsplit(list1[i,1],",")
      reviseddf1 <- reviseddf[[1]]
      reviseddf1 <- trimws(reviseddf1,which="both")
      
      #extract everything before "_"
      text <- str_split(reviseddf1[1], fixed("_"))[[1]][1]
      
      #remove everything before "_"
      reviseddf1 <- gsub(".*_","",reviseddf1)
      
      reviseddf1 <- paste(text, reviseddf1, sep = "_")
      ##
    }else{
      reviseddf <-  strsplit(list1[i,1],",")
      reviseddf1 <- reviseddf[[1]]
      reviseddf1 <- trimws(reviseddf1,which="both")
    }
    
    ###############
    for (j in 1:length(reviseddf1)) {
      var1 <- reviseddf1[j]
      var1 <- trimws(var1,which="both")
      
      if(var1 %in% names(tariff_data)){
        
        tariff_data[,var1] <- paste(tariff_data[,var1],tariff_data[,list1[i,1]],sep="||")
        
      }else{
        
        tariff_data[,var1] <- tariff_data[,list1[i,1]]
        
      }
      
      
    }
    
    ###############
    
    tariff_data[,list1[i,1]] <- NULL
    
  }
}


for(r in 1:ncol(tariff_data)){
  tariff_data[,r]<-gsub('^\\||\\|$', '', tariff_data[,r])
  tariff_data[,r]<-gsub('^\\||\\|$', '', tariff_data[,r])
  tariff_data[,r]<-gsub('^\\||\\|$', '', tariff_data[,r])
  tariff_data[,r]<-gsub('^\\||\\|$', '', tariff_data[,r])
  tariff_data[,r]<-gsub('^\\||\\|$', '', tariff_data[,r])
  tariff_data[,r]<-gsub('^\\||\\|$', '', tariff_data[,r])
  tariff_data[,r]<-gsub('^\\||\\|$', '', tariff_data[,r])
  tariff_data[,r]<-gsub('^\\||\\|$', '', tariff_data[,r])
  tariff_data[,r]<-gsub('^\\||\\|$', '', tariff_data[,r])
  tariff_data[,r]<-gsub('^\\||\\|$', '', tariff_data[,r])
  tariff_data[,r]<-gsub('^\\||\\|$', '', tariff_data[,r])
  tariff_data[,r]<-gsub('^\\||\\|$', '', tariff_data[,r])
}


#############################################################################################
######## BEWLOW LINES ARE NEWLY ADDED FORMATTING ############################################

##remove "Product Description" column from tariff_data column
tariff_data <- tariff_data[!names(tariff_data) %in% "Product Description"]

##MERGE TREEVIEW AND DUTYRATES
final_dataframe<-left_join(tree_data,tariff_data,by="HS Code")



######## TO ADD FOUR DIGIT HEADERS #######################

initial <- "0101"

i <- 1

find <- FALSE

while (i <= nrow(final_dataframe)) {
  
  
  dat<- substr(final_dataframe[i,1], start = 1, stop = 4)
  
  if(initial == dat && find == FALSE) {
    
    if(nchar(final_dataframe[i,1])==4){
      
      find <- TRUE
      
      
    }else{
      
      find <- TRUE
      
      if(dat == str_remove(final_dataframe[i,1], "0+$")){
        
        final_dataframe <- add_row(final_dataframe,`HS Code`=dat,"Product Description"=final_dataframe[i,2],.before = i) 
        
      }else{
        
        final_dataframe <- add_row(final_dataframe,`HS Code`=dat,"Product Description"="",.before = i) 
        
      }
      
    }
    
  }
  
  if(initial != "" & initial != dat & dat!="" & dat!="****") {
    
    initial <- dat
    find <- FALSE
    i <- i-1
    
    
  } 
  
  i <- i+1
  
}


###########################################
############################### To add columns dutable and Sysgen ###################

final_dataframe$Dutiable <- ""
final_dataframe$Sysgen <- ""
final_dataframe$comments <- ""
final_dataframe$comments1 <- ""
final_dataframe$comments2 <- ""
final_dataframe$comments3 <- ""
final_dataframe$comments4 <- ""


final_dataframe$`HS Code` <- trimws(final_dataframe$`HS Code`,which = "both")

final_dataframe$Dutiable[ nchar(final_dataframe$`HS Code`)==12] <- "Y"
final_dataframe$Dutiable[ nchar(final_dataframe$`HS Code`)<12] <- "N"

######################################################################################
############################### To generate Sysgen ###################################

final_dataframe$Sysgen[final_dataframe$`HS Code`==""] <- "Y"
final_dataframe$Sysgen[final_dataframe$`HS Code`!=""] <- "N"

######################################################################################

############################### To check the order ###################################

final_dataframe$`HS Code` <- trimws(final_dataframe$`HS Code`,which = "both")


for (i in 1:(nrow(final_dataframe)-1)) {
  
  if(nchar(final_dataframe[i,1])!=0){
    
    if(nchar(final_dataframe[i+1,1])!=0){
      
      
      if(nchar(final_dataframe[i,1]) <= nchar(final_dataframe[i+1,1])){
        
        if(final_dataframe[i,1] > substr(final_dataframe[i+1,1],start = 1,stop = nchar(final_dataframe[i,1]))){
          
          final_dataframe[i,"comments1"] <- "NOT IN PROPER ORDER"
          
        } 
        
      }else{
        
        if(final_dataframe[i+1,1] < substr(final_dataframe[i,1],start = 1,stop = nchar(final_dataframe[i+1,1]))){
          
          final_dataframe[i,"comments1"] <- "NOT IN PROPER ORDER"
          
        }
        
      } 
      
      
    }else{
      
      if(nchar(final_dataframe[i,1]) <= nchar(final_dataframe[i+2,1])){
        
        if(final_dataframe[i,1] > substr(final_dataframe[i+2,1],start = 1,stop = nchar(final_dataframe[i,1]))){
          
          final_dataframe[i,"comments1"] <- "NOT IN PROPER ORDER"
          
        } 
        
      }else{
        
        if(final_dataframe[i+2,1] < substr(final_dataframe[i,1],start = 1,stop = nchar(final_dataframe[i+2,1]))){
          
          final_dataframe[i,"comments1"] <- "NOT IN PROPER ORDER"
          
        }
        
      } 
      
      
    }
    
  }
  
}


############################## TO CHECK SYSGEN RULES #################################

final_dataframe$`Product Description` <- trimws(final_dataframe$`Product Description`,which = "both")

final_dataframe$comments2[ final_dataframe$Sysgen=="Y" & (str_sub(final_dataframe$`Product Description`, - 1, - 1) != ':' )] <- "VERIFY SYSGEN IS NOT AS EXPECTED"

#####################################################################################
############## comment for four digit header start with ' - ' ########################

final_dataframe$comments2[ nchar(final_dataframe$`HS Code`)==4 & (str_detect(substr(final_dataframe$`Product Description`,1,1) ,"-"))] <- "VERIFY HEADER STARTED WITH '-'"

######################################################################################

#generating comments if duplication in immediate description
for(m in 1: nrow(final_dataframe)){
  if(!m==nrow(final_dataframe)){
    if(final_dataframe[m,2]!=""){
      if(final_dataframe[m,1]!=final_dataframe[m+1,1] & final_dataframe[m,2]==final_dataframe[m+1,2]){
        final_dataframe[m,"comments3"]<-"Verify Immediate Description Duplication"
      }
    }
  }
}

#generating comments description empty check
for(m in 1: nrow(final_dataframe)){
  if(!m==nrow(final_dataframe)){
    if(final_dataframe[m,1]!="" & final_dataframe[m,2]==""){
      final_dataframe[m,"comments4"]<-"Verify Description is Empty"
    }
  }
}

##merge "comments1", "comments2", "comments3" and "comments4" columns and save it in "comments" column
##remove "comments1", "comments2", "comments3" and "comments4" columns from final_dataframe dataframe(remove=TRUE)
final_dataframe<- unite(final_dataframe,col='comments'
                        , c('comments1', 'comments2', "comments3", "comments4") , sep = "//", remove = TRUE)


##remove "comments1", "comments2", "comments3" and "comments4" columns from final_dataframe dataframe
#final_dataframe <- final_dataframe[!names(final_dataframe) %in% c("comments1", "comments2", "comments3", "comments4")]

#removing leading and trailing forward slashes(from column 3 to ncol)
for(r in 1:10){
  final_dataframe[,"comments"]<-gsub('^\\//|\\//$', '', final_dataframe[,"comments"]) #same action is repeating multiple times to remove unnecessary forward slashes
}

###newly added lines to merge "BOS-HERZ" and "BOS-HERZ "---(01Jan2023)--------
final_dataframe[is.na(final_dataframe)] <- ""
#check "BOS-HERZ" and "BOS-HERZ " columns available in final_dataframe
if ( (!is.null(final_dataframe$`BOS-HERZ`)) & (!is.null(final_dataframe$`BOS-HERZ `)) ) {
  final_dataframe$`BOS-HERZ`<- paste(final_dataframe$`BOS-HERZ`, final_dataframe$`BOS-HERZ `, sep = "//")
  
  ##remove leading and trailing hyphens(-)
  final_dataframe$`BOS-HERZ`<-gsub('^\\//|\\//$', '', final_dataframe$`BOS-HERZ`)
  final_dataframe$`BOS-HERZ`<-gsub('^\\//|\\//$', '', final_dataframe$`BOS-HERZ`)
  
  #remove "BOS-HERZ " column
  final_dataframe$`BOS-HERZ `<-NULL
}

###----------------------------------------------

#remove leading&trailing spaces from column names
col_names_order<-names(final_dataframe)
col_names_order<-as.data.frame(col_names_order)
col_names_order[,1]<-trimws(col_names_order[,1], which = "both")
names(final_dataframe)<-col_names_order[,1]

##Remove ADDITIONAL or UNWANTED SPACES from all columns
for(i in 2:ncol(final_dataframe)){
  final_dataframe[,i] <- trimws(final_dataframe[,i], which = "both")
  final_dataframe[,i] <- gsub("\\s+", " ", str_trim(final_dataframe[,i]))
}

tstamp <- Sys.time()
tstamp <- format(as.POSIXct(tstamp),"%d-%b-%Y_%H-%M")

list_of_datasets <- list("formatted" = final_dataframe)
write.xlsx(list_of_datasets, file = paste(getwd(), "/Turkey_Tariff_",tstamp,".xlsx", sep = ''))

###---------------------------------------------------------------##
