Jlist1 <- list.files("C:\\Program Files\\Java")
Jfolder <- Jlist1[1]
Sys.setenv('JAVA_HOME' = paste("'C:/Program Files/Java/",Jfolder,"/jre'",sep = ""))

#Sys.setenv('JAVA_HOME' = 'C:/Program Files/Java/jdk1.7.0_45/jre')

library(tibble)
library(dplyr)
library(openxlsx)
library(stringr)
library(gsubfn)
library(stringr)
library(RSelenium)
library(htm2txt)
library(xml2)
library(rvest)


# set working directory
setwd(dirname(rstudioapi::getSourceEditorContext()$path))
getwd()

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
remDr <- driver[['client']]
remDr$maxWindowSize()

##### CUSTOM SCRIPT END

#Function to scrape links from a given URL
scraplinks <- function(url){
  # Create an html document from the url
  webpage <- remDr$getPageSource
  # Extract the URLs
  url_ <- webpage %>%
    rvest::html_nodes("a") %>%
    rvest::html_attr("href")
  # Extract the link text
  link_ <- webpage %>%
    rvest::html_nodes("a") %>%
    rvest::html_text()
  return(tibble(link = link_, url = url_))
}

#Function to merge data
rbind.all.columns <- function(x, y) {
  
  x.diff <- setdiff(colnames(x), colnames(y))
  y.diff <- setdiff(colnames(y), colnames(x))
  
  x[, c(as.character(y.diff))] <- NA
  
  y[, c(as.character(x.diff))] <- NA
  
  return(rbind(x, y))
}

#Import lookup table for chapter HSN
imp_data <- read.xlsx(paste(getwd(),'/INPUT_HSN.xlsx',sep = ""),sheet = 1,startRow = 1)

#nrow(imp_data)
#Creating an empty dataframe
final_table <- data.frame()

#Fire up RemoteDriver using Selenium
weburl <- "https://www.bangladeshtradeportal.gov.bd/index.php?r=tradeInfo/listAll"
#remDr <- remoteDriver(browserName="chrome", port=4444) # instantiate remote driver to connect to Selenium Server
#remDr$open(silent = TRUE) # open web browser
remDr$navigate(url = weburl)
Sys.sleep(5)

#remDr$maxWindowSize()
#Fetching the English page

#langButton <- remDr$findElement(using = 'xpath','//*[@id="bs-example-navbar-collapse-1"]/div[2]/div/a/span[2]')
#langButton$clickElement()

#newOption <- remDr$findElement(using = 'xpath','//*[@id="bs-example-navbar-collapse-1"]/div[2]/div/div/div/ul/li[1]/a/span')
#newOption$clickElement()
Sys.sleep(5)
#b
#b=1
#Giving input chapter for fetching data chapter wise
for (b in 01:nrow(imp_data)){
  remDr$navigate(url = weburl)
  #Giving input chapter 
  sbar <- remDr$findElement(using = 'xpath','/html')
  sbar$sendKeysToElement(list(key = "home"))
  
  radio <- remDr$findElement(using = 'xpath','//*[@id="CommoditySearchForm_searchType_0"]')
  radio$clickElement()
  
  
  inpbox <- remDr$findElement(using = 'xpath','//*[@id="CommoditySearchForm_searchValue"]')
  inpbox$clearElement()
  inpbox$sendKeysToElement(list(imp_data$HSN[b]))
  inpbox$sendKeysToElement(list(key = "enter"))
  Sys.sleep(5)
  
  #Expanding nodes
  sbar <- remDr$findElement(using = 'xpath','/html')
  sbar$sendKeysToElement(list(key = "home"))
  
  try(while (is.null(remDr$findElement(using = 'xpath','//div[@class = "hitarea hasChildren-hitarea expandable-hitarea"]')) == FALSE) {
    exclick <- remDr$findElement(using = 'xpath','//div[@class = "hitarea hasChildren-hitarea expandable-hitarea"]')
    exclick$clickElement()
    Sys.sleep(1)  
  })              
  
  sbar <- remDr$findElement(using = 'xpath','/html')
  sbar$sendKeysToElement(list(key = "home"))
  
  try(while (is.null(remDr$findElement(using = 'xpath','//div[@class = "hitarea hasChildren-hitarea expandable-hitarea lastExpandable-hitarea"]')) == FALSE) {
    lastclick <- remDr$findElement(using = 'xpath','//div[@class = "hitarea hasChildren-hitarea expandable-hitarea lastExpandable-hitarea"]') 
    lastclick$clickElement()  
    Sys.sleep(1)
  })
  
  sbar <- remDr$findElement(using = 'xpath','/html')
  sbar$sendKeysToElement(list(key = "home"))
  
  try(while (is.null(remDr$findElement(using = 'xpath','//div[@class = "hitarea hasChildren-hitarea expandable-hitarea"]')) == FALSE) {
    exclick <- remDr$findElement(using = 'xpath','//div[@class = "hitarea hasChildren-hitarea expandable-hitarea"]')
    exclick$clickElement()
    Sys.sleep(1)  
  })
  
  sbar <- remDr$findElement(using = 'xpath','/html')
  sbar$sendKeysToElement(list(key = "home"))
  
  try(while (is.null(remDr$findElement(using = 'xpath','//div[@class = "hitarea hasChildren-hitarea expandable-hitarea lastExpandable-hitarea"]')) == FALSE) {
    lastclick <- remDr$findElement(using = 'xpath','//div[@class = "hitarea hasChildren-hitarea expandable-hitarea lastExpandable-hitarea"]') 
    lastclick$clickElement()  
    Sys.sleep(1)
  })
  
  sbar <- remDr$findElement(using = 'xpath','/html')
  sbar$sendKeysToElement(list(key = "home"))
  
  try(while (is.null(remDr$findElement(using = 'xpath','//div[@class = "hitarea hasChildren-hitarea expandable-hitarea"]')) == FALSE) {
    exclick <- remDr$findElement(using = 'xpath','//div[@class = "hitarea hasChildren-hitarea expandable-hitarea"]')
    exclick$clickElement()
    Sys.sleep(1)  
  })
  
  sbar <- remDr$findElement(using = 'xpath','/html')
  sbar$sendKeysToElement(list(key = "home"))
  
  try(while (is.null(remDr$findElement(using = 'xpath','//div[@class = "hitarea hasChildren-hitarea expandable-hitarea lastExpandable-hitarea"]')) == FALSE) {
    lastclick <- remDr$findElement(using = 'xpath','//div[@class = "hitarea hasChildren-hitarea expandable-hitarea lastExpandable-hitarea"]') 
    lastclick$clickElement()  
    Sys.sleep(1)
  })
  
  sbar <- remDr$findElement(using = 'xpath','/html')
  sbar$sendKeysToElement(list(key = "home"))
  
  try(while (is.null(remDr$findElement(using = 'xpath','//div[@class = "hitarea hasChildren-hitarea expandable-hitarea"]')) == FALSE) {
    exclick <- remDr$findElement(using = 'xpath','//div[@class = "hitarea hasChildren-hitarea expandable-hitarea"]')
    exclick$clickElement()
    Sys.sleep(1)  
  })
  
  sbar <- remDr$findElement(using = 'xpath','/html')
  sbar$sendKeysToElement(list(key = "home"))
  
  try(while (is.null(remDr$findElement(using = 'xpath','//div[@class = "hitarea hasChildren-hitarea expandable-hitarea lastExpandable-hitarea"]')) == FALSE) {
    lastclick <- remDr$findElement(using = 'xpath','//div[@class = "hitarea hasChildren-hitarea expandable-hitarea lastExpandable-hitarea"]') 
    lastclick$clickElement()  
    Sys.sleep(1)
  })
  
  sbar <- remDr$findElement(using = 'xpath','/html')
  sbar$sendKeysToElement(list(key = "home"))
  
  try(while (is.null(remDr$findElement(using = 'xpath','//div[@class = "hitarea hasChildren-hitarea expandable-hitarea"]')) == FALSE) {
    exclick <- remDr$findElement(using = 'xpath','//div[@class = "hitarea hasChildren-hitarea expandable-hitarea"]')
    exclick$clickElement()
    Sys.sleep(1)  
  })
  
  Sys.sleep(5)
  
  #Scrape webpage links
  source <- remDr$getPageSource()
  webpage <- read_html(source[[1]])
  # Extract the URLs
  url_ <- webpage %>%
    rvest::html_nodes("a") %>%
    rvest::html_attr("href")
  # Extract the link text
  link_ <- webpage %>%
    rvest::html_nodes("a") %>%
    rvest::html_text()
  data <- tibble(link = link_, url = url_)
  
  links <- dplyr::filter(data, grepl("tradeInfo/view&id", data$url))
  
  
  sectionlinks <- paste('https://www.bangladeshtradeportal.gov.bd/',links$url, sep = "")
  sections_final <- as.list(sectionlinks)
  sections_final <- as.character(sections_final)
  
  #remDr$close()  
  #i=1
  #Open all links extracted within chapters
  for (i in 1:length(sections_final)){
    chapurl <- sections_final[i]
    
    remDr$navigate(chapurl)
    
    Sys.sleep(1)
    
    m <- remDr$getPageSource()
    
    webpage <- read_html(m[[1]])
    print(i)
    print(chapurl)
    #chapurl
    #Extraction of tariff table
    extract <- webpage %>%
      html_nodes(xpath = '(//*[@id="list-tariff-grid"]//following::table)[1]') %>%
      html_table(fill = TRUE)
    
    if (length(extract) != 0) {
      tar_table <- extract[[1]]
      tar_table <- tar_table[c(5,1,4,6,7)]  
    }else{
      tar_table <- data.frame()
    }
    Sys.sleep(3)
    #Extraction of HSN Data
    dummy <- webpage %>%
      html_nodes(xpath = '/html/body/div/div[2]/div[3]/div[1]/text()') %>%
      html_text(trim = TRUE)
    
    hsn1 <- dummy[[5]]
    #hsn1
    
    hsn2 <- dummy[[7]]
    #hsn2
    
    hsn3 <- dummy[[9]]
    #hsn3
    
    df <- c(hsn1,hsn2,hsn3)
    #typeof(df)
    df <- as.data.frame(df)
    #typeof(df)
    df <- str_split_fixed(df$df,pattern = " ",n=2)
    df <- as.data.frame(df)
    
    Sys.sleep(2)
    #Appending the extracted data to final table
    if (nrow(final_table) == 0){
      final_table <- rbind(final_table,df) 
    }else{
      final_table <- rbind.all.columns(final_table, df)  
    }
    rnum <- nrow(final_table)
    
    uncol <- 0
    for (u in 1:length(final_table)){
      if (colnames(final_table[u]) == "Unit"){
        uncol <- 1
        break
      }
    }
    
    if (uncol <- 0){
      final_table <- add_column(final_table,Unit = "",.after = length(final_table))
    }
    
    if (nrow(tar_table) != 0){
      for (j in 1:nrow(tar_table)) {
        unit <- trimws(tar_table[j,1])
        act <- trimws(tar_table[j,2])
        rate <- trimws(tar_table[j,3])
        startDate <- trimws(tar_table[j,4])
        endDate <- trimws(tar_table[j,5])
        
        final_table[rnum,3] <- unit
        
        counter <- 0
        for (k in 1:length(final_table)) {
          if (act == colnames(final_table[k])){
            final_table[rnum,k] <- rate
            final_table[rnum,k+1] <- startDate
            final_table[rnum,k+2] <- endDate
            
            #message("COLUMN FOUND")
            counter <- 1
            break
          }
        }
        if (counter == 0){
          final_table <- add_column(final_table, cty = "", .after = length(final_table))
          names(final_table)[length(final_table)] <- act
          final_table[rnum,length(final_table)] <- rate
          
          final_table <- add_column(final_table, sdate = "", .after = length(final_table))
          names(final_table)[length(final_table)] <- paste0(act," Start Date")
          final_table[rnum,length(final_table)] <- startDate
          
          final_table <- add_column(final_table, edate = "", .after = length(final_table))
          names(final_table)[length(final_table)] <- paste0(act," End Date")
          final_table[rnum,length(final_table)] <- endDate
        }
      }
    }
  }
  if(i %% 15==0){
    remDr$refresh()
  }
}


# Remove duplicates based on HSN columns
tab <- final_table
tab <-tab[!duplicated(tab$V1),]

tab <- sapply(tab, as.character)
tab[is.na(tab)] <- ""
tab <- as.data.frame(tab)

#Rename Column Names
names(tab)[1] <- "HSN"
names(tab)[2] <- "Description"
names(tab)[3] <- "Units"

#Adding Length column at First Position
tab <- add_column(tab, 'Length' = "", .before = 1)
tab$HSN <- as.character(tab$HSN)
for (i in 1:nrow(tab)){
  tab$Length[i] <- nchar(tab$HSN[i])  
}


#Adding header HSN's
i = nrow(tab)
try(while(i >= 1) {
  Found <- ""
  if ((tab$Length[i] != "4") & (tab$Length[i] != "0")){
    searchHSN <- substr(tab$HSN[i],1,4)
    
    j = i
    while(j >= 1) {
      if ((searchHSN != substr(tab$HSN[j],1,4)) & (tab$HSN[j] != "")){
        if (searchHSN == tab$HSN[j+1]){
          Found <- 0
          break
        }
        else{
          Found <- 1
          break 
        }
      }
      j=j-1
    }
    
    if  ((Found == 1) & (Found != 0)){
      tab <- add_row(tab,Length = nchar(searchHSN), HSN = searchHSN, Description = tab$Description[j+1],.after = j)
      i= j
    }
    i=j
  }
},silent = TRUE)


#Adding Sysgen column after third Position
tab <- add_column(tab, 'Sysgen' = "", .after =  3)
for (i in 1:nrow(tab)){
  if (tab$Length[i] == "0"){
    tab$Sysgen[i] <- "Y"
  }
  else{
    tab$Sysgen[i] <- "N"
  }
}

#Adding Dutiable column after fourth Position
tab <- add_column(tab, 'Dutiable' = "", .after =  4)

for (i in 1:nrow(tab)){
  if (tab$Length[i] == "8"){
    tab$Dutiable[i] <- "Y"
  }
  else {
    tab$Dutiable[i] <- "N"
  }
}


#Add Comments Column at the last
tab <- add_column(tab, 'Comments' = "", .after =  length(tab))

for(i in 1:nrow(tab)){
  comments <- ""
  
  #Write Sysgen comments
  if (tab$Sysgen[i] == "Y"){
    try(if (startsWith(tab$Description[i],"-") == FALSE){
      comments <- "Verify Sysgen does not begin with hyphen"      
    },silent = TRUE)
    try(if (endsWith(tab$Description[i],":") == FALSE){
      if (comments == ""){
        comments <- "Verify Sysgen does not end with colon" 
      }
      else{
        comments <- paste0(comments,"|Verify Sysgen does not end with colon") 
      }
    },silent = TRUE)
  }
  #Write Dutiable comments
  if (tab$Dutiable[i] == "Y"){
    try(if (tab$Ordinary[i] == ""){
      if (comments == ""){
        comments <- "Verify dutiable HSN does not have rates" 
      }
      else{
        comments <- paste0(comments,"|Verify dutiable HSN does not have rates")
      }
    },silent = TRUE)
  }
  
  tab$Comments[i] <- comments
}

#Removing the length column
tab <- tab[,-c(1)]

#Removing the No results found columns
colLength = length(tab) 
for (i in 1:colLength){
  try(if (grepl("No results",colnames(tab)[i])){
    tab <- tab[,-c(i)]
    colLength = length(tab)
  },silent = TRUE)
}


#Writing the output to excel file
tstamp <- Sys.time()
tstamp <- format(as.POSIXct(tstamp),"%d-%b-%Y_%H-%M")
filepath <- paste(getwd(), "/Bangladesh_Tariff98",tstamp,".xlsx", sep="")
tab= subset(tab, select = -c(7,8,10,11,13,14,16,17,19,20,22,23) )
names(tab) <- c("HSN","DESCRIPTION","Sysgen","Dutiable","UNIT","Customs_Duty(CD)","Supplementary_Duty(SD)","Value_Added_Tax(VAT)","Advance_Income_Tax(AIT)","Advance_Tax(AT)","Regulatory_Duty(RD)","COMMENT")


write.xlsx(tab,file = filepath,rowNames = FALSE)
