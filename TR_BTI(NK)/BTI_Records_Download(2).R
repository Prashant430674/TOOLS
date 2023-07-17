# set working directory
setwd(dirname(rstudioapi::getSourceEditorContext()$path))

# load libraries
library(xlsx)
detach("package:xlsx", unload = TRUE)
library(rvest)
library(dplyr)
library(stringr)
library(openxlsx)
library(purrr)
library(netstat)

# function for pagewait
pageWait <- function(a) {
  for(v in 1:15) {
    pageCheck <- 0
    Sys.sleep(1)
    x <- remdr$getPageSource()
    funcData <- read_html(x[[1]])
    funCheck <- funcData %>%
      html_nodes(xpath = a) %>%
      html_text(trim = TRUE)
    if(length(funCheck) > 0) {
      #print("page load : success")
      #print(funCheck)
      pageCheck <<- 1 
      break()
      
      
    }
    if(v == 15) {
      #print("page load : fail")
      pageCheck <<- 0
    }
  }
}

# function to extract the current table data
extracTable <- function() {
  x <- remdr$getPageSource()
  xPage <- read_html(x[[1]])
  xTable <- xPage %>%
    html_nodes(xpath = '//*[@id="ctl00_ContentPlaceHolder1_GridView1"]') %>%
    html_table(fill = TRUE)
  myTbl <- xTable[[1]]
  myTbl <- myTbl[which(substring(myTbl$DETAY, 1, 2) == 'TR'), ]
  myTbl <<- myTbl[, c(1:4)]
}

# open browser
port <- free_port(random = T)
library(RSelenium)
rd <- rsDriver(port=14421L,browser="firefox",chromever = "110.0.5481.30")
remdr <- rd[['client']]

rdriver <- rsDriver(browser = "firefox",
                    chromever = "108.0.5359.71",
                    verbose = TRUE,
                    port = free_port())
remdr <- rdriver[['client']]
remdr$maxWindowSize()
# Go to the source link
remdr$navigate("https://uygulama.gtb.gov.tr/BTBBasvuru/AnaSayfa")
Sys.sleep(format(runif(1,2,5),digits = 1))

# click the first link (Bağlayıcı Tarife Bilgisi Sorgulama İşlemleri)
firstClick <- remdr$findElement(using = 'id', value = 'LinkButton1')
firstClick$clickElement()

# click Search button (BUL)
searchBut <- remdr$findElement(using = 'id', value = 'ctl00_ContentPlaceHolder1_btnBul')
searchBut$clickElement()
pageWait('//*[@id="footer"]')
Sys.sleep(5)

# get the total record count
c <- remdr$getPageSource()
pageData <- read_html(c[[1]])
resCount <- pageData %>%
  html_nodes(xpath = '//*[@id="ctl00_ContentPlaceHolder1_lblCount"]') %>%
  html_text(trim = TRUE) %>%
  str_extract(., '\\d+')

# get the total pages count (each page has 16 records except the very last page)
totPages <- ceiling(as.numeric(resCount)/16)

# create a DF
fDF <- data.frame()

# get a dummy DF for checking purpose
checkTbl <- mtcars[c(1:1), c(1:4)]

# set same names as original table
extracTable()
names(checkTbl) <- names(myTbl)
fDF <- rbind(fDF, checkTbl)
# go to each page and extract the table data
for(i in 1:totPages) {
  
  # execute the script
  myscript <- paste("javascript:__doPostBack('ctl00$ContentPlaceHolder1$GridView1",
                    "','Page$",
                    i,
                    "')",
                    sep = '')
  nextPg <- remdr$executeScript(myscript) # execute the javascript available in page
  Sys.sleep(runif(1,5,10))
  
  # extract the current table
  extracTable()
  
  while(as.character(fDF[nrow(fDF), 1]) == as.character(myTbl[nrow(myTbl), 1])) {
    
    nextPg <- remdr$executeScript(myscript)
    Sys.sleep(runif(1,5,10))
    extracTable()
    message('Clicking once again to avoid duplicates')
  }
  
  # bind to the DF
  fDF <- rbind(fDF, myTbl)
  fDF <- unique(fDF)
  message(paste("captured ", nrow(fDF)-1," of ", resCount," records",sep = ''))
  
  
}

# remove the dummy data (mtcars)
fDF <- fDF[-c(1:1), ]

DF <- unique(fDF) # just to be sure

# save the total records data
write.xlsx(fDF, paste(getwd(),"/TR_BTI Records ", format(Sys.Date(),'%d-%b-%Y'), ".xlsx",sep = ''), row.names = FALSE)
