# set working directory
setwd(dirname(rstudioapi::getSourceEditorContext()$path))

# load libraries
library(xlsx)
detach("package:xlsx", unload = TRUE)
library(RSelenium)
library(readr)
library(rvest)
library(dplyr)
library(stringr)
library(openxlsx)
library(purrr)
library(abbyyR)
library(tesseract)
library(netstat)

# download Turkish language OCR data
tesseract_download('tur')

# Function to get the HS Code from the page
getHSCode <- function() {
  # get HS Code from the page
  t <- remdr$getPageSource()
  tariffPg <- read_html(t[[1]])
  hscode <<- tariffPg %>%
    html_nodes(xpath = '//*[@id="ctl00_ContentPlaceHolder1_lblBtbNo"]') %>%
    html_text(trim = TRUE)
  if(is_empty(hscode)) {
    hscode <<- ""
  }
}


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


# Function to input HS Number and get the duties page
hsInput <- function() {
  Sys.sleep(format(runif(1,2,5),digits = 1))
  # input HS Number
  suppressMessages(tryCatch(
    expr = remdr$findElement(using = 'id', value = 'ctl00_ContentPlaceHolder1_txtBtbno'),
    error = function(error){
      Sys.sleep(format(runif(1,10,15),digits = 1))
      remdr$findElement(using = 'id', value = 'ctl00_ContentPlaceHolder1_txtBtbno')
      
    }  
  ))
  inputBx <- remdr$findElement(using = 'id', value = 'ctl00_ContentPlaceHolder1_txtBtbno')
  inputBx$clearElement()
  inputBx$sendKeysToElement(list(BTI_list[x]))
  Sys.sleep(format(runif(1,1,2),digits = 1))
  
  # click Search button (BUL)
  searchBut <- remdr$findElement(using = 'id', value = 'ctl00_ContentPlaceHolder1_btnBul')
  searchBut$clickElement()
  pageWait('//*[@id="footer"]')
  Sys.sleep(format(runif(1,2,5),digits = 1))
  
  # open the BTI info page
  remdr$executeScript('javascript:WebForm_DoPostBackWithOptions(new WebForm_PostBackOptions("ctl00$ContentPlaceHolder1$GridView1$ctl02$btn", "", true, "", "", false, true))')
  Sys.sleep(2)
  pageWait('//*[@id="ctl00_ContentPlaceHolder1_btnKapat2"]')
  Sys.sleep(format(runif(1,2,3),digits = 1))
}

# open browser
port <- free_port(random = T)
library(RSelenium)
rd <- rsDriver(port=as.integer(port),browser="firefox", check = F)
remdr <- rd[['client']]

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

# read the BTI List
BTI_list <- read_lines(paste(getwd(),"/bti.txt",sep = ''))

# create a dataframe
BTI_df <- data.frame()

# go to the search page and key in the BTI number
print(length(BTI_list))
for(x in 1:length(BTI_list)) { # LOOP : BTI Input
  
  hsInput() # Input the HS Number
  getHSCode() # get the HS Code

  # check if the data rendered is correct
  while(BTI_list[x] != hscode) {
    remdr$close()
    # open browser
    port <- free_port(random = T)
    library(RSelenium)
    rd <- rsDriver(port=as.integer(port),browser="firefox", check = F)
    remdr <- rd[['client']]
    
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
    
    hsInput()
    getHSCode()
  }
  
  # extract the details
  v <- remdr$getPageSource()
  btiPage <- read_html(v[[1]])
  BTI_num <- btiPage %>%
    html_nodes(xpath = '//*[@id="ctl00_ContentPlaceHolder1_lblBtbNo"]') %>%
    html_text(trim = TRUE)
  if(!is_empty(BTI_num)) {
    HS_num <- btiPage %>%
      html_nodes(xpath = '//*[@id="ctl00_ContentPlaceHolder1_lblGtip"]') %>%
      html_text(trim = TRUE)
    Effective_Date <- btiPage %>%
      html_nodes(xpath = '//*[@id="ctl00_ContentPlaceHolder1_lblGbastar"]') %>%
      html_text(trim = TRUE)
    classification <- btiPage %>%
      html_nodes(xpath = '//*[@id="ctl00_ContentPlaceHolder1_lblSinger"]') %>%
      html_text(trim = TRUE)
    item_Desc <- btiPage %>%
      html_nodes(xpath = '//*[@id="ctl00_ContentPlaceHolder1_lblEstanim"]') %>%
      html_text(trim = TRUE)
    image_link <- btiPage %>%
      html_nodes(xpath = '//*[@id="ctl00_ContentPlaceHolder1_Image1"]') %>%
      html_attr('src') %>%
      paste0('https://uygulama.gtb.gov.tr/BTBBasvuru/', .)
    
    # download image in the 'images' folder
    download.file(image_link, destfile = paste(getwd(),"/images/",hscode,'.jpg',sep = ''), mode = "wb", quiet = TRUE)
    
    # run OCR on the image
    ocrTxt <- suppressMessages(tryCatch(
      expr = ocr(paste(getwd(),"/images/",hscode,'.jpg',sep = ''), engine = tesseract(language = "tur")),
      error = function(error){
        ocrTxt <<- ""
      }  
    ))
    if(ocrTxt == "" || str_detect(ocrTxt, 'BULUNMAMAKTADIR')) {
      imgTag <- 'N'
    } else {
      imgTag <- 'Y'
    }
    
    # create a dataframe
    tempDF <- data.frame(
      BTI_Number = BTI_num,
      HS_Number = HS_num,
      Effective_Date = Effective_Date,
      Classification = classification,
      Item_Description = item_Desc,
      Image_Link = image_link,
      Image = imgTag,
      check.names = FALSE
    )
    
    # bind to the master DF
    BTI_df <- rbind(BTI_df, tempDF)
    
    # click the close button
    closeBut <- remdr$findElement(using = 'xpath', value = '//*[@id="ctl00_ContentPlaceHolder1_btnKapat2"]')
    closeBut$clickElement()
  }
  
  
  print(length(BTI_list) - x)
  
} # /LOOP : BTI Input

# save the BTI records data
BTI_df <- unique(BTI_df) # just to be sure
write.xlsx(BTI_df, paste(getwd(),"/TR_BTI Data ", format(Sys.Date(),'%d-%b-%Y'), ".xlsx",sep = ''))

# add the data to the template file
templateDF <- loadWorkbook(list.files(getwd(), pattern = 'TR_BTI_Update_'))

# add BTI data to the first sheet
writeData(templateDF, 1, BTI_df)

# save the template file
saveWorkbook(templateDF, paste(getwd(),"/TR_BTI_Update_", format(Sys.Date(),'%d%b%Y'), ".xlsx",sep = ''), overwrite = TRUE)
