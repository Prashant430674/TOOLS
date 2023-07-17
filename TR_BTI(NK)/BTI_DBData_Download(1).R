#dealing with Java memory space error
options(java.parameters = "- Xmx7024m")

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
driver <- rsDriver(browser = 'chrome', port = free_port(random = T), chromever = realChromeVer[1], verbose = F, check = T)
remdr <- driver[['client']]
remdr$closeWindow()

# function to launch session
launchSession <- function() {
  remdr$open()
  remdr$maxWindowSize()
}

# function to close session
closeSession <- function() {
  remdr$closeWindow()
}

launchSession()
##### CUSTOM SCRIPT END

# Set Working Directory #####
setwd(dirname(rstudioapi::getSourceEditorContext()$path))

# Load Library ####
library(rvest)
library(tidyverse)
library(xlsx)

# Define Functions ####

# Function to navigate
goToURL <- function(url, xpath = '/html/body') {
  pgStatus <- NULL
  while(is.null(pgStatus)) {
    tryCatch(
      remdr$navigate(url = url),
      pageWait(xpath),
      pgStatus <- 'OK', 
      error = function(error) {
        Sys.sleep(5)
        pgStatus <- NULL
      }
    )
  }
}

# Function to wait until a specific text is available in page
waitUntilText <- function(XpathVal, TextToWait) {
  mytext <- ''
  while(mytext != TextToWait) {
    tempEle <- remdr$findElement(using = 'xpath', value = XpathVal)
    mytext <- unlist(tempEle$getElementText())
  }
}

# Function for pagewait
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

# Go to landing Page ####
goToURL('http://pla11038.amberroad.com:7060/CONTENT/mdi/html/desktop/login.jsp', '//*[@id="user_id"]')
# input username and password
ID_INPUT <- remdr$findElement(using = 'id',value = 'user_id' )
ID_INPUT$sendKeysToElement(list("CURRENCY_USER"))
PW_INPUT <- remdr$findElement(using = 'id',value = 'password' )
PW_INPUT$sendKeysToElement(list("password", key = "enter"))
Sys.sleep(runif(1,2,3))
waitUntilText('//*[@id="titleTD"]', 'My Dashboard')
# click on Hamburger menu
menuBut <- remdr$findElement(using = 'xpath', value = "//button[@title='View menu items']")
menuBut$clickElement()
Sys.sleep(runif(1,2,3))
# click on Utility
CLK_UTILITY <- remdr$findElement(using = 'xpath',value = "//button[@value='Utility']")
CLK_UTILITY$clickElement()
Sys.sleep(runif(1,2,3))
# click on Master Data Extraction
CLK_MDE <- remdr$findElement(using = 'xpath',value = "//div[@id='desktop-header__menu-favorites-row']/nav[1]/div[1]/section[1]/div[1]/ul[1]/li[3]/a[1]")
CLK_MDE$clickElement()
Sys.sleep(runif(1,2,3))
waitUntilText('//*[@id="titleTD"]', 'Master Data Extraction')
# switching frames
mainFrame <- remdr$findElement(using = 'id', value = 'mainFrame')
remdr$switchToFrame(mainFrame)
Sys.sleep(runif(1,2,3))

# enter required data in the fields
remdr$findElement(using = 'xpath', value = '//*/option[@value="GCQA"]')$clickElement()
Sys.sleep(runif(1,2,3))
remdr$findElement(using = 'xpath', value = '//*/option[@value="Binding Tariff Information"]')$clickElement()
Sys.sleep(runif(1,2,3))
remdr$findElement(using = 'xpath', value = '//*/option[@value="LCS_BTI_DESC"]')$clickElement()
Sys.sleep(runif(1,2,3))
remdr$findElement(using = 'xpath', value = "(//td[text()='Country ID']/following::option[@value='TR'])[1]")$clickElement()
Sys.sleep(runif(1,2,3))
remdr$findElement(using = 'xpath', value = "(//td[text()='Language Cd']/following::option[@value='TR'])[1]")$clickElement()
Sys.sleep(runif(1,2,3))

label6 <- remdr$findElement(using = 'xpath', value = "//td[text()='Date']/following::input")
label6$sendKeysToElement(list(format(Sys.Date(),'%d-%b-%Y')))
Sys.sleep(runif(1,2,3))

# click Extract button
extractBut <- remdr$findElement(using = 'id', value = 'extractbutton')
extractBut$clickElement()
Sys.sleep(runif(1,2,3))
# wait for the download success alert
downloadRes <- 0
while(downloadRes == 0) {
  tryCatch({
    expr = remdr$getAlertText()
    downloadRes <<- 1
  },
  error = function(error) {
    downloadRes <<- 0
  }
  )
}
if(downloadRes == 1) {
  remdr$acceptAlert()
  Sys.sleep(runif(1,2,3))
}
# get file list before downloading the file
default_dir <- file.path("", "Users", Sys.info()[["user"]], "Downloads")
files_Before <- list.files(default_dir, pattern = '*.xlsx')
# get file list again
files_After <- list.files(default_dir, pattern = '*.xlsx')

# get the downloaded file name
reqFile <- setdiff(files_After, files_Before) 
# click Download button
dwldBut <- remdr$findElement(using = 'id', value = 'downloadbutton')
dwldBut$clickElement()
Sys.sleep(runif(1,2,3))
# wait until the new file is available in the downloads folder


while(length(reqFile) == 0) {
  Sys.sleep(10)
  
  # get file list 
  files_After <- list.files(default_dir, pattern = '*.xlsx')
  
  # get the new file
  reqFile <- setdiff(files_After, files_Before)
  
}

# signout the browser and close the session
remdr$switchToFrame(NA)
clickProfile <- remdr$findElement(using = 'xpath', value = '//*[@id="desktop-user"]/span/span[1]')
clickProfile$clickElement()
Sys.sleep(runif(1,2,3))

signOut <- remdr$findElement(using = 'id', value = 'desktop-user__signout')
signOut$clickElement()
Sys.sleep(runif(1,2,3))
remdr$close()

# read the downloaded file
myDF <- read.xlsx(paste0(default_dir,'/',reqFile), sheetIndex = 1, encoding="UTF-8")
for(vi in 1:ncol(myDF)) {
  myDF[, vi] <- as.character(myDF[, vi])
}
# add first row as headers 
names(myDF) <- myDF[1, ]
myDF <- myDF[-1, ]

# save the DB data somewhere
myDF[is.na(myDF)] <- ""
write.xlsx(myDF, paste('DB_Data ', format(Sys.Date(),'%d-%b-%Y'),'.xlsx', sep = '' ), row.names = FALSE)
