#install.packages('rvest')

#Loading the packages required
library('rvest')
library('xml2')
library("WriteXLS")
library("RCurl")
# Set path to xls file
setwd('C:\\Users\\oddgu\\Documents')


# Enter excel as CSV file
data1 <- read.csv(file.choose(), header=TRUE, sep=";")

# Creating list for phone numbers to append to excel sheet
phone_number_vec <- c()
shareholders <- c()
for(i in 1:nrow(data1)){
  
  # Setting found data to F
  fd = FALSE
  is_AS = FALSE
  # Getting shareholders
  row <- data1[i, "shareholders"]
  shareh <- row[1][1]
  first_shareh <- unlist(strsplit(as.character(shareh), " "))
   # Construct full name
  if(is.na(first_shareh[1])){
    phone_number_vec <- c(phone_number_vec, "Did not get input")
    first_shareh <- paste(first_shareh, collapse = ' ')
    shareholders <- c(shareholders, first_shareh)
    next
  }
  url_name <- first_shareh[1]
  # Constructed url appropriate name
  for(j in 2:length(first_shareh)){
    if(first_shareh[j]== "AS" | first_shareh[j] == "ASA"){
      phone_number_vec <- c(phone_number_vec, "Not implemented for AS/ASA types")
      first_shareh <- paste(first_shareh, collapse = ' ')
      shareholders <- c(shareholders, first_shareh)
      fd = TRUE
      is_AS = TRUE
      break
    }
    else if(nchar(first_shareh[j])>1){
      url_name <- paste(url_name, first_shareh[j], sep="%20")
    }
  }
  # if we got an AS we skip to next iteration
  if(is_AS){
    next
  }
  # need to correct for Æ Ø Å errors from the csv file
  # UTF-8  
  url_name <- gsub("®","AE",url_name)
  url_name <- gsub("¯","OE",url_name)
  url_name <- gsub("Å","AA",url_name)
  url_name <- gsub("f","E",url_name)

  # Finding the area code
  row <- data1[i, "Addresse"]
  addr <- row[1][1]
  first_addr <- strsplit(as.character(addr), " ")
  
  # Trying to get area code
  area <- substring(addr, 1, 4)
  
  # Specifying the url for desired website to be scrapped
  urlname <- paste('https://www.gulesider.no/person/resultat/',url_name, sep="")
  
  # Handling exceptions
  if(!url.exists(urlname)){
    if(is_AS){
      next
    }else{
    print("URL does not exist!")
    phone_number_vec <- c(phone_number_vec, "Did not find a phone")
    first_shareh <- paste(first_shareh, collapse = ' ')
    shareholders <- c(shareholders, first_shareh)
    next
    }
  }
  
  # Reading the HTML code from the website
  webpage <- read_html(urlname)
  
  # Using the CSS key .hit-address
  rank_data_html <- html_nodes(webpage,'.hit-address')
  
  # Converting the ranking data to text
  rank_data <- html_text(rank_data_html)
  
  # Strips for correct phone
  for(k in 1:length(rank_data)){
    if(grepl(area, rank_data[k])){
      alpha = unlist(strsplit(rank_data[k], "[\n]"))
      phone_number <- alpha[2]
      phone_number_vec <- c(phone_number_vec,phone_number)
      fd = TRUE
      break
    }
  }
  
  # Failproofing for loss of data
  if(!fd){
    print("failproof for loss of data")
    phone_number_vec <- c(phone_number_vec, "Did not find a phone")
  }
  
  # need to correct for Æ Ø Å errors from the csv file
  # UTF-8  
  first_shareh <- paste(first_shareh, collapse = ' ')
  first_shareh <- gsub("®","AE",first_shareh)
  first_shareh <- gsub("¯","OE",first_shareh)
  first_shareh <- gsub("Å","AA",first_shareh)
  first_shareh <- gsub("f","E",first_shareh)
  shareholders <- c(shareholders, first_shareh)

  # Resetting values
  url <- ""
  url_name <- ""
  webpage <- ""
  phone_number  <- ""
}

# Appending phone number to last array
insert_index <- ncol(data1) + 1
data1[[insert_index]]<- phone_number_vec

# Appending reformatted names on first array
data1$shareholders <- shareholders

# Remaming column to phone number
names(data1)[insert_index] <- "Phone Number"


# To rename excel file simply edit the name currently set as phone_numbers
WriteXLS(data1, ExcelFileName = paste(getwd(), "/phone_numbers.xlsx",sep=""),SheetNames = "Aspaas Oppslagsverk", perl = "perl", Encoding = ("latin1"))

