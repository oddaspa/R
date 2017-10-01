#install.packages('rvest')
#install.packages("ggplot2")

#Loading the packages required
library('rvest')
library('xml2')
library("WriteXLS")

# Set path to xls file
setwd('C:\\Users\\oddgu\\Documents')


# Enter excel as CSV file
data1 <- read.table(file.choose(), header=TRUE, sep=";")


# Creating list for phone numbers to append to excel sheet
phone_number_vec <- c()
for(i in 1:nrow(data1)){
  
  # Setting found data to F
  fd = FALSE
  
  # Getting shareholders
  row <- data1[i, "X1"]
  shareh <- row[1][1]
  first_shareh <- unlist(strsplit(as.character(shareh), " "))
  
  # Construct full name
  url_name <- first_shareh[[1]]
  
  # Constructed url appropriate name
  for(j in 2:length(first_shareh)){
    if(j<3){
      url_name <- paste(url_name, first_shareh[j], sep="%20")
    }
  }
  
  # Finding the area code
  row <- data1[i, "Addresse"]
  addr <- row[1][1]
  first_addr <- strsplit(as.character(addr), " ")
  
  # Trying to get area code
  area <- substring(addr, 1, 4)
  
  # Specifying the url for desired website to be scrapped
  url <- paste('https://www.gulesider.no/person/resultat/',url_name, sep="")
  
  # Reading the HTML code from the website
  webpage <- read_html(url)
  
  # Using the CSS key .hit-address
  rank_data_html <- html_nodes(webpage,'.hit-address')
  
  # Converting the ranking data to text
  rank_data <- html_text(rank_data_html)
  
  # Strips for correct phone
  for(k in 2:length(rank_data)){
    if(grepl(area, rank_data[k])){
      alpha = unlist(strsplit(rank_data[k], "[\n]"))
      phone_number <- alpha[2]
      phone_number_vec <- c(phone_number_vec,phone_number)
      fd <- TRUE
    }
  }
  
  # Failproofing for loss of data
  if(!fd){
    phone_number_vec <- c(phone_number_vec, "Did not find a phone")
  }
  
  # Resetting values
  url <- ""
  url_name <- ""
  webpage <- ""
  phone_number  <- ""
}

# Appending phone number to last array
insert_index <- ncol(data1) + 1
data1[[insert_index]]<- phone_number_vec

# Remaming column to phone number
names(data1)[insert_index] <- "Phone Number"
data1

# To rename excel file simply edit the name currently set as phone_numbers
WriteXLS(data1, ExcelFileName = paste(getwd(), "/phone_numbers.xlsx",sep=""),SheetNames = "Aspaas Oppslagsverk", perl = "perl")

