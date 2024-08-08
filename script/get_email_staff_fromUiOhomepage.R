#################
## This script scrapes the staff list from the Department of Sociology and Human Geography, University of Oslo
#  Make sure the following is correct: 
#   1. The URL to staff page is correct
#   2. The number of pages to scrape is correct
#   2. The xpath to the elements are correct
#
# Note: I've had great help of gpt.uio.no to sort this code out
#################

invisible(Sys.setlocale(locale='no_NB.utf8'))
library(stringr)
library(httr)
library(XML)
library(stringr)
library(ggplot2)
library(rvest)
library(tidyverse)


# Loops through pages, scrape name, email and title for each person. 
person_data <- list()
for(i in seq(1,6)){
  url <- paste("http://www.sv.uio.no/iss/personer/?page=", i, sep="")
  UrlPage <- read_html(url) 
  person_data[[i]] <- UrlPage %>%
    html_elements(xpath = '//tr[starts-with(@class, "vrtx-person")]') %>%
    map_df(~{
      tibble(
        name = html_element(.x, xpath = './td[contains(@class, "vrtx-person-listing-name")]/a[not(contains(@class, "vrtx-image"))]') %>% html_text(trim = TRUE),
        email = html_element(.x, xpath = './td[contains(@class, "vrtx-person-listing-email")]/a') %>% html_attr("href") %>% sub('mailto:', '', .),
        title = html_element(.x, xpath = './td[contains(@class, "vrtx-person-listing-name")]/span') %>% html_text(trim = TRUE),
        
      )
    })
  
}

ansattliste <- bind_rows(person_data) %>% 
  mutate(etternavn = str_split(name, ",") %>% map_chr(~.x[1]),
         fornavn = str_split(name, ",") %>% map_chr(~.x[2]) %>% str_trim()) %>% 
  mutate(navn = paste(fornavn, etternavn)) %>% 
  select(navn, name, email, title) %>% 
  rename(navn2 = name)
         

ansattliste %>% 
  filter(str_detect(navn, "Trude"))



# Save the data

write.xlsx(ansattliste, "data/ansattliste.xlsx")

ansattliste <- read.xlsx("data/ansattliste.xlsx")



rett_navn <- read_excel("data/navn_variasjoner.xlsx")


ansattliste2 <- ansattliste %>% 
  left_join(rett_navn, by = c("navn" = "navn")) %>% 
  mutate(navn = ifelse(is.na(riktig_navn), navn, riktig_navn)) %>% 
  select(navn, email, title) %>% 
  distinct() %>% 
  arrange(navn) %>% 
  filter( !str_detect( tolower(title), "rådgiver|konsulent|sjef|emerit"))

ansattliste2 %>% 
  filter(str_detect(navn, "Trude"))

# Create file for sending emails

library(tidyverse)
library(readxl)
library(openxlsx)

write.xlsx( data.frame(navn = personer_stab), "personrapport/send_epost_alle.xlsx")

filnavn <- paste0("personrapport/send_epost_alle.xlsx")

wb <- loadWorkbook(filnavn)
addWorksheet(wb, "epost")

writeData(wb, "epost", x = ansattliste2)

#writeFormula(wb, sheet = "Sheet 1", x = "=XLOOKUP(A1;epost!A:A;epost!B:B)", startCol = 2, startRow = 2)

saveWorkbook(wb, filnavn, overwrite = TRUE)




