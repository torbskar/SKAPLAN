

## Kjører ut rapporter 
library(tidyverse)
library(readxl)

rm(list=(ls()))


## Log fil
sink("log.Rtxt", append = TRUE)
  print( paste("Sist rapport kjørt:", Sys.time()))
sink()


## Create file structure 
if(dir.exists("personrapport/")){ 
  files2zip <- dir('personrapport', full.names = TRUE)
  zip(zipfile = paste0('arkiv/personrapport_', Sys.Date()), files = files2zip)
  unlink("personrapport/", recursive = TRUE)
  dir.create("personrapport")
} else {
  dir.create("personrapport")
}



#source("hente_data.R")

#dato <- Sys.Date()


# Leser inn data
budsjett <- read_rds("data/timebudsjett.rds") %>% 
  filter(!is.na(timer))


## Oversikt over person

stab <- read_excel("stab.xlsx") %>%   # Fil med alle i stab
  select(navn, navn2, navn3) %>% 
  mutate(navn3 = ifelse(navn == navn3, NA, navn3))  

staff <- unique(c(stab$navn, stab$navn2, stab$navn3)) %>% str_squish()

personer <- budsjett %>% 
  filter(navn != "xxx") %>% 
  group_by(navn) %>% 
  slice(1) %>% 
  pull(navn) %>% 
  str_squish()

# Hvem som ikke er i stab
personer[!(personer %in% staff) ]



pers2 <- stab %>%  
  select(navn, navn2) %>% 
  filter(!is.na(navn2))

pers3 <- stab %>% 
  select(navn, navn3) %>% 
  filter(!is.na(navn3))

# Omkoder 3 stavemåter av navn
budsjett <- left_join(budsjett, pers2, join_by(navn == navn2)) %>%
  mutate(navn = ifelse(!is.na(navn.y), navn.y, navn)) %>%
  select(-navn.y) %>% 
  left_join(pers3, join_by(navn == navn3)) %>%
  mutate(navn = ifelse(!is.na(navn.y), navn.y, navn)) %>% 
  select(-navn.y)  
  #filter(navn == "Kjell Kjellman" | navn == "Kjell Erling Kjellman" | navn == "Kjell E. Kjellman")

personer <- budsjett %>% 
  filter(navn != "xxx") %>% 
  group_by(navn) %>% 
  slice(1) %>% 
  pull(navn) %>% 
  str_squish()


# selecterer bare stab
personer <- personer[personer %in% stab$navn]

#personer <- "Axel Petter Kristensen"

pathdir <- paste0(getwd(), "/personrapport/")

for(person in personer){
  print(person)
  
  persondata <- budsjett %>% 
    filter(navn == person)
  
  saveRDS(persondata, file="data/timebudsjett_person.rds")
  
  ## FUll report
  filnavn <- paste(person, ".pdf", sep='')
  quarto::quarto_render(input = "personrapport.qmd", 
                    output_file = filnavn) 

  file.copy(from = paste0(here::here(), "/", filnavn), 
            to = paste0(here::here(), "/personrapport/", filnavn))
  file.remove(paste0(here::here(), "/", filnavn))
  
  #rm(list=ls()[!(ls() %in% c("person", "personer", "pathrmd", "pathdir", "wd", "dato"))])
  #file.remove("emne.RDat")
}


# 

# # Remove all old files
# files <- list.files()[!(list.files() %in% c(list.files(pattern = ".R"), list.files(pattern = ".zip")) )]
# unlink(files, recursive = FALSE)
