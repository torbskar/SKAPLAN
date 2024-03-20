
invisible(Sys.setlocale(locale='no_NB.utf8'))

library(tidyverse)
library(readxl)
library(openxlsx)


## Oversikt over person #### 
source("script/stab_andreFolk.R")


faste <- stab %>% 
  filter(str_sub(Stillingsgruppe, 1, 4) %in% c("1013", "1011", "1475", "1198")) %>% 
  mutate(etternavn = word(navn, -1)) %>% 
  select(navn, etternavn) %>% 
  arrange(etternavn) %>% 
  group_by(etternavn) %>% 
  # mutate(likenavn = n()) %>% 
  # arrange(-likenavn) %>% 
  slice(1) %>% 
  ungroup()


tr_mappe <- "C:/Users/torbskar/Universitetet i Oslo/Timeregnskap ved ISS - General/" 

forskningstermin <- read_xlsx( paste0(tr_mappe, "Forskningsterminer_ISS.xlsx"), skip = 1) %>% 
  janitor::clean_names() %>% 
  select(v23:ncol(.)) %>% 
  rename("forsktilgode" = colnames(.)[ncol(.)-3],
         "etternavn" = colnames(.)[ncol(.)-2],
         "kommentar1" = colnames(.)[ncol(.)-1],
         "kommentar2" = colnames(.)[ncol(.)] ) %>% 
  mutate(etternavn = word( str_remove(etternavn, ","))) %>% 
  mutate(etternavn = ifelse(etternavn == "García-Godos", "Naveda", etternavn)) %>% 
  left_join(faste, by = "etternavn") %>% 
  select(-c(kommentar1, kommentar2)) %>%
  filter(!is.na(forsktilgode)) %>% 
  select(-   colnames(.)[str_detect(colnames(.), "_")]  ) %>% 
  pivot_longer(-c(navn, etternavn, forsktilgode), names_to = "semester", values_to = "f_termin") %>% 
  filter(!is.na(f_termin), f_termin < 0)


save(forskningstermin, file = "data/forskningstermin.Rdat")

