# Oppdater med stab

invisible(Sys.setlocale(locale='no_NB.utf8'))

library(tidyverse)
library(readxl)
library(openxlsx)


## Oversikt over stab siste år #### 
stab0 <- read_excel("data/stab.xlsx") %>% 
  mutate(navn = `Etter- og fornavn`, 
         Fratredelsesdato = ifelse(Fratredelsesdato == "-", NA_character_, Fratredelsesdato),
         sluttdato = openxlsx::convertToDate(Fratredelsesdato)) %>% 
  mutate(navn = paste(str_squish( str_split_i(navn, ",", 2) ),
                      paste0(str_squish( str_split_i(navn, ",", 1) ))
  )) %>% 
  select(navn, Stillingsgruppe, Dellønnsprosent, sluttdato) %>% 
  mutate(int_ekst = "intern")

# Nyansatte som ikke har begynt ennå etc. og ikke finnes i andre filer
stab_nye <- readxl::read_excel("data/stab_nye.xlsx") %>% 
  arrange(navn) %>% 
  mutate(int_ekst = "intern")

stab <- bind_rows(stab_nye, stab0) %>% 
  mutate(navn = str_to_title(navn)) %>% 
  mutate(navn = str_replace(navn, "' ", "'")) %>% 
  group_by(navn) %>% 
  slice(1)            # Fjerner dubletter - hvis folk er i flere filer 

#save(stab, file = "data/stab.Rdat")



## Oversikt over eksterne undervisere og andre navn til lister ####
andrefolk <- readxl::read_excel("data/eksterne_undervisere.xlsx") %>% 
  arrange(navn) %>% 
  select(navn) %>% 
  mutate(navn = str_to_title(navn)) %>% 
  mutate(navn = str_replace(navn, "' ", "'")) 

# loop over csv-filer i data-mappen som starter med TP  
# Oppdateres herfra: https://tp.educloud.no/uio/report/ velg fagpersonperspektiv
# Åpne i Notepad og lagre som UTF-8
csvfiler <- list.files("data")[startsWith(list.files("data"), "TP_personrapport")]
TP <- list()
j <- 0
for(i in csvfiler){
  #print(i)
  j <- j + 1
  TP[[j]] <- read.csv(paste0("data/", i), sep = ";", 
                      fileEncoding = "UTF-8",
                      encoding = "UTF-8") %>% 
    arrange(Fagperson) %>% 
    select(Fagperson) %>% 
    mutate(navn = paste(str_squish( str_split_i(Fagperson, ",", 2) ),
                        paste0(str_squish( str_split_i(Fagperson, ",", 1) )))) %>% 
    select(navn) %>% 
    mutate(navn = str_replace(navn, "' ", "'")) 
}

TPfolk <- bind_rows(TP) %>% unique() %>% as_tibble()

#View(TPfolk)

# TPfolk %>% 
#   filter(str_detect(navn, "David"))


navneliste <- bind_rows(stab, andrefolk, TPfolk) %>% 
  mutate(navn = str_squish(navn)) %>%  # fjern whitespace og dubletter
  group_by(navn) %>% 
  slice(1) %>% 
  ungroup() %>%
  mutate(fornavn = str_split_i(navn, " ", 1),     # fjerne dublett på for- og etternavn. Beholder lengste navn
         etternavn = str_split_i(navn, " ", -1), 
         antsplit = str_count(navn, " ")) %>% 
  arrange(etternavn, fornavn, desc(antsplit)) %>% 
  group_by(fornavn, etternavn) %>%
  slice(1) %>% 
  ungroup() %>%
  mutate(sort = ifelse(navn == "Annen", 0, 1),
         sluttdato = replace_na(sluttdato, as.Date("9999-01-01", format = "%Y-%m-%d"))) %>% 
  arrange( desc(int_ekst), navn, desc(sluttdato)) %>%   # hvis dublett: intern før ekstern 
  arrange( is.na(int_ekst)) %>% 
  group_by(navn) %>% 
  slice(1) %>% 
  ungroup() %>%
  arrange(sort, navn) %>%
  select(-sort) %>% 
  mutate(int_ekst = ifelse(is.na(int_ekst), "ekstern", int_ekst)) %>% 
  as_tibble() %>% 
  select(navn) %>% 
  mutate(navn = case_when(navn == "Michael Gentile" ~ "Michael Paul Gentile", 
                          navn == "Magne Flemmen" ~ "Magne Øyvind Paalgard Flemmen", 
                          navn == "Andrea Nightingale" ~ "Andrea Joslyn Nightingale",
                          navn == "Jemima Garcia-Godos" ~ "Jemima Victoria Garcia-Godos Naveda",
                          navn == "Trude Lappegaard" ~ "Trude Lappegård",
                          tolower(navn) %in% c("karen o'brien", 
                                               "karen o' brien",
                                               "karen linda o' brien") ~ "Karen Linda O'Brien", 
                          
                          TRUE ~ navn))

save(navneliste, file = "data/navneliste.Rdat")

