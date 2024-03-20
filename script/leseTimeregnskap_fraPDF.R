#################################;
# Lese inn fra timeregnskap  ####
# INNFILER: fra en mappe med alle timeregnskapsrapporter 
#           Ved ISS ligger disse i separate mapper for faste og midlertidige 
#           slik at det er lett å skille mellom disse.
#           For midlertidige trengs en fil med gjenstående saldo fra forrige semester. Egen Excel-fil fra admin!! 
# INNDATA :  "N:/iss-admfelles/Timeregnskap/2023 Høst/PDF FAST"
#            "N:/iss-admfelles/Timeregnskap/2023 Høst/PDF MIDLERTIDIG"
#             data/saldo_V23.xlsx
# Source  : "script/stab_andreFolk.R"
#################################;

invisible(Sys.setlocale(locale='no_NB.utf8'))

library(pdftools)
library(tidyverse)


source("script/stab_andreFolk.R")


# fast ansatte ####
baneF <- "N:/iss-admfelles/Timeregnskap/2023 Høst/PDF FAST"

filerF <- list.files(baneF)
filerF <- tibble(filerF)$filerF 


# funksjon for å trekke ut nummer fra filnavn
parse_number2 <- function(x){
  x %>% 
    str_extract("\\d{4}") %>% 
    as.numeric()
}

filerF <- tibble(filerF) %>% 
  filter(str_detect(filerF, "VÅR|HØST")) %>% 
  mutate(sort = ifelse(str_detect(filerF, "oppdatert"), 1, 9),
         id = str_sub(filerF,1,12), 
         aar = parse_number2(filerF),
         semester = ifelse(str_detect(filerF, "HØST"), 2, 1)) %>% 
  arrange(sort) %>% 
  group_by(id) %>%
  slice(1) %>% 
  pull(filerF)

fF <- paste0(baneF, "/", filerF)


fil_opps <- tibble(filer = filerF) %>% 
  mutate(total = 0, 
         navn = NA_character_)


#View(fil_opps)

# Loop gjennom alle pdf'er for faste, plukk ut totaltall
j = 0
#i <- fF[1]
for(i in fF){
  if(i == fF[1]){ 
    j = 1} else{j = j+1}
  text <- pdf_text(i)
  raw <- map(text, ~ str_split(.x, "\\n") %>% unlist()) %>% 
    reduce(., c)
  totalrader <- stringr::str_which(tolower(raw), "totalt")
  fil_opps$total[j] <- parse_number(raw[totalrader][2], 
                                    locale = locale(decimal_mark = ",", 
                                                    grouping_mark = ".")  )
  fil_opps$navn[j] <- raw[8] %>% str_squish()
}

saldo_faste <- fil_opps %>% 
  mutate(aar = parse_number2(filer),
         semester = ifelse(str_detect(filer, "HØST"), "H", "V")) %>% 
  arrange(navn, desc(aar), semester) %>%  #sortering gir riktig semester ut ved slice
  group_by(navn) %>% 
  slice(1) %>% 
  mutate(aar = paste(aar, semester), 
         timer = total) %>% 
  mutate(emne = "Faste poster", aktivitet = paste("SALDO", aar)) %>% 
  select(navn, timer, aar, emne, aktivitet) %>% 
  mutate(navn = case_when(navn == "Magne Flemmen" ~ "Magne Øyvind Paalgard Flemmen", 
                          navn == "Andrea Nightingale" ~ "Andrea Joslyn Nightingale",
                          navn == "Jemima Garcia-Godos" ~ "Jemima Victoria Garcia-Godos Naveda",
                          navn == "Trude Lappegaard" ~ "Trude Lappegård",
                          tolower(navn) %in% c("karen o'brien", 
                                               "karen o' brien",
                                               "karen linda o' brien") ~ "Karen Linda O'Brien", 
                          
                          TRUE ~ navn))

saldo_faste %>%
  filter(str_detect(navn, "Andrea"))
# 
# 
# navn_f <- saldo_faste %>% 
#   pull(navn) 
# 
# navn_s <- stab %>% 
#   pull(navn)
# 
# navn_f[!(navn_f %in% navn_s)]


# midlertidige ansatte ####
baneM <- "N:/iss-admfelles/Timeregnskap/2023 Høst/PDF MIDLERTIDIG"
filerM <- list.files(baneM)


filerM <- tibble(filerM) %>% 
  filter(str_detect(filerM, "VÅR|HØST")) %>% 
  mutate(sort = ifelse(str_detect(filerM, "oppdatert"), 1, 9),
         id = str_sub(filerM,1,12), 
         aar = parse_number2(filerM),
         semester = ifelse(str_detect(filerM, "HØST"), 2, 1)) %>% 
  arrange(sort) %>% 
  group_by(id) %>%
  slice(1) %>% 
  pull(filerM)


fM <- paste0(baneM, "/", filerM)

fil_oppsM <- tibble(filer = filerM) %>% 
  mutate(total = 0, navn = NA_character_)
# Loop gjennom alle pdf'er for faste, plukk ut totaltall
j = 0
#i <- fM[8]
for(i in fM){
  if(i == fM[1]){ 
    j = 1} else{j = j+1}
  text <- pdf_text(i)
  raw <- map(text, ~ str_split(.x, "\\n") %>% unlist()) %>% 
    reduce(., c)
  totalrader <- stringr::str_which(tolower(raw), "aktiviteter")
  fil_oppsM$total[j] <- parse_number(raw[totalrader][1], 
                                     locale = locale(decimal_mark = ",", 
                                                     grouping_mark = ".")  )
  fil_oppsM$navn[j] <- raw[8] %>% str_squish()
}

midl_aktivitet <- fil_oppsM %>% 
  mutate(aar = parse_number2(filer),
         semester = ifelse(str_detect(filer, "HØST"), "H", "V")) %>% 
  #mutate(navn = paste(fornavn, etternavn) %>% str_squish() ) %>% 
  arrange(navn, desc(aar), semester) %>%  #sortering gir riktig semester ut ved slice
  group_by(navn) %>% 
  slice(1) %>% 
  mutate(oppdatert = paste(aar, semester)) %>% 
  select(navn, total, oppdatert)

oppd <- unique(midl_aktivitet$oppdatert)[1]

# Forrige oppsummering: 
saldo_midl <- read_excel("data/saldo_V23.xlsx", sheet = "midlertidige") %>% 
  mutate(aar = oppd, 
         emne = "Faste poster", 
         aktivitet = paste("SALDO_GJENSTÅENDE")) %>% 
  select(navn, aar, timer, emne, aktivitet) %>% 
  mutate(navn = case_when(navn == "Beatrice Johannessen" ~ "Beatrice Irene Johannessen",
                          navn == "Ludvig Sunnemark" ~ "Erik Ludvig Sunnemark",
                          TRUE ~ navn )) %>%
  right_join(midl_aktivitet, by = c("navn")) %>% 
  mutate(timer = timer + total) %>% 
  select(-total, - oppdatert) 


# Samler for både faste og midlertidige #### 
saldo <- bind_rows(saldo_faste, saldo_midl) %>% 
  select(navn, timer, aar, emne, aktivitet) %>% 
  mutate(semester = ifelse(str_sub(aar, -1, -1) == "V", 1, 2),
         aar = str_sub(aar, 1, 4)) %>%
  mutate(navn = str_to_title(navn)) %>% 
  left_join(stab, by = "navn") 

save(saldo, file = "data/saldo.Rdat")
