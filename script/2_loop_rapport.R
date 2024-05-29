#------------------------------------------------------;
# Funktion: Loop for å kjøre rapporter for personer og emner
# Inndata : timebudsjett.rds
#           stab.Rdat
#           
# Quarto  : personrapport.qmd
#           emnerapport.qmd
#           stab_andreFolk.R
# 
# Ut      : En pdf-fil for hver person i staben
#           En pdf-fil for hvert emne i porteføljen
#------------------------------------------------------;


invisible(Sys.setlocale(locale='no_NB.utf8'))

library(tidyverse)
library(readxl)
library(stringr)

rm(list=(ls()))


#source("script/1_leseInnData.R")

# Log fil mm. ####
sink("log.Rtxt", append = TRUE)
  print( paste("Sist rapport kjørt:", Sys.time()))
sink()

## Siste år/semester for timeregnskapet #### 
sisteaar <- 2023
sistesemester <- 2


## Create file structure 
# if(dir.exists("personrapport/")){ 
#   files2zip <- dir('personrapport', full.names = TRUE)
#   zip(zipfile = paste0('arkiv/personrapport_', Sys.Date()), files = files2zip)
#   unlink("personrapport/", recursive = TRUE)
#   dir.create("personrapport")
# } else {
#   dir.create("personrapport")
# }


# Lese inn data og rydde ####

budsjett <- readRDS("data/timebudsjett.rds") %>% 
  filter(str_detect(aktivitet, "Saldo") | 
           (aar > sisteaar | 
              (aar == sisteaar & semester > sistesemester) )) %>% 
  mutate(aktivitet = ifelse(str_sub(aktivitet,1,5) == "Saldo", "Saldo", aktivitet)) %>% 
  mutate(aktivitet = factor(aktivitet)) %>% 
  mutate(aktivitet = fct_relevel(aktivitet, 
                                 "Saldo", "Lopendefaglig", "Kontakttid",
                                 "Emneansvar", "Forelesning", 
                                 "Lage Eksamen", "Lage Utsatt Eksamen", 
                                 "Seminargrupper", 
                                 "Ordinær Sensur", "Sensur Utsatt", "Klagesensur",
                                 "Annet")) %>% 
  mutate(aktivitet = fct_relevel(aktivitet, "Annet", after = Inf)) %>% 
  arrange(navn, desc(Stillingsgruppe)) %>% 
  group_by(navn) %>% 
  fill(Stillingsgruppe)


# budsjett %>% 
#   filter(emne == "SGO1910") %>% 
#   head()

budsjett %>%
  filter(str_detect(navn, "Frøja")) %>%
  #filter(navn == person  ) %>%
  #select(navn, emne, aar, semester, aktivitet, timer, Stillingsgruppe) %>%
  select(navn, emne, aar, semester, Stillingsgruppe) %>%
  head()
# 
# 
# budsjett %>%
#   filter(str_detect(navn, "Karen")) %>%
#   arrange(aar, semester, emne, aktivitet) %>%
#   #filter(str_detect(aktivitet, "Saldo")) %>%
#   View()

# budsjett %>% 
#   filter(emne == "SOSGEO1120", aar == 2023) %>% 
#   select(navn, emne, aar, semester, aktivitet)


## Personer i budsjett ####
personer <- budsjett %>% 
  filter(navn != "Annen") %>% 
  group_by(navn) %>% 
  slice(1) %>% 
  pull(navn) %>% 
  str_squish()

# Hvem som ikke er i stab
# personer[!(personer %in% stab$navn) ]

## Oppdatere alle personer og finn stab #### 
# Hvis eldre enn 8 uker, kjør på nytt
if(file.exists("data/stab.Rdat") & 
   difftime(Sys.time(), 
            file.info("data/stab.Rdat")$mtime,
            units = "weeks") |> as.numeric() < 8 ){
  load(file = "data/stab.Rdat")
} else{
  source("script/stab_andreFolk.R", local = TRUE)
}

#View(stab)

# Kjører rapporter per person ####
stab_navn <- stab %>% 
  pull(navn)

# str_equal(stab_navn[str_detect(stab_navn, "Arno")]  |> str_to_title(),
#           personer[str_detect(personer, "Arno")] |> str_to_title(),
#           ignore_case = FALSE)

personer_stab <- personer[( tolower(personer) %in% tolower(stab_navn)) ] %>% unique()
#personer_stab <- personer_stab[54:length(personer_stab)]
#personer_stab <- personer_stab[str_detect( personer_stab, "Saga")]
#person <- personer_stab

pathdir <- paste0(getwd(), "/personrapport/")
for(person in personer_stab){
  print(person)
  
  persondata <- budsjett %>% 
    filter(navn == person)
  
  saveRDS(persondata, file= "timebudsjett_person.rds")
  
  ## FUll report
  person <- gsub("[[:punct:]]", " ", person)  # Fjerner punktum i navn
  filnavn <- paste(person, ".pdf", sep='')
  quarto::quarto_render(input = "personrapport.qmd", 
                    output_file = filnavn) 

  file.copy(from = paste0(here::here(), "/", filnavn), 
            to = paste0(here::here(), "/personrapport/", filnavn),
            overwrite = TRUE)
  file.remove(paste0(here::here(), "/", filnavn))
  
  #rm(list=ls()[!(ls() %in% c("person", "personer", "pathrmd", "pathdir", "wd", "dato"))])
  #file.remove("emne.RDat")
}


## Fjerner midlertidige filer ####
rmfiler <- list.files()
rmfiler <- rmfiler[endsWith(rmfiler, ".log") | 
                     endsWith(rmfiler, ".tex") | 
                     endsWith(rmfiler, ".aux")]
if(!is_empty(rmfiler)){
  file.remove(paste0(here::here(), "/", rmfiler))}

file.remove(paste0(here::here(), "/", "timebudsjett_person.rds"))  # emnefil til rapportering

gc()


# Kjører oversiktsrapport for personer #### 
#  Lagrer i overordnet mappe først - deretter flytte. Pga teknisk problem som sikkert kan løses. 
quarto::quarto_render(input = "personoversikt.qmd", 
                      output_file = "personoversikt.pdf") 

file.copy(from = paste0(here::here(), "/", "personoversikt.pdf"), 
          to = paste0(here::here(), "/oversiktsrapporter/", "personoversikt.pdf"),
          overwrite = TRUE)
file.remove(paste0(here::here(), "/", "personoversikt.pdf"))




# Kjører rapporter per emne ####
pathdir <- paste0(getwd(), "/emnerapport/")
emner <- budsjett %>% 
  filter(str_sub(emne,1,3) %in% c("SOS", "KUL", "UTV", "SVL", "SVM", "HGO", "SGO", "OLA")) %>% 
  pull(emne) %>% 
  unique() 

#emner <- "SOS1004"

for(emnet in emner){
  print(emnet)
  
  emnedata <- budsjett %>% 
    #filter(str_sub(emne,1,3) %in% c("SOS", "KUL", "UTV", "SVL", "SVM", "HGO", "SGO", "OLA")) %>% 
    filter(emne == emnet) %>% 
    group_by(navn, emne, aktivitet, aar, semester) %>% 
    summarise(timer = sum(timer), antall = sum(antall)) %>% 
    ungroup()
  
  saveRDS(emnedata, file= "timebudsjett_emne.rds")
  
  ## FUll report
  filnavn <- paste(emnet, ".pdf", sep='')
  quarto::quarto_render(input = "emnerapport.qmd", 
                        output_file = filnavn) 
  
  file.copy(from = paste0(here::here(), "/", filnavn), 
            to = paste0(here::here(), "/emnerapport/", filnavn),
            overwrite = TRUE)
  file.remove(paste0(here::here(), "/", filnavn))
  
}



## Fjerner midlertidige filer ####
rmfiler <- list.files()
rmfiler <- rmfiler[endsWith(rmfiler, ".log") | 
                     endsWith(rmfiler, ".tex") | 
                     endsWith(rmfiler, ".aux")]
if(!is_empty(rmfiler)){
  file.remove(paste0(here::here(), "/", rmfiler))}

file.remove(paste0(here::here(), "/", "timebudsjett_emne.rds"))  # emnefil til rapportering




# Kjører rapport over emneansvarlige per emne

budsjett <- readRDS("data/timebudsjett.rds")
# budsjett <- readRDS("data/timebudsjett.rds") %>% 
#   filter(avtalt == "Avtalt")


emneansvarlige <- budsjett %>% 
  filter( tolower(aktivitet) == "emneansvar") %>% 
  select(aar, semester, emne, navn) %>% 
  # kilde: https://stackoverflow.com/questions/52120034/extract-first-letter-in-each-word-in-r
  mutate(initialer = gsub('\\b(\\pL)\\pL{2,}|.','\\U\\1', navn, perl = TRUE)) %>%  
  mutate(navn = paste( str_sub(initialer, 1, -2), word(navn, -1))) %>%
  select(-initialer) %>% 
  group_by(aar, semester, emne) %>% 
  summarise(navn = paste(navn, collapse = " & "))  %>% 
  arrange(aar, desc(semester), emne) %>%
  pivot_wider(names_from = c(aar, semester), values_from = navn, names_sort = TRUE) %>% 
  ungroup() %>% 
  mutate(program = gsub("\\d", "", emne),
         program = case_when(parse_number(emne) >= 9000  ~ "phd_program", 
                             program %in% c("SGO", "HGO") | 
                               emne %in% c("SOSGEO2040") ~ "Samfunnsgeografi",
                             program %in% c("SOS") | 
                               emne %in% c("SVMET1010") ~ "Sosiologi",
                             program %in% c("OLA", "SVPRO") ~ "Organisasjon, ledelse og arbeid",
                             program %in% c("KULKOM") ~ "Kultur og kommunikasjon",
                             program %in% c("SVLEP") ~ "Lektorprogrammet",
                             program %in% c("UTV") ~ "Utviklingsstudier og bærekraft", 
                             TRUE ~ "Sosiologi")) %>%
  mutate(program = case_when(parse_number(emne) >= 4000 & parse_number(emne) < 5000 ~ paste(program, "MA"),
                             parse_number(emne) < 4000 & parse_number(emne) >= 1000 ~ paste(program, "BA"),
                             TRUE ~program)) 


prog <- unique(emneansvarlige$program)
#programx <- "Samfunnsgeografi"

for(programx in prog){
  print(programx)
  
  ansvarligdata <- emneansvarlige %>% 
    filter(program == programx)
  
  saveRDS(ansvarligdata, file= "ansvarligdata.rds")
  
  ## FUll report
  
  filnavn <- paste(programx, ".pdf", sep='')
  quarto::quarto_render(input = "oversikt_emne_ansvarlig.qmd", 
                        output_file = filnavn) 
  
  file.copy(from = paste0(here::here(), "/", filnavn), 
            to = paste0(here::here(), "/oversiktsrapporter/", filnavn),
            overwrite = TRUE)
  file.remove(paste0(here::here(), "/", filnavn))
  
}


