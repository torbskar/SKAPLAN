#*****************************************************************
#* Leser inn data fra excel-filer og lager en oversikt over timer *
#* og ressursbruk.                                                *
#* Skrevet av Torbjørn Skardhamar, 2024                           *
#*****************************************************************

rm(list = ls())
invisible(Sys.setlocale(locale='no_NB.utf8'))  # Sikrer at æøå etc blir riktig. Kan avhenge av oppsett på lokal maskin 

library(tidyverse)
library(readxl)
library(openxlsx)


# LESER INN DATA ####

## All ordinær undervisning ####
# Henter fra mappe med excel-filer
source("script/emnefiler.R")


## Veiledning MA #### 
source("script/veiledning_ma.R")

## Ekstra ####
source("script/ekstra.R")

## Uniped ####
source("script/uniped.R")

## Norskkurs ####
source("script/norskkurs.R")


## Utland phd ####
source("script/utenlandsopphold.R")

## Lederverv og programleder ####
source("script/verv.R")

## Frikjøp ####
source("script/frikjop.R")




## Forskningstermin ####
# Leser tidligere lagret fil hvis mindre enn 8 uker gammel
if(file.exists("data/forskningstermin.Rdat") & 
               difftime(Sys.time(), 
                        file.info("data/forskningstermin.Rdat")$mtime,
                        units = "weeks") |> as.numeric() < 8 ){
  load(file = "data/forskningstermin.Rdat")
} else{
  source("script/forskningstermin_planlagt.R", local = TRUE)
}



# setter filene sammen #### 

# Gjeldende satser i timeregnskapet ####
satser <- read_xlsx("data/satser.xlsx") %>% 
  filter(!is.na(uttelling)) %>% 
  select(kode, uttelling) %>% 
  rename(aktivitet = kode) %>% 
  arrange(aktivitet) %>% 
  filter(!is.na(uttelling) & !is.na(aktivitet))

# Emner som har ikke-standard uttelling
emner_stor_sensur <- read_xlsx("data/Emner_ekstra_uttelling.xlsx") %>% 
  filter(!is.na(eksamen)) %>% 
  pull(emne)
emner_stor_eksamen <- read_xlsx("data/Emner_ekstra_uttelling.xlsx") %>% 
  filter(!is.na(lage_eksamen)) %>% 
  pull(emne)

navn_var <- read_xlsx("data/navn_variasjoner.xlsx")

#  Noen navn er forskjellig i Aura og ansattlista, og TP. Kan gi feil i emnearkene 
#  Navn med apostrof (Karen) gir trøbbel pga tekstfunksjoner som brukes div steder

budsjett <- list_all %>% 
  bind_rows() %>% 
  rename(antall = timer ) %>%          # Det er antall aktiviteter som registreres. Må regnes om.  
  bind_rows(., verv) %>% 
  bind_rows(., utlandstimer) %>% 
  bind_rows(., veiledning_ma) %>%
  bind_rows(., ekstratimer) %>%
  bind_rows(., unipedtimer) %>% 
  bind_rows(., frikjop) %>% 
  bind_rows(., norskkurs) %>% 
  mutate(navn = str_to_title(navn)) %>%    # Sikrer store/små bokstaver hvis uhell
  mutate(navn = str_replace(navn, "' ", "'")) %>% 
  left_join(navn_var, by = "navn") %>% 
  mutate(navn = ifelse( !is.na(riktig_navn), riktig_navn, navn )) %>% 
  mutate(antall = case_when(aktivitet == "forelesning" ~ antall*2, 
                            aktivitet == "emneansvar" & antall < 1 ~ antall*1.5, 
                            aktivitet == "seminargrupper" ~ antall*2, 
                            TRUE ~ antall)) %>% 
  mutate(aktivitet_midl = aktivitet) %>%   # Lager kopi av detaljert aktivitet før omkoding og sater
  mutate(aktivitet = case_when(tolower(aktivitet) %in% c("ordinær_sensur", "sensur_utsatt", "klagesensur") ~ "sensur",
                               tolower(aktivitet) %in% ("lage_utsatt_eksamen") ~ "lage_eksamen",
                               TRUE ~ aktivitet), 
         aktivitet = case_when(tolower(aktivitet) == "sensur" & 
                                 emne %in% emner_stor_sensur ~ "sensur_stor",
                               TRUE ~ aktivitet),
         aktivitet = case_when(tolower(aktivitet) == "lage_eksamen" & 
                                 emne %in% emner_stor_eksamen ~ "lage_eksamen2",
                               TRUE ~ aktivitet)) %>%
  left_join(satser, by = "aktivitet") %>% 
  mutate(aktivitet = aktivitet_midl,
         uttelling = replace_na(uttelling, 1), 
         timer = ifelse(is.na(timer), antall * uttelling, timer)) %>%           # Regner om til timer der antall er angitt. Ellers beholdes timer angitt. 
  mutate(aktivitet_grov = case_when(str_detect(aktivitet, "sensur") ~ "sensur",
                                    str_detect(aktivitet, "eksamen") ~ "lage eksamen", 
                                    tolower(aktivitet) %in% c("emneansvar", "forelesning", "seminargrupper") ~ "undervisnings") 
         )%>% 
  select(-aktivitet_midl) # drop midlertidig


# unique(budsjett$aktivitet)
# 
budsjett %>%
  filter( str_detect(navn, "Tork")) %>%
  arrange(aktivitet) %>%
  select(navn, aar, emne, aktivitet, antall, uttelling, timer) %>%
  View()


## Oversikt over person #### 
# Hvis eldre enn 8 uker, kjør på nytt
if(file.exists("data/stab.Rdat") & 
      difftime(Sys.time(), 
               file.info("data/stab.Rdat")$mtime,
               units = "weeks") |> as.numeric() < 8 ){
  load(file = "data/stab.Rdat")
} else{
  source("script/stab_andreFolk.R", local = TRUE)
}


## Faste tillegg ####
# 
iaar <- year(Sys.Date())  # for å avgrense antall år fremover

fastetillegg <- stab %>% 
  mutate(sluttdato = ifelse(is.na(sluttdato), as_date(ymd("2099-01-01")), sluttdato) |> as_date()) %>% 
  group_by(navn, Stillingsgruppe, sluttdato) %>% 
  expand(semester = c(1,2), 
         aar = iaar:(iaar+4)) %>%  
  arrange(navn, aar, semester) %>% 
  mutate(sluttsemester = ifelse(month(sluttdato) <= 7, 1, 2 )) %>% 
  filter(aar < year(sluttdato) | (aar == year(sluttdato) & semester <= sluttsemester ) ) %>% 
  mutate(Lopendefaglig = 30, 
         kontakttid = 20) %>% 
  pivot_longer(cols = c("Lopendefaglig", "kontakttid"),  
               values_to = "timer", 
               names_to = "aktivitet") %>% 
  ungroup() %>% 
  mutate(aar = paste(aar), 
         emne = "Faste poster") %>% 
  select(navn, aar, semester, emne, aktivitet, timer) %>% 
  left_join(stab, by = "navn")


fastetillegg %>%
  filter( str_detect(navn, "Martin") ) %>% 
  # filter(navn == "Agnes Fauske") %>%
  View()



# Saldo fra timeregnskapet #### 
# Sjekk om fil er lagret til disk først. Hvis ikke: kjør script på nytt. 
#   Poenget er å spare litt tid. 
if(file.exists("data/saldo.Rdat") & 
   difftime(Sys.time(), 
            file.info("data/stab.Rdat")$mtime,
            units = "weeks") |> as.numeric() < 8 ){
  load(file = "data/saldo.Rdat")
} else{
  source("script/leseTimeregnskap_fraPDF.R")
}




#allenavn <- c(stab$navn, budsjett$navn) %>% unique() %>% sort()


# saldo_midl %>%
#   filter(str_detect(navn, "Frith"))

# Plasserer faste poster og saldo øverst ved sortering  ####
fastsaldo <- bind_rows(fastetillegg, saldo) %>% 
  mutate(sort = str_sub(aktivitet, 1, 5) %in% c("SALDO", "SALDO_GJENSTÅENDE")+0) %>% 
  arrange(navn, aar, semester, desc(sort)) %>% 
  group_by(navn, aar, semester) %>% 
  filter( (first(sort) == 1 & sort == 1) | first(sort) != 1) %>%
  ungroup %>% 
  select(-sort) 

# Legger til ansattgruppe og sluttdato #### 
#  Fjerner underscore fra aktivitet-streng
#  Gir 15% tillegg til stipendiater for undervisningsaktiviteter
budsjett <- budsjett %>% 
  left_join(stab[, c("navn", "Stillingsgruppe", "sluttdato")], by = "navn") %>% 
  mutate(timer = case_when(Stillingsgruppe == "1017 Stipendiat" & 
                             (aktivitet %in% c("forelesning", 
                                               "seminargrupper", 
                                               "ordinær_sensur",
                                               "emneansvar", 
                                               "lage_eksamen", 
                                               "lage_utsatt_eksamen", 
                                               "sensur_utsatt", 
                                               "veiledning_ma") ) ~ timer*1.15,
                           TRUE ~ timer)) %>% 
  mutate(semester = ifelse(str_sub(aar, -1, -1) == "V", 1, 2),
         aar = str_sub(aar, 1, 4)) %>% 
  filter(aar > 2023 | (aar == 2023 & semester == 2))  %>%    ## Velger startår/semester 
  bind_rows(., fastsaldo) %>% 
  mutate(aktivitet = str_replace_all(aktivitet, "_", " ") |> str_to_title()) 
  
# budsjett %>%
#   filter(str_detect(navn, "Lars Erik ")) %>% 
#   View()


#  budsjett %>%
# #   #mutate(aktivitet = str_replace_all(aktivitet, "_", " ") |> str_to_title()) %>%
#    filter(str_detect(navn, "Andrea ")) %>%
# # #   filter(navn == "Alf Jørgen Schnell") %>%
# #    View()
#    pull(navn) %>% unique() ==
 # saldo_faste %>%
 #   filter(str_detect(navn, "Andrea ")) %>% 
 #   pull(navn)
 
## Lagrer permanent fil ####
budsjett %>% 
  write_rds("data/timebudsjett.rds")


# budsjett <- read_rds("data/timebudsjett.rds")
# 
# budsjett %>%
#   filter(str_detect(navn, "Martin")) %>%
#   select(aar, aktivitet, antall, timer, uttelling) %>%
#   View()


# Oversiktsrapport #### 
#  Lagrer i overordnet mappe først - deretter flytte. Pga teknisk problem som sikkert kan løses. 
quarto::quarto_render(input = "personoversikt.qmd", 
                      output_file = "personoversikt.pdf") 

file.copy(from = paste0(here::here(), "/", "personoversikt.pdf"), 
          to = paste0(here::here(), "/emner/", "personoversikt.pdf"),
          overwrite = TRUE)
file.remove(paste0(here::here(), "/", "personoversikt.pdf"))




