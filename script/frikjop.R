# Frikjøp 

invisible(Sys.setlocale(locale='no_NB.utf8'))  # Sikrer at æøå etc blir riktig. Kan avhenge av oppsett på lokal maskin 


# I frikjøpsfila er det bare skrevet etternavn. 
# Derfor må vi hente etternavn fra stab og koble på fullt navn
if(file.exists("data/stab.Rdat") & 
   difftime(Sys.time(), 
            file.info("data/stab.Rdat")$mtime,
            units = "weeks") |> as.numeric() < 8 ){
  load(file = "data/stab.Rdat")
} else{
  source("script/stab_andreFolk.R", local = TRUE)
}

stabsnavn <- stab %>% 
  mutate(etternavn = word(navn, -1)) %>% 
  select(navn, etternavn) %>% 
  group_by(etternavn) %>%
  slice(1) %>%
  ungroup() 






# Mappe i administrasjonsområde der oversikt over frikjøp ligger
mappe <- "N:/iss-forskningsadministrasjon/7. Til ledelsen/Frikjøp"

filer <- list.files(mappe) 
frikjopfil <- filer[str_detect(filer, "Frikjøp")]



navn <- read_excel(paste0(mappe, "/", frikjopfil), 
                  sheet = 1,
                  skip = 3,
                  col_names = FALSE)[1:2,] 

navn[1,] <- navn[1,]  %>% 
  pivot_longer(cols = -1, values_to = "aar", names_to = "typ") %>% 
  fill(aar) %>% 
  pivot_wider(names_from = typ, values_from = aar)


nyenavn <- paste(navn[1,], navn[2,]) %>% 
  str_replace(., " ", "_") %>% 
  str_replace_all("_NA", "") %>% 
  tolower()

fri <- read_excel(paste0(mappe, "/", frikjopfil), 
                  sheet = 1,
                  skip = 5,
                  col_names = c(nyenavn)) 


frikjop <- fri %>% 
  fill(navn) %>% 
  filter(!is.na(navn)) %>%
  filter(!is.na(prosjektno.)) %>%
  select(navn, contains("ekstern")) %>%
  pivot_longer(cols = -navn, values_to = c("pct"), names_to = "intekst") %>% 
  mutate(aar = as.numeric(str_sub(intekst, 1,4)),
         type = str_sub(intekst, 6,14)) %>% 
  select(-intekst) %>% 
  filter(!is.na(pct)) %>% 
  group_by(navn, aar) %>% 
  reframe(frikjop = sum(pct)/100) %>% 
  filter(!is.na(frikjop))


frikjop <- frikjop %>% 
  arrange(navn, aar) %>%
  group_by(navn, aar) %>% 
  expand(nesting(semester = c("V", "H"))) %>% 
  left_join(frikjop, by = c("navn", "aar")) %>% 
  mutate(timer = frikjop*898*.53,   ## Totalt timer og andel arbeidsplikt
         emne = "Faste poster", 
         aktivitet = "Frikjøp", 
         aar = paste(aar, semester)) %>% 
  select(-semester) %>% 
  rename(etternavn = navn) %>% 
  left_join(stabsnavn, by = "etternavn") %>% 
  select(-etternavn)



# frikjop %>%
#   filter(str_detect(navn, "Tork"))
# 
# View(frikjop)
