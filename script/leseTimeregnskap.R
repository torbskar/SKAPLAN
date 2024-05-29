#################################;
# Lese inn fra timeregnskap  
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

library(tidyverse)

# Henter navn på alle faste ansatte ####
source("script/stab_andreFolk.R")

# Alle fulle navn for senere kobling. Vær obs på dubletter på etternavn! 
fulltnavn_faste <- stab %>% 
  filter(Stillingsgruppe %in% c("1011 Førsteamanuensis", 
                                "1013 Professor", "1198 Førstelektor")) %>% 
  select(navn) %>%
  mutate(etternavn = str_extract(navn, "[\\w-]+\\b$")) %>%  # siste ord i string, inkluder bindestrek 
  mutate(etternavn = case_when(etternavn == "Brien" ~ "O'Brien",
                               etternavn == "Naveda" ~ "Garcia-Godos",
                               TRUE ~ etternavn)) %>% 
  group_by(etternavn) %>% 
  slice(1) %>%              # OBS! Flere med samme etternavn er trøbbel her. Bør fikses i saldo-filen. Gjelder foreløpig kun Bjørn og Synøve
  ungroup()


#View(fulltnavn_faste)

### Rett stavemåte for navn ####
rettnavn <- read_excel("data/navn_variasjoner.xlsx") %>% 
  mutate(navn_low = str_to_lower(navn), 
         riktig_navn = str_squish(riktig_navn)) 

# Hente ut saldo fra forrige semester for midlertidige ####
bane_timeregnskap <-  "N:/iss-admfelles/Timeregnskap/2023 Høst/"
bane_timeregnskap_forrige <-  "N:/iss-admfelles/Timeregnskap/2023 Vår/"

# Sjekker om det finnes en fil med navn som inneholder "saldo" i filnavnet
# Dersom det finnes, les inn saldo fra forrige semester
# Hvis ikke, kjør scraping av alle pdf-filer for faste og midlertidige

# OBS! I denne filen har faste ansatte kun oppgitt etternavn. Separat fane for SOS og HGO
#      Midlertidige har fullt navn i to kolonner. 

if( any(grepl("saldo",list.files(bane_timeregnskap)) ) ){
  
  saldofil <- list.files(bane_timeregnskap) %>% 
    str_subset("saldo") %>% 
    str_subset("~", negate = TRUE) %>% 
    paste0(bane_timeregnskap, .)

  sheet_names <- excel_sheets(saldofil)[1:4]
  
  ### Faste ansatte ####
  saldo_faste <- lapply(sheet_names[1:2], function(x) {          # Read all sheets to list
    as.data.frame(read_excel(saldofil, 
                             sheet = x,
                             skip = 1, 
                             col_types = c("text") )) 
  } 
  ) %>% 
    bind_rows() %>% 
    filter( !grepl("[[:alpha:]]", .[,2])  ) %>% 
    filter(!is.na(.[,2])) %>% 
    mutate(semester = str_sub( colnames(.)[2], 1, 1),
           aar0 = parse_number(colnames(.)[2]),
           timer = as.numeric(.[,2])) %>%
    mutate(emne = "Faste poster", 
           aktivitet = paste("SALDO")) %>% 
    mutate(aar = paste(aar0, semester)) %>% 
    mutate(etternavn = word(Navn), 
           etternavn = str_remove(etternavn, '\\*'),
           etternavn = str_remove(etternavn, ",")) %>%
    left_join(fulltnavn_faste, by = "etternavn") %>%
    mutate(navn = ifelse(is.na(navn), etternavn, navn)) %>% # hvis ikke match, behold etternavn
    select(navn, timer, aar, emne, aktivitet) %>% 
    mutate(navn = ifelse(str_ends(tolower(navn), "brien"),       # OBS! Karen O'Brien gir alltid trøbbel i tekstfunksjoner. Gjør manuell endring.
                                  "Karen Linda O'Brien", navn)) 
# View(saldo_faste)
   
  ### Midlertidige ansatte ####
  saldo_midl <- lapply(sheet_names[3], function(x) {          # Read all sheets to list
    as.data.frame(read_excel(saldofil, 
                             sheet = x,
                             skip = 2, 
                             col_types = c("text") ))
    }) %>% 
      bind_rows() %>% 
      filter( !grepl("[[:alpha:]]", .[,6])  ) %>% 
      filter(!is.na(.[,6])) %>% 
      mutate(semester = str_sub( colnames(.)[6], -3, -3), # OBS! -3 fordi angitt med f.eks. H23
             aar0 = parse_number(colnames(.)[6])+2000,     # OBS! +2000 for å få riktig årstall fra 2-sifret
             timer = as.numeric(.[,6]),
             navn = paste(Fornavn, Etternavn) |> str_squish() ) %>% 
      mutate(aar = paste(aar0, semester),
           emne = "Faste poster", 
           aktivitet = paste("SALDO_GJENSTÅENDE")) %>% 
      select(navn, timer, aar, emne, aktivitet) %>%
      mutate(navn = case_when(navn == "Eirik J Berger" ~ "Eirik Jerven Berger",   # Manuell fix :(
                              navn == "Beatrice Johannessen" ~ "Beatrice Irene Johannessen",
                              navn == "Ludvig Sunnemark" ~ "Erik Ludvig Sunnemark",
                              TRUE ~ navn )) 
  
  # ## Samler for både faste og midlertidige ####
  # saldo <- bind_rows(saldofaste, saldo_midl) %>% 
  #   left_join(., rettnavn, by = "navn") %>% 
  #   mutate(navn = case_when( !is.na(navn) & navn != riktig_navn ~ riktig_navn,
  #                            TRUE ~ navn)) %>% 
  #   select(-riktig_navn) 
  
  #glimpse(saldo)
  

} else {    # Hvis må lese fra pdf-filer ####      

    library(pdftools)  
    
    ## Pdf'er for fast ansatte ####
    baneF <- paste0(bane_timeregnskap, "PDF FAST")  
    filerF <- list.files(baneF)
    filerF <- tibble(filerF)$filerF 
    
    # funksjon for å trekke ut nummer fra filnavn
    parse_number2 <- function(x){
      x %>% 
        str_extract("\\d{4}") %>% 
        as.numeric()
    }
    
    
    ### Velger riktig fil hvis oppdatert ####
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
    
    ### Loop faste, plukk ut totaltall ####
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
      left_join(., rettnavn, by = "navn") %>% 
      mutate(navn = case_when( !is.na(navn) & navn != riktig_navn ~ riktig_navn,
                               TRUE ~ navn)) %>% 
      select(-riktig_navn) 
    
    
    #View(saldo_faste)
    
    # saldo_faste %>%
    #   filter(str_detect(navn, "Magne"))
    # 
    # 
    # navn_f <- saldo_faste %>% 
    #   pull(navn) 
    # 
    # navn_s <- stab %>% 
    #   pull(navn)
    # 
    # navn_f[!(navn_f %in% navn_s)]
    
    
    ### Loop midlertidige, plukk ut aktivitet siste år ####
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
      filter(!is.na(navn)) %>% 
      mutate(aar = parse_number2(filer),
             semester = ifelse(str_detect(filer, "HØST"), "H", "V")) %>% 
      arrange(navn, desc(aar), semester) %>%  #sortering gir riktig semester ut ved slice
      group_by(navn) %>% 
      slice(1) %>% 
      mutate(oppdatert = paste(aar, semester)) %>% 
      select(navn, total, oppdatert) %>% 
      mutate(navn_low = str_to_lower(navn)) %>% 
      left_join(., rettnavn[, c("riktig_navn", "navn_low")], by = "navn_low") %>% 
      mutate(navn = case_when( !is.na(navn) & navn != riktig_navn ~ riktig_navn,
                               TRUE ~ navn)) %>% 
      select(-c(riktig_navn, navn_low)) %>% 
      filter(!is.na(total))
#  View(midl_aktivitet)
      
    oppd <- unique(midl_aktivitet$oppdatert)[1]
    
    ### Finn forrige saldo fra oppsummeringsfil: ####
    saldofil_f <- list.files(bane_timeregnskap_forrige) %>% 
      str_subset("saldo") %>% 
      str_subset("~", negate = TRUE) %>% 
      paste0(bane_timeregnskap_forrige, .)
    
    sheet_names <- excel_sheets(saldofil_f)[1:3]
    

    ### Midlertidige ansatte ####
    saldo_midl <- lapply(sheet_names[3], function(x) {          # Read all sheets to list
      as.data.frame(read_excel(saldofil_f, 
                               sheet = x,
                               skip = 2, 
                               col_types = c("text") ))
    }) %>% 
      bind_rows() %>% 
      filter( !grepl("[[:alpha:]]", .[,6])  ) %>% 
      filter(!is.na(.[,6])) %>% 
      mutate(semester = str_sub( colnames(.)[6], -3, -3), # OBS! -3 fordi angitt med f.eks. H23
             aar0 = parse_number(colnames(.)[6])+2000,     # OBS! +2000 for å få riktig årstall fra 2-sifret
             timer = as.numeric(.[,6]),
             navn = paste(Fornavn, Etternavn) |> str_squish(), 
             navn_low = str_to_lower(navn)) %>% 
      mutate(aar = paste(aar0, semester),
             emne = "Faste poster", 
             aktivitet = paste("SALDO_GJENSTÅENDE")) %>% 
      select(navn, navn_low, timer, aar, emne, aktivitet) %>%
      left_join(rettnavn[, c("navn_low", "riktig_navn")], by = "navn_low") %>%
      mutate(navn = case_when( !is.na(navn) & navn != riktig_navn ~ riktig_navn,
                               TRUE ~ navn)) %>% 
      select(-riktig_navn) %>%   
      full_join(midl_aktivitet, by = c("navn")) %>% 
      mutate(total = replace_na(total, 0), 
             timer = timer + total) %>% 
      select(-total, - oppdatert) 

    
    ## Avslutt ifelse for saldo-fil ####
  } 

#View(saldo_faste)

## Samler for både faste og midlertidige #### 
saldo <- bind_rows(saldo_faste, saldo_midl) %>% 
  select(navn, timer, aar, emne, aktivitet) %>% 
  mutate(semester = ifelse(str_sub(aar, -1, -1) == "V", 1, 2),
         aar = str_sub(aar, 1, 4)) %>%
  mutate(navn = str_to_title(navn),
         navn_low = str_to_lower(navn)) %>% 
  left_join(rettnavn[, c("navn_low", "riktig_navn")], by = "navn_low") %>%
  mutate(navn = case_when( !is.na(navn) & navn != riktig_navn ~ riktig_navn,
                           TRUE ~ navn)) %>% 
  select(-riktig_navn) %>% 
  left_join(stab, by = "navn") 

#View(saldo)

# Lagre saldo-fil####
save(saldo, file = "data/saldo.Rdat")
