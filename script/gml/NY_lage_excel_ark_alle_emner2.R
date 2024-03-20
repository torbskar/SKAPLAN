#-------------------------------;  
# Lage excel-ark 
#-------------------------------;  


library(tidyverse)
library(lubridate)
library(openxlsx)

invisible(Sys.setlocale(locale='no_NB.utf8'))
#year(Sys.Date())





# Fra fil med oversikt over emneansvarlige 
alleemner <- readxl::read_excel("Undervisningsplan_Sosiologi_kopi.xlsx") %>% 
  rename_with(tolower) %>% 
  setNames( c(names(.)[1:5], paste0('aar_', names(.)[6:ncol(.)]) ))

# Rydder litt opp. beholder bare emner som fremdeles går
emneplan <- alleemner %>% 
  select(1:5, num_range("aar_", year(Sys.Date()):2028)) %>%
  filter(is.na(avsluttet)) %>% 
  mutate(emnekode = as.character(map(strsplit(emne, split = " "), 1))) %>% 
  mutate(emnekode = str_remove(emnekode, ",")) %>% 
  group_by(emnekode) %>% 
  slice(1) %>% 
  ungroup()


## Eksempel: Henter for ISS ####
inst <- "700"
sem <- "23v"
# https://tp.uio.no/report/semreport.php?inst=185170700&semester=21v&type=staff
filnavn <- paste0("https://tp.uio.no/report/semreport.php?inst=185170", inst, "&semester=", sem, "&type=staff")

require(R.utils)
df <- downloadFile("https://tp.uio.no/report/semreport.php?inst=185170700&semester=20h&type=staff", username = "torbskar@uio.no", password ="!6Arranmalt17!U", method = "wget")

library(RCurl)
??GET
GET(filnavn, authenticate("torbskar", "!6Arranmalt17!U"), content_type("text/csv"), write_disk("temp"))
readLines("temp")


dt <- read.csv2(filnavn)
View(dt)



# Henter fra web: https://www.uio.no/studier/emner/sv/iss/ og her: https://www.uio.no/studier/emner/sv/sv/ 
fleremner <- read.delim("data/emner_fraWeb_H.txt") %>% 
  separate_wider_delim(Emne, names = c("emnekode", "tittel"), delim = "–") %>% 
  mutate(semester = "H") %>%
  bind_rows( read.delim("data/emner_fraWeb_V.txt") %>% 
               separate_wider_delim(Emne, names = c("emnekode", "tittel"), delim = "–") %>% 
               mutate(semester = "V") ) %>%
  mutate(emnekode = str_squish(emnekode), 
         tittel = str_squish(tittel)) %>% 
  select(emnekode, tittel, semester) 


# Slår filene sammen
emneplan <- emneplan %>% 
  select(emnekode, semester) %>% 
  bind_rows(fleremner) %>% 
  mutate(semester = str_sub(semester, 1,1)) %>% 
  bind_rows(alleemner) %>% 
  group_by(emnekode) %>%
  slice(1) %>% 
  ungroup() %>%
  filter(!is.na(emnekode)) 

#emner_kode <- c(emner_kode, fleremner) %>% unique()


#emner <- paste0(emneplan$emnekode, "_", str_sub(emneplan$semester, 1,1) ) 
emner_kode <- emneplan$emnekode
semester <- str_sub(emneplan$semester, 1,1)
length(emner_kode) == length(semester)

View(emneplan)

# Oppsett for hver side i excel-arkene: variablenavn
years <- seq(year(Sys.Date()),year(Sys.Date())+5)
columns <- c("Navn", "Avtalt", "Emneansvar",	"Forelesning",	"Seminargrupper",	"Lage_eksamen",	"Ordinær_sensur",	"Klagesensur",	"Lage_utsatt_eksamen",	"Sensur_utsatt",	"Annet",	"Kommentar")

d <- data.frame((matrix(ncol = length(columns), nrow = 1)))
names(d) <- columns


stab0 <- readxl::read_excel("stab.xlsx") %>% 
  pull(navn)

dataValidation <- c("Avtalt", "Avtalt med forbehold", "Ikke avtalt")

# Instruksjon på første linjer
instruks <- 
  c("Instruksjoner: ", 
    "I første kolonner oppgis undervisers fulle navn. Bruk drop-down meny.", 
    "Hvis det mangler navn i menyen, velg 'Annen' og skriv i kommentarfeltet i siste kollonne.",
    "Antall oppgaver av hver type angis for hver person i etterfølgende kollonner. F.eks. en vanlig dobbeltforelesning angis som 1.", 
    "Du kan oppgi brøker. F.eks. ved delt emneansvar angir du 0,5 i stedet for 1.",
    "Se kommentarfelt i hver kollonne for mer informasjon og eksempler."
  )


celleinstruks <- c("Fullt navn må staves nøyaktig. Bruk drop-down listen. 
Hvis du åpner filen i webapplikasjonen vil autofill fungere. 
MERK! Hvis du ikke finner navnet i listen, velg 'Annen' og skriv navnet i kommentarfeltet i siste kollonne.", 
                   "Er oppgaven avtalt med utførende person?", 
                   "For emneansvar alene skriv 1, hvis delt ansvar skriv 0,5",
                   "Skriv antall dobbeltforelesninger. 1 forelesning på 2x45 minutter er da 1",
                   "Skriv antall seminarganger. F.eks. Hvis personen har to seminargrupper som møtes 7 ganger à 2 x 45 minutter, skriv: 14",
                   "Den som lager eksamen skriver 1 her. (Normalt emneansvarlig). Kan skrive brøk hvis flere lager eksamen sammen.",
                   "Antall oppgaver å sensurere, uansett uttelling",
                   "Som ordinær sensur",
                   "Som å lage ordinær eksamen",
                   "Som ordinær sensur",
                   "Antall avtalte timer ekstra. Legg til kommentar.",
                   "Kommenter ved behov. Må fylles ut hvis navn på person mangler i listen og hvis annet er fylt ut.") %>% 
  data.frame() %>% 
  mutate(id = row_number()) %>% 
  pivot_wider(values_from = 1, names_from = id)

# Stil på instruksjon ####

# Overskrift instruksjon
hs0b <- createStyle(fgFill = "#ffe6b3", textDecoration = "Bold") 
# Generell instruksjon
hs0 <- createStyle(fgFill = "#ffe6b3", textDecoration = "italic", wrapText = TRUE)
# Celleinstruksjon
hs0c <- createStyle(fgFill = "#eeeeee", textDecoration = "italic", wrapText = TRUE)

# Stil på header-raden i tabellen
hs1 <- createStyle(fgFill = "#4F81BD", halign = "CENTER", textDecoration = "Bold",
                   border = "Bottom", fontColour = "white")


# KJØRER GJENNOM ####

for(i in 1:length(emner_kode)){
  #print(i)
  
  wb <- createWorkbook()
  addWorksheet(wb, "stab", visible = FALSE)  # skjult ark med stab til drop-down meny
  addWorksheet(wb, "dataValidation", visible = FALSE)  # skjult ark med stab til drop-down meny
  writeData(wb, "stab", x = stab)
  writeData(wb, "dataValidation", x = dataValidation)
  
  filename <- paste0("emner/", emner_kode[i], ".xlsx")
  
  if(file.exists(filename)){
    print("Filen finnes fra før")
  } else{ # Filer som IKKE er opprettet ennå
   for(j in 1:length(years)){
    
      sheetname <- paste(years[j], semester[i])
      addWorksheet(wb, sheetname)
      
      setColWidths(wb, sheet = sheetname, cols = c(1, 3:11), widths = 20)
      setColWidths(wb, sheet = sheetname, cols = 2, widths = 15)
      setColWidths(wb, sheet = sheetname, cols = 12, widths = 50)
      
      for(k in 1:length(instruks)){
        mergeCells(wb, sheetname, 1:12, k)
      }
      
      ## Generell instruks her 
      writeData(wb, sheetname, startRow = 1,
                x = instruks)
      addStyle(wb, sheetname, hs0b, 1, 1)
      addStyle(wb, sheetname, hs0, 2:length(instruks), 1)
      
      ## Celleinstruks her (legg til tom linje)
      writeData(wb, sheetname, startRow = length(instruks)+2, 
                x = celleinstruks, colNames = FALSE)
      setRowHeights(wb, sheetname, length(instruks)+2, 190)
      addStyle(wb, sheetname, hs0c, length(instruks)+2, 1:12)
      
      ## Kollonner til data her
      writeData(wb, sheetname, x = d, 
                startRow = length(instruks)+3, startCol = 1, 
                colNames = TRUE, headerStyle = hs1)
      #addStyle(wb, sheetname, style = hs1, rows = length(instruks)+3, cols = 1:12, gridExpand = TRUE)
      #setColWidths(wb, sheet = sheetname, cols = 1:12, widths = 20)
      
      # data validation
      data_validation_source <- "'stab'!$A$1:$A$1000"
      data_validation_source2 <- "'dataValidation'!$A$1:$A$4" 
      
        dataValidation(wb, sheetname, 
                       col = 1, rows = length(instruks)+4:50, 
                       type = "list", 
                       value = data_validation_source)
        dataValidation(wb, sheetname, 
                       col = 2, rows = length(instruks)+4:50, 
                       type = "list", 
                       value = data_validation_source2)
        dataValidation(wb, sheetname,
                       col = 3:10, rows = length(instruks)+4:50, type = "decimal",
                       operator = "between", value = c(0, 60))
        dataValidation(wb, sheetname,
                       col = 11, rows = length(instruks)+4:50, type = "decimal",
                       operator = "between", value = c(0, 99))

        
        for(rad in length(instruks)+4:50){ 
          # Når oppgaver er fylt ut, men ikke person
          regel <- paste0('AND( SUM(B',rad,':I',rad, ') > 0, LEN(A', rad, ') = 0)' )
          conditionalFormatting(wb, sheetname, 
                                cols = 11, rows = rad, 
                                rule = regel, 
                                style = createStyle(bgFill = "#FFC7CE"))
        }
        
        # Når antall timer er veldig høyt
        regel_h <- paste0('AND( SUM(B',rad,':J',rad, ') > 0, LEN(A', rad, ') = 0)' )
        conditionalFormatting(wb, sheetname, 
                              cols = 11, rows = rad, 
                              rule = regel, 
                              style = createStyle(bgFill = "#FFC7CE"))
        
                

        

      # Lås ark   
      protectWorksheet(wb, sheetname, protect = TRUE, 
                       lockFormattingCells = FALSE, lockFormattingColumns = FALSE, 
                       lockInsertingColumns = TRUE, lockDeletingColumns = TRUE)
        
        #This line allows specified cells to be unlocked so that users can enter values.
      for(l in 1:12){
        addStyle(wb, sheetname, 
                 style = createStyle(locked = FALSE), 
                 rows = length(instruks)+4:50, cols = l)
        }
    
   }
    # Lagrer til fil
    sheetVisibility(wb)[1:2] <- "veryHidden"
    saveWorkbook(wb, filename, overwrite = TRUE)
  }

} 


