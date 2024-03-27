#-------------------------------;  
# Lage excel-ark 
#-------------------------------;  


library(tidyverse)
library(lubridate)
library(openxlsx)

invisible(Sys.setlocale(locale='no_NB.utf8'))
#year(Sys.Date())


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


emner_kode <- fleremner$emnekode
semester <- fleremner$semester
length(emner_kode) == length(semester)

# Oppsett for hver side i excel-arkene: variablenavn
years <- seq(year(Sys.Date()),year(Sys.Date())+5)
columns <- c("Navn",	"Emneansvar",	"Forelesning",	"Seminargrupper",	"Lage_eksamen",	"Ordinær_sensur",	"Klagesensur",	"Lage_utsatt_eksamen",	"Sensur_utsatt",	"Annet",	"Kommentar")

d <- data.frame((matrix(ncol = length(columns), nrow = 1)))
names(d) <- columns

# Oversikt over stab i en fil levert fra kontorsjef
source("script/stab_andreFolk.R")

# Instruksjon på første linjer
instruks <- 
  c("Instruksjoner: ", 
  "I første kolonner oppgis undervisers fulle navn. Bruk drop-down meny.", 
  "Hvis det mangler navn i menyen, velg 'Annen' og skriv i kommentarfeltet i siste kollonne.",
  "Antall oppgaver av hver type angis for hver person i etterfølgende kollonner. F.eks. en vanlig dobbeltforelesning angis som 1.", 
  "Du kan oppgi brøker. F.eks. ved delt emneansvar angir du 0,5 i stedet for 1.",
  "Se kommentarfelt i hver kollonne for mer informasjon og eksempler."
  )

# OBS! For comments vil linjeskift og tab etc. bli med i kommentaren. 
c1 <- createComment(comment = "Fullt navn må staves nøyaktig. Bruk drop-down listen. 
Hvis du åpner filen i webapplikasjonen vil autofill fungere. 
MERK! Hvis du ikke finner navnet i listen, velg 'Annen' og skriv navnet i kommentarfeltet i siste kollonne.",
visible = FALSE)

c2 <- createComment(comment = "For emneansvar alene skriv 1, hvis delt ansvar skriv 0,5",
                    visible = FALSE)
c3 <- createComment(comment = "Skriv antall dobbeltforelesninger. 1 forelesning på 2x45 minutter er da 1",
                    visible = FALSE)
c4 <- createComment(comment = "Skriv antall seminarganger. F.eks. Hvis personen har to seminargrupper som møtes 7 ganger à 2 x 45 minutter, skriv: 14",
                    visible = FALSE)
c5 <- createComment(comment = "Den som lager eksamen skriver 1 her. (Normalt emneansvarlig). Kan skrive brøk hvis flere lager eksamen sammen.",
                    visible = FALSE)
c6 <- createComment(comment = "Antall oppgaver å sensurere, uansett uttelling",
                    visible = FALSE)
c7 <- createComment(comment = "Som ordinær sensur",
                    visible = FALSE)
c8 <- createComment(comment = "Som å lage ordinær eksamen",
                    visible = FALSE)
c9 <- createComment(comment = "Som ordinær sensur",
                    visible = FALSE)
c10 <- createComment(comment = "Antall avtalte timer ekstra. Legg til kommentar.",
                    visible = FALSE)
c11 <- createComment(comment = "Kommenter ved behov. Må fylles ut hvis navn på person mangler i listen og hvis annet er fylt ut.",
                     visible = FALSE)

# Stil på instruksjon
hs0 <- createStyle(fgFill = "#ffe6b3", textDecoration = "italic")
hs0b <- createStyle(fgFill = "#ffe6b3", textDecoration = "Bold")

# Stil på header-raden i tabellen
hs1 <- createStyle(fgFill = "#4F81BD", halign = "CENTER", textDecoration = "Bold",
                   border = "Bottom", fontColour = "white")


for(i in 1:length(emner_kode)){
  #print(i)
  
  wb <- createWorkbook()
  addWorksheet(wb, "stab", visible = FALSE)  # skjult ark med stab til drop-down meny
  writeData(wb, "stab", x = stab)
  
  filename <- paste0("emner/", emner_kode[i], ".xlsx")
  
  if(file.exists(filename)){
    print("Filen finnes fra før")
  } else{ # Filer som IKKE er opprettet ennå
   for(j in 1:length(years)){
    
      sheetname <- paste(years[j], semester[i])
      addWorksheet(wb, sheetname)
      for(k in 1:length(instruks)){
        mergeCells(wb, sheetname, 1:6, k)
      }

      writeData(wb, sheetname, startRow = 1,
                x = instruks)
      addStyle(wb, sheetname, hs0b, 1, 1)
      addStyle(wb, sheetname, hs0, 2:length(instruks), 1)
      
      writeData(wb, sheetname, x = d, startRow = length(instruks)+1, colNames = TRUE, headerStyle = hs1)
      setColWidths(wb, sheet = sheetname, cols = 1:11, widths = 20)
      
      writeComment(wb, sheetname, col = "A", row = length(instruks)+1, comment = c1)
      writeComment(wb, sheetname, col = "B", row = length(instruks)+1, comment = c2)
      writeComment(wb, sheetname, col = "C", row = length(instruks)+1, comment = c3)
      writeComment(wb, sheetname, col = "D", row = length(instruks)+1, comment = c4)
      writeComment(wb, sheetname, col = "E", row = length(instruks)+1, comment = c5)
      writeComment(wb, sheetname, col = "F", row = length(instruks)+1, comment = c6)
      writeComment(wb, sheetname, col = "G", row = length(instruks)+1, comment = c7)
      writeComment(wb, sheetname, col = "H", row = length(instruks)+1, comment = c8)
      writeComment(wb, sheetname, col = "I", row = length(instruks)+1, comment = c9)
      writeComment(wb, sheetname, col = "J", row = length(instruks)+1, comment = c10)
      writeComment(wb, sheetname, col = "K", row = length(instruks)+1, comment = c11)
      
      # data validation
      data_validation_source <- "'stab'!$A$1:$A$1000"
      
        dataValidation(wb, sheetname, 
                       col = 1, rows = length(instruks)+2:50, 
                       type = "list", 
                       value = data_validation_source)
        dataValidation(wb, sheetname,
                       col = 2:9, rows = length(instruks)+2:50, type = "decimal",
                       operator = "between", value = c(0, 60))
        dataValidation(wb, sheetname,
                       col = 10, rows = length(instruks)+2:50, type = "decimal",
                       operator = "between", value = c(0, 99))

        
        for(rad in length(instruks)+2:50){ 
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
      for(l in 1:11){
        addStyle(wb, sheetname, 
                 style = createStyle(locked = FALSE), 
                 rows = length(instruks)+2:50, cols = l)
        }
    
   }
    # Lagrer til fil
    sheetVisibility(wb)[1] <- "veryHidden"
    saveWorkbook(wb, filename, overwrite = TRUE)
  }

} 


