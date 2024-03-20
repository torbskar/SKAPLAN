

library(openxlsx)
library(tidyverse)
filnavn <- "emner/SOS1004 - Copy.xlsx"

# Laster arbeidsbok
wbx <- loadWorkbook(filnavn, xlsxFile = NULL)

# trekker ut navn på ark 
sheetname <- names(wbx)[-1]
instruks <- read.xlsx(wbx, sheet = sheetname[1], rows=1:6 )  # Eksisterende instruks 

datadel <- read.xlsx(wbx, sheet = sheetname[1], start = 7) 

# Tester med nytt ark 
nyttsheet <- "tester"

addWorksheet(wbx, nyttsheet)

setColWidths(wbx, sheet = nyttsheet, cols = 1:10, widths = 20)
setColWidths(wbx, sheet = nyttsheet, cols = 11, widths = 50)

#removeCellMerge(wbx, sheet, cols, rows
for(k in 1:nrow(instruks)){
  mergeCells(wbx, nyttsheet, 1:10, k)
}

# Stil på instruksjon
hs0 <- createStyle(fgFill = "#ffe6b3", textDecoration = "italic", wrapText = TRUE)
hs0b <- createStyle(fgFill = "#ffe6b3", textDecoration = "Bold", wrapText = TRUE)

# Stil på header-raden i tabellen
hs1 <- createStyle(fgFill = "#4F81BD", halign = "CENTER", textDecoration = "Bold",
                   border = "Bottom", fontColour = "white", wrapText = TRUE)

writeData(wbx, nyttsheet, startRow = 1,
          x = instruks)
writeData(wbx, nyttsheet, startRow = 6, x = "DU MA REDIGERE FILEN I TEAMS")




# Legg comments i egen celle i stedet for

celleinstruks <- c("Fullt navn må staves nøyaktig. Bruk drop-down listen. 
Hvis du åpner filen i webapplikasjonen vil autofill fungere. 
MERK! Hvis du ikke finner navnet i listen, velg 'Annen' og skriv navnet i kommentarfeltet i siste kollonne.", 
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

writeData(wbx, nyttsheet, startRow = 8, x = celleinstruks, colNames = FALSE)
setRowHeights(wbx, nyttsheet, 8, 190)
addStyle(wbx, nyttsheet, hs0, 8, 1:11)

addStyle(wbx, nyttsheet, hs0b, 1, 1)
addStyle(wbx, nyttsheet, hs0, 2:nrow(instruks)+1, 1)
addStyle(wbx, nyttsheet, hs0b, nrow(instruks)+3, 1)


writeData(wbx, nyttsheet, startRow = 9, x = datadel, 
          colNames = TRUE, headerStyle = hs1)


saveWorkbook(wbx, filnavn, overwrite = TRUE)
