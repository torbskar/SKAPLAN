# Oppdater med stab i alle regneark #### 

library(tidyverse)
library(readxl)
library(openxlsx)

invisible(Sys.setlocale(locale='no_NB.utf8'))


# Felles script for liste med folk og rydding ####
# Brukes objektet "navneliste" videre
## Oversikt over person #### 
# Hvis eldre enn 2 uker, kjør på nytt
if(file.exists("data/stab.Rdat") & 
   difftime(Sys.time(), 
            file.info("data/eksterne_undervisere.xlsx")$mtime,
            units = "weeks") |> as.numeric() > 2 ){
  load(file = "data/stab.Rdat")
} else{
  source("script/stab_andreFolk.R", local = TRUE)
}


# Laster arbeidsbok

emnefiler <- list.files("emner/")
emnefiler <- emnefiler[str_sub(emnefiler, -5, -1) == ".xlsx"] 
emnefiler <- emnefiler[!(emnefiler %in% c("ekstra.xlsx", "veiledning_MA.xlsx", "veiledning_phd.xlsx","uniped.xlsx"))]


# Loop gjennom alle filer. Les inn og erstatt fanen "stab"
for(i in 1:length(emnefiler)){ 
  print(emnefiler[i])
  filnavn <- paste0("emner/", emnefiler[i])
  wb <- loadWorkbook(filnavn)
  deleteData(wb, "stab", cols=1:2, rows = 1:999, gridExpand = TRUE)

  writeData(wb, "stab", x = navneliste)
  sheetVisibility(wb)[1:2] <- "veryHidden"

  saveWorkbook(wb, filnavn, overwrite = TRUE)
  }



andrefiler <- list.files("data/") 
andrefiler <- andrefiler[endsWith(andrefiler, ".xlsx")]
andrefiler <- andrefiler[(str_detect(tolower(andrefiler), 
                                     "ekstra|uniped|nork|frikj|andre_oppg|forskningst|verv|veiledning"))]


#getSheetNames("data/verv.xlsx")
#i <- 7
# Loop gjennom alle filer. Les inn og erstatt fanen "stab"
for(i in 1:length(andrefiler)){ 
  print(andrefiler[i])
  filnavn <- paste0("data/", andrefiler[i])
  fanenavn <- getSheetNames(filnavn)    
  
  if("stab" %in% fanenavn){
      wb <- loadWorkbook(filnavn)
      deleteData(wb, "stab", cols=1:2, rows = 1:999, gridExpand = TRUE)
      
      writeData(wb, "stab", x = (navneliste %>% pull(navn)) )
      sheetVisibility(wb)[which(names(wb)=="stab")] <- "veryHidden"

      saveWorkbook(wb, filnavn, overwrite = TRUE)
  }else{
    print(paste("Arket stab finnes ikke i", filnavn))
  }
}

