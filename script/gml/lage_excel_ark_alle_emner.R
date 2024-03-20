
library(tidyverse)
library(lubridate)

#year(Sys.Date())

alleemner <- readxl::read_excel("Undervisningsplan_Sosiologi_kopi.xlsx") %>% 
  rename_with(tolower) %>% 
  setNames( c(names(.)[1:5], paste0('aar_', names(.)[6:ncol(.)]) ))



emneplan <- alleemner %>% 
  select(1:5, num_range("aar_", year(Sys.Date()):2028)) %>%
  filter(is.na(avsluttet)) %>% 
  mutate(emnekode = as.character(map(strsplit(emne, split = " "), 1))) %>% 
  mutate(emnekode = str_remove(emnekode, ",")) %>% 
  group_by(emnekode) %>% 
  slice(1) %>% 
  ungroup()


glimpse(emneplan)


#emner <- paste0(emneplan$emnekode, "_", str_sub(emneplan$semester, 1,1) ) 
emner_kode <- emneplan$emnekode
semester <- str_sub(emneplan$semester, 1,1)
length(emner_kode) == length(semester)


years <- year(Sys.Date()):2028
columns <- c("Navn",	"Emneansvar",	"Forelesning",	"Seminargrupper",	"Lage_eksamen",	"Ordinær_sensur",	"Klagesensur",	"Lage_utsatt_eksamen",	"Sensur_utsatt",	"Annet",	"Kommentar")

d <- data.frame((matrix(ncol = length(columns), nrow = 1)))
names(d)<-columns


i <- 1
ifelse(emneplan$emnenavn == emner_kode[i]){
  planaar <- emneplan[i,] %>% 
    pivot_longer(cols = 6:(ncol(.)-1), names_to = "varnavn", values_to = "person") %>% 
    mutate(aar = str_sub(varnavn, -4, -1)) %>% 
    filter(!is.na(person)) %>% 
    select(emnekode, semester, nivå, aar, person)

}



#install.packages("xlsx")

library(xlsx)

emner_kode[1]

for(i in 1:length(emner_kode)){
  
  filename <- paste0("test/", emner_kode[i], ".xlsx")
  
  planaar <- emneplan[i,] %>% 
    pivot_longer(cols = 6:(ncol(.)-1), names_to = "varnavn", values_to = "person") %>% 
    mutate(aar = str_sub(varnavn, -4, -1)) %>% 
    filter(!is.na(person)) %>% 
    select(emnekode, semester, nivå, aar, person) %>% 
    replace(is.na(.), 0)

  for(j in 1:length(years)){
    
    d$Navn <- planaar$person[j]
    
    if(file.exists(filename)){
      
      # Emner over to semestre
      if(emner_kode[i] == "KULKOM1001"){ 
        for(k in c("H", "V")){
          write.xlsx(d, file = filename, sheetName = paste(years[j], k), append = TRUE, row.names = FALSE, showNA = FALSE)
        }
        
      } else { # ALLE andre emner
        write.xlsx(d, file = filename, sheetName = paste(years[j]), append = TRUE, row.names = FALSE, showNA = FALSE)
      }
      
      } else{ # Filer som IKKE er opprettet ennå
        # Emner over to semestre
        if(emner_kode[i] == "KULKOM1001"){ 
          for(k in c("H", "V")){
            if(k=="H"){
              write.xlsx(d, file = filename, sheetName = paste(years[j], k), append = FALSE, row.names = FALSE, showNA = FALSE)
              
            } else{  # ALLE andre emner
                write.xlsx(d, file = filename, sheetName = paste(years[j], k), append = TRUE, row.names = FALSE, showNA = FALSE)
              }
          }
        } else {
          
        write.xlsx(d, file = filename, sheetName = paste(years[j]), append = FALSE, row.names = FALSE, showNA = FALSE)
        }
      }
  }
}



# require(openxlsx)
# 
# wb <- createWorkbook(creator = "TS", title = emner[1], subject = "planlegging")
# 
# #add sheet and enable gridlines
# for(j in 1:length(years)){
#   addWorksheet(wb, x=d, sheet = paste(years[j]), gridLines = TRUE)
# }
# 
# 
# writeData(wb, 
#           sheet=paste(years[j]), 
#           d, 
#           paste0("test/", emner[i], ".xlsx"))
             