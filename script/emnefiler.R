## Ordinær undervisning: emnefiler 

files <- list.files("emner/") 
files <- files[endsWith(files, ".xlsx") & 
                 !str_detect(files, "veiledning")]

list_all <- list(length(files))

for(i in 1:length(files)){ 
  print(files[i])
  fil <- paste0("emner/", files[i])
  
  sheet_names <- excel_sheets(fil)[-c(1:2)]  # unntatt første sheet
  
  emne_aar <- lapply(sheet_names, function(x) {          # Read all sheets to list
    as.data.frame(read_excel(fil, 
                             sheet = x,
                             skip = 8,
                             col_types = c("text", "text", rep("numeric", 9), "text")) 
    ) |> 
      mutate(aar = x,
             emne = gsub('.xlsx', '', files[i]))
  } 
  ) %>% 
    bind_rows() %>% 
    rename_with(tolower) %>% 
    mutate(navn = ifelse(navn == "Annen", kommentar, navn)) %>%  ## Hvis navn ikke finnes i nedtrekkslista skal det ligge i kommentarfeltet. Kopier fra kommentar til navn. 
    #select(-kommentar) %>% 
    pivot_longer(cols = -c(emne, aar, navn, avtalt, kommentar), 
                 names_to = "aktivitet", 
                 values_to = "timer") %>% 
    filter(!is.na(timer) ) %>%  
    mutate(navn = str_to_title(navn) %>% str_squish()) 
  
  list_all[[i]] <- emne_aar 
}



## 

if(file.exists("data/manglernavn.Rdat") & 
   difftime(Sys.time(), 
            file.info("data/manglernavn.Rdat")$mtime,
            units = "weeks") |> as.numeric() < 4 ){
  load(file = "data/forskningstermin.Rdat")
} else{


list_all_x <- list(length(files))
for(i in 1:length(files)){ 
  print(files[i])
  fil <- paste0("emner/", files[i])
  
  sheet_names <- excel_sheets(fil)[-c(1:2)]  # unntatt første sheet
  
  emne_aar_x <- lapply(sheet_names, function(x) {          # Read all sheets to list
    as.data.frame(read_excel(fil, 
                             sheet = x,
                             skip = 8,
                             col_types = c("text", "text", rep("numeric", 9), "text")) 
    ) |> 
      mutate(aar = x,
             emne = gsub('.xlsx', '', files[i]))
  } 
  ) %>% 
    bind_rows() %>% 
    rename_with(tolower) %>% 
    mutate(navn = str_to_title(navn) %>% str_squish()) 

  list_all_x[[i]] <- emne_aar_x 
}

head(list_all_x[[2]])

sjekker <- bind_rows(list_all_x) %>% 
  filter(navn == "Annen" & !is.na(kommentar)) %>% 
  mutate(nyttnavn = str_sub(kommentar, 1, 35) ) %>% 
  select(navn, aar, emne, nyttnavn) 

save(sjekker, file = "data/manglernavn.Rdat")

}