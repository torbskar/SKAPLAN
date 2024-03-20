


# Master veiledning ####
files_v <- list.files("emner/") 
files_v <- files_v[endsWith(files_v, ".xlsx") & 
                     str_detect(files_v, "veiledning")]

fil_v <- paste0("emner/", files_v)



list_veil <- list(length(fil_v))

for(i in 1:length(fil_v)){ 
  fil <- fil_v[i]
  print(str_remove(fil, "emner/") %>% str_remove("_veiledning.xlsx"))

  sheet_names <- excel_sheets(fil)[-c(1)]  # unntatt første sheet
  
  emne_aar <- lapply(sheet_names, function(x) {          # Read all sheets to list
    as.data.frame(read_excel(fil, 
                             sheet = x,
                             skip = 6,
                             col_types = c("text", rep("numeric", 2), "text")) 
    ) |> 
      mutate(aar = x,
             emne = str_remove(fil, "emner/") %>% str_remove("_veiledning.xlsx"))
  } 
  ) %>% 
    bind_rows() %>% 
    rename_with(tolower) %>% 
    #select(-kommentar) %>% 
    pivot_longer(cols = -c(emne, aar, navn, kommentar), 
                 names_to = "aktivitet", 
                 values_to = "timer") %>% 
    filter(!is.na(timer) ) %>%  
    select(-kommentar) %>% 
    mutate(navn = str_to_title(navn) %>% str_squish() ) %>% 
    mutate(aktivitet = str_remove(aktivitet, "antall") |> str_squish()) %>%  
    mutate(timer = case_when(str_sub(emne, 1, 3) == "OLA" & aktivitet == "veiledning" ~ timer * 30,
                             aktivitet == "veiledning" ~ timer * 40,
                             aktivitet == "sensur" ~ timer * 8))
  
  list_veil[[i]] <- emne_aar 
}

veiledning_ma <- list_veil %>% 
  bind_rows()
