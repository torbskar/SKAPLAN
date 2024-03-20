## Norskkurs
fil <- paste0("data/norskkurs.xlsx")
sheet_names <- excel_sheets(fil)[excel_sheets(fil) != "stab"]

norskkurs <- lapply(sheet_names, function(x) {          # Read all sheets to list
  as.data.frame(read_excel(fil, 
                           sheet = x,
                           col_types = "text") ) 
} 
) %>% 
  bind_rows() %>% 
  mutate(aar = paste(aar, semester),
         timer = as.numeric(timer),
         emne = "Norskkurs") %>% 
  select(-semester)