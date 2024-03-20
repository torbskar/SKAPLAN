## Planlagt Uniped

fil <- paste0("data/uniped.xlsx")
sheet_names <- excel_sheets(fil)[excel_sheets(fil) != "stab"]

unipedtimer <- lapply(sheet_names, function(x) {          # Read all sheets to list
  as.data.frame(read_excel(fil, 
                           sheet = x,
                           col_types = "text") ) 
} 
) %>% 
  bind_rows() %>% 
  mutate(aar = paste(aar, semester),
         timer = as.numeric(timer),
         emne = "Uniped") %>% 
  select(-semester)