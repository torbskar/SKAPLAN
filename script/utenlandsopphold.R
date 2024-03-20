# Utenlandsopphold (phd/postdoc)

fil <- paste0("data/utenlandsopphold.xlsx")
sheet_names <- excel_sheets(fil)[excel_sheets(fil) != "stab"]

utlandstimer <- lapply(sheet_names, function(x) {          # Read all sheets to list
  as.data.frame(read_excel(fil, 
                           sheet = x,
                           col_types = "text") ) 
} 
) %>% 
  bind_rows() %>% 
  mutate(aar = paste(aar, semester),
         timer = as.numeric(timer),
         emne = "Utenlandsopphold fratrekk") %>% 
  select(-semester)