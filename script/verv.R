# Verv og programlederansvar

fil <- paste0("data/verv.xlsx")
sheet_names <- excel_sheets(fil)[excel_sheets(fil) != "stab"]

verv <- lapply(sheet_names, function(x) {          # Read all sheets to list
  as.data.frame(read_excel(fil, 
                           sheet = x,
                           col_types = "text") ) 
} 
) %>% 
  bind_rows() %>% 
  mutate(aar = paste(aar, semester),
         antall = as.numeric(antall),
         emne = "Verv",
         aktivitet = verv) %>% 
  select(-semester, -verv)