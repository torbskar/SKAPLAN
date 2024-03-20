## Ekstra avtalt timer

fil <- paste0("data/ekstra.xlsx")
sheet_names <- excel_sheets(fil)[excel_sheets(fil) != "stab"]

ekstratimer <- lapply(sheet_names, function(x) {          # Read all sheets to list
  as.data.frame(read_excel(fil, 
                           sheet = x,
                           col_types = "text") ) |> 
    mutate(aar = x,
           timer = as.numeric(timer)) 
} 
) %>% 
  bind_rows()