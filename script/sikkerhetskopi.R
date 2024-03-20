
dato <- Sys.Date()

# Lag zip kopi av emner
if(dir.exists("arkiv")){ 
  files2zip <- dir('emner', full.names = TRUE)
  zip::zip(zipfile = paste0('arkiv/emner_Zip_versjon_', dato, ".zip"), files = files2zip)
} 

# Lag zip kopi av data
if(dir.exists("arkiv")){ 
  files2zip <- dir('data', full.names = TRUE)
  zip::zip(zipfile = paste0('arkiv/data_Zip_versjon_', dato, ".zip"), files = files2zip)
} 


# Lag zip kopi av personrapport
if(dir.exists("arkiv")){ 
  files2zip <- dir('personrapport', full.names = TRUE)
  zip::zip(zipfile = paste0('arkiv/personrapport_Zip_versjon_', dato, ".zip"), files = files2zip)
} 


# # Lag zip kopi av script
# if(dir.exists("arkiv")){ 
#   files2zip <- dir('script', full.names = TRUE)
#   zip::zip(zipfile = paste0('arkiv/script_Zip_versjon_', dato, ".zip"), files = files2zip)
# } 
