# SKAPLAN
Versjon til deling uten data

System for undervisningsplanlegging basert på excel-filer og R/Markdown. 
Hvert emne som går planlegges bemanning i Excel-ark. SKAPLAN samler opp alt dette og kan dermed lage rapporter fordelt på person, emner, aktiviteter osv. 
Alle aktiviteter noteres som antall standardaktiviteter (f.eks. en dobbeltforelesning telles som 1), og så regner SKAPLAN ut etter gjeldende timesatser. 

SKAPLAN henter også inn informasjon om forskningsfri, frikjøp, uniped osv. omregnet i timer slik at summen av timer skal tilsvare undervisningsplikten per semester. 

SKAPLAN henter også inn saldo fra timeregnskapet forrige semester. For fast ansatte er dette kumulativt, men for midlertidige ansatte er det gjenstående timer som er det relevante tallet. 

Datastruktur og plassering av disse filene er basert på rutiner ved ISS. Det er et poeng å gjøre minst mulig "ekstra", men benytte de systemer som finnes fra før. 



