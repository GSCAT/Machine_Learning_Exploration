library(dplyr)
library(readr)
library(RSQLServer)
library(RODBC)
library(formattable)
library(RJDBC)
library(rChoiceDialogs)
library(lubridate)
library(RCurl)

# Create RODBC connection---- 
my_connect <- odbcConnect(dsn= "IP EDWP", uid= my_uid, pwd= my_pwd)
# sqlTables(my_connect, catalog = "EDWP", tableName  = "tables")
sqlQuery(my_connect, query = "SELECT  * from dbc.dbcinfo;")


# I'm just showing Viral the coolness

 z <-  2+2 

 abc <- 50
 

