

for (val in 2:6) {
  file <- paste ('C:/Users/vshah/Documents/Files-Reports/2017/test/','book', val , '.csv', sep = '')
  dat <- read.csv( file ,sep = ",")
  dat1 <- rbind(dat1,dat)
   
}
dat1

# trying another method#

library (data.table)
library(plyr)
File_names <- list.files ( path="C:/Users/vshah/Documents/Files-Reports/2017/test/", pattern = ".csv") 
File_names
temp <- do.call("rbind",lapply(file_names,read.csv,header = TRUE))

temp1<-do.call(rbind, lapply(File_names, read.table, header=TRUE, sep=","))
temp1 <- ldply(file_names,read.csv)