
for (val in 2:6) {
  file <- paste ('C:/Users/vshah/Documents/Files-Reports/2017/test/','book', val , '.csv', sep = '')
  dat <- read.csv( file ,sep = ",")

  dat1 <- rbind(dat1,dat)
   
}
dat1