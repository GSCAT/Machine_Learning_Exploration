install.packages("dplyr")
library(dplyr)
library(magrittr)
library(readxl)
library(readr)
library(tidyr)
library(xlsx)
library(rChoiceDialogs)

# Setup Environment ----
choose_file_directory <- function()
{
  v <- jchoose.dir()
  return(v)
}

my_directory <- choose_file_directory()

# Read in Data ----
PCF_file <-  "AUR AUC 2017 - Jan Fcst.xlsx"
PCF_file2 <-  "AUR AUC 2017 - Feb Fcst.xlsx"

PCF_Forecast <- read_excel(PCF_file2, sheet = "Corp CP Essbase Pull LY KB" )
PCF_Budget <- read_excel(PCF_file, sheet = "Corp CP Essbase Pull B KB" )
PCF_LY <- read_excel(PCF_file, sheet = "Corp CP Essbase Pull LY KB" )
  
  
product_key <- read_excel("Area Product Key.xlsx", sheet = 1)
quarter_mapping <- read_csv("quarter_mapping.csv")
BMC_table <- read_csv("BMC.csv")

last_col_forecast <- length(PCF_Forecast) 
last_col_budget <- length(PCF_Budget) 
last_col_LY <- length(PCF_LY) 

# Add Source column to Forecast----
PCF_Forecast[last_col_forecast+1] <- as.factor("Forecast")
names(PCF_Forecast)[length(PCF_Forecast)] <- "Source"

# Add Source column to Budget ----
PCF_Budget[last_col_budget+1] <- as.factor("Budget")
names(PCF_Budget)[length(PCF_Budget)] <- "Source"

# Add Source column to LY ----
PCF_LY[last_col_LY+1] <- as.factor("LY")
names(PCF_LY)[length(PCF_LY)] <- "Source"


# Function for converting columns to factors ----
# Only works when on first n columns. Pass a sequence (i.e. 1:3 for first 3 columns)
conv_fact <- function(x, my_table){
  for(i in seq_along(x))
 my_table[[i]] <- as.factor(my_table[[i]])
  return(my_table)
}
# Change first number of columns (vec_1) to Factor ----

PCF_Forecast <- conv_fact(1:4, PCF_Forecast)
PCF_Budget <- conv_fact(1:4, PCF_Budget)
PCF_LY <- conv_fact(1:4, PCF_LY)

product_key <- conv_fact(1:3, product_key)
BMC_table <- conv_fact(1:8, BMC_table)
quarter_mapping <-  conv_fact(1:3, quarter_mapping)

rbind_PCF <- rbind(PCF_Forecast, PCF_Budget, PCF_LY)

# Left join to prod_key table for "Business Unit" field and arrange
PCF_post_proc <-  left_join(rbind_PCF, product_key, by= c('Area', 'Product')) %>% 
  select(`Years`, `Accounts`, `Business Unit`,`Source`, `February`, `March`, `April`, 
         `May`, `June`, `July`, `August`, `September`, `October`, `November`, `December`, `January`) %>% 
  gather("Month", "Value", 5:16) %>% 
  spread(Accounts, Value) %>% 
  left_join(quarter_mapping, by = c('Month'= 'Fiscal Month')) %>% 
  left_join(BMC_table, by = c('Business Unit' = 'BMC'))


# PCF_Budget_post_proc <-  left_join(PCF_Budget, product_key, by= c('Area', 'Product')) %>% 
#   select(`Years`, `Accounts`, `Business Unit`, `Source`, `February`, `March`, `April`, 
#          `May`, `June`, `July`, `August`, `September`, `October`, `November`, `December`) %>% 
#   gather("Month", "Value", 5:15) %>% 
#   spread(Accounts, Value) %>% 
#   left_join(quarter_mapping, by = c('Month'= 'Fiscal Month'))

# PCF_Forecast_post_proc[,5:12] <- lapply(PCF_Forecast_post_proc[, 5:12], function (x) as.numeric(x))

PCF_post_proc[,5:12] <- lapply(PCF_post_proc[, 5:12], function (x) as.numeric(x))
PCF_post_proc$Month <- as.factor(PCF_post_proc$`Month`)

# Output PCF ----
Output_PCF_Brand_Market <- PCF_post_proc %>% 
  group_by(`Fiscal Quarter`, `BMC_short_desc`) %>% 
 summarise("Forecast TY AUR of Sales" = sum(subset(`Retail$`,    Source == "Forecast"), na.rm = TRUE)/sum(subset(`Unit Sales`, Source == "Forecast"), na.rm = TRUE),
                 "TY AUC of Receipts" = sum(subset(`Cost Rcpts`, Source == "Forecast"), na.rm = TRUE)/sum(subset(`Unit Rcpts`, Source == "Forecast"), na.rm = TRUE),
                 "Budget AUR of Sales"= sum(subset(`Retail$`,    Source == "Budget"),na.rm = TRUE)/sum(subset(`Unit Sales`, Source == "Budget"), na.rm = TRUE),
             "Budget AUC of Receipts" = sum(subset(`Cost Rcpts`, Source == "Budget"), na.rm = TRUE)/sum(subset(`Unit Rcpts`, Source == "Budget"), na.rm = TRUE),
                    "LY AUR of Sales" = sum(subset(`Retail$`,    Source == "LY"), na.rm = TRUE)/sum(subset(`Unit Sales`, Source == "LY"), na.rm = TRUE),
                 "LY AUC of Receipts" = sum(subset(`Cost Rcpts`, Source == "LY"), na.rm = TRUE)/sum(subset(`Unit Rcpts`, Source == "LY"), na.rm = TRUE),
        "AUR % Change (TY vs Budget)" = as.numeric((`Forecast TY AUR of Sales`-`Budget AUR of Sales`)/`Budget AUR of Sales`)*100,
        "AUC % Change (TY vs Budget)" = as.numeric((`TY AUC of Receipts`-`Budget AUC of Receipts`)/`Budget AUC of Receipts`)*100,
            "AUR % Change (TY vs LY)" = as.numeric((`Forecast TY AUR of Sales`-`LY AUR of Sales`)/`LY AUR of Sales`)*100,
            "AUC % Change (TY vs LY)" = as.numeric((`TY AUC of Receipts`-`LY AUC of Receipts`)/`LY AUC of Receipts`)*100,
                          "GM Budget" = sum((subset(`Retail$`, Source == "Budget") - subset(`Cost$`, Source == "Budget")), na.rm = TRUE)/ sum(subset(`Retail$`, Source == "Budget"), na.rm = TRUE)*100,
                 "GM Forecast/Actual" = sum((subset(`Retail$`, Source == "Forecast") - subset(`Cost$`, Source == "Forecast")), na.rm = TRUE)/ sum(subset(`Retail$`, Source == "Forecast"), na.rm = TRUE)*100,
                              "GM LY" = sum((subset(`Retail$`, Source == "LY") - subset(`Cost$`, Source == "LY")), na.rm = TRUE)/ sum(subset(`Retail$`, Source == "LY"), na.rm = TRUE)*100,
                "GM Budget (Dollars)" = sum((subset(`Retail$`, Source == "Budget") - subset(`Cost$`, Source == "Budget")), na.rm = TRUE),
       "GM Forecast/Actual (Dollars)" = sum((subset(`Retail$`, Source == "Forecast") - subset(`Cost$`, Source == "Forecast")), na.rm = TRUE),
                    "GM LY (Dollars)" = sum((subset(`Retail$`, Source == "LY") - subset(`Cost$`, Source == "LY")), na.rm = TRUE))


# Output PCF Gap Inc ----
Output_PCF_GapInc <- PCF_post_proc %>% 
  group_by( `Fiscal Quarter`, `Gap Inc`) %>% 
  summarise("Forecast TY AUR of Sales" = sum(subset(`Retail$`,    Source == "Forecast"), na.rm = TRUE)/sum(subset(`Unit Sales`, Source == "Forecast"), na.rm = TRUE),
            "TY AUC of Receipts" = sum(subset(`Cost Rcpts`, Source == "Forecast"), na.rm = TRUE)/sum(subset(`Unit Rcpts`, Source == "Forecast"), na.rm = TRUE),
            "Budget AUR of Sales"= sum(subset(`Retail$`,    Source == "Budget"),na.rm = TRUE)/sum(subset(`Unit Sales`, Source == "Budget"), na.rm = TRUE),
            "Budget AUC of Receipts" = sum(subset(`Cost Rcpts`, Source == "Budget"), na.rm = TRUE)/sum(subset(`Unit Rcpts`, Source == "Budget"), na.rm = TRUE),
            "LY AUR of Sales" = sum(subset(`Retail$`,    Source == "LY"), na.rm = TRUE)/sum(subset(`Unit Sales`, Source == "LY"), na.rm = TRUE),
            "LY AUC of Receipts" = sum(subset(`Cost Rcpts`, Source == "LY"), na.rm = TRUE)/sum(subset(`Unit Rcpts`, Source == "LY"), na.rm = TRUE),
            "AUR % Change (TY vs Budget)" = as.numeric((`Forecast TY AUR of Sales`-`Budget AUR of Sales`)/`Budget AUR of Sales`)*100,
            "AUC % Change (TY vs Budget)" = as.numeric((`TY AUC of Receipts`-`Budget AUC of Receipts`)/`Budget AUC of Receipts`)*100,
            "AUR % Change (TY vs LY)" = as.numeric((`Forecast TY AUR of Sales`-`LY AUR of Sales`)/`LY AUR of Sales`)*100,
            "AUC % Change (TY vs LY)" = as.numeric((`TY AUC of Receipts`-`LY AUC of Receipts`)/`LY AUC of Receipts`)*100,
            "GM Budget" = sum((subset(`Retail$`, Source == "Budget") - subset(`Cost$`, Source == "Budget")), na.rm = TRUE)/ sum(subset(`Retail$`, Source == "Budget"), na.rm = TRUE)*100,
            "GM Forecast/Actual" = sum((subset(`Retail$`, Source == "Forecast") - subset(`Cost$`, Source == "Forecast")), na.rm = TRUE)/ sum(subset(`Retail$`, Source == "Forecast"), na.rm = TRUE)*100,
            "GM LY" = sum((subset(`Retail$`, Source == "LY") - subset(`Cost$`, Source == "LY")), na.rm = TRUE)/ sum(subset(`Retail$`, Source == "LY"), na.rm = TRUE)*100,
            "GM Budget (Dollars)" = sum((subset(`Retail$`, Source == "Budget") - subset(`Cost$`, Source == "Budget")), na.rm = TRUE),
            "GM Forecast/Actual (Dollars)" = sum((subset(`Retail$`, Source == "Forecast") - subset(`Cost$`, Source == "Forecast")), na.rm = TRUE),
            "GM LY (Dollars)" = sum((subset(`Retail$`, Source == "LY") - subset(`Cost$`, Source == "LY")), na.rm = TRUE))


# Output PCF by Brand Region ----
Output_PCF_Brand_Region <- PCF_post_proc %>% 
  group_by(`Fiscal Quarter`, `Brand Region`) %>% 
  summarise("Forecast TY AUR of Sales" = sum(subset(`Retail$`,    Source == "Forecast"), na.rm = TRUE)/sum(subset(`Unit Sales`, Source == "Forecast"), na.rm = TRUE),
            "TY AUC of Receipts" = sum(subset(`Cost Rcpts`, Source == "Forecast"), na.rm = TRUE)/sum(subset(`Unit Rcpts`, Source == "Forecast"), na.rm = TRUE),
            "Budget AUR of Sales"= sum(subset(`Retail$`,    Source == "Budget"),na.rm = TRUE)/sum(subset(`Unit Sales`, Source == "Budget"), na.rm = TRUE),
            "Budget AUC of Receipts" = sum(subset(`Cost Rcpts`, Source == "Budget"), na.rm = TRUE)/sum(subset(`Unit Rcpts`, Source == "Budget"), na.rm = TRUE),
            "LY AUR of Sales" = sum(subset(`Retail$`,    Source == "LY"), na.rm = TRUE)/sum(subset(`Unit Sales`, Source == "LY"), na.rm = TRUE),
            "LY AUC of Receipts" = sum(subset(`Cost Rcpts`, Source == "LY"), na.rm = TRUE)/sum(subset(`Unit Rcpts`, Source == "LY"), na.rm = TRUE),
            "AUR % Change (TY vs Budget)" = as.numeric((`Forecast TY AUR of Sales`-`Budget AUR of Sales`)/`Budget AUR of Sales`)*100,
            "AUC % Change (TY vs Budget)" = as.numeric((`TY AUC of Receipts`-`Budget AUC of Receipts`)/`Budget AUC of Receipts`)*100,
            "AUR % Change (TY vs LY)" = as.numeric((`Forecast TY AUR of Sales`-`LY AUR of Sales`)/`LY AUR of Sales`)*100,
            "AUC % Change (TY vs LY)" = as.numeric((`TY AUC of Receipts`-`LY AUC of Receipts`)/`LY AUC of Receipts`)*100,
            "GM Budget" = sum((subset(`Retail$`, Source == "Budget") - subset(`Cost$`, Source == "Budget")), na.rm = TRUE)/ sum(subset(`Retail$`, Source == "Budget"), na.rm = TRUE)*100,
            "GM Forecast/Actual" = sum((subset(`Retail$`, Source == "Forecast") - subset(`Cost$`, Source == "Forecast")), na.rm = TRUE)/ sum(subset(`Retail$`, Source == "Forecast"), na.rm = TRUE)*100,
            "GM LY" = sum((subset(`Retail$`, Source == "LY") - subset(`Cost$`, Source == "LY")), na.rm = TRUE)/ sum(subset(`Retail$`, Source == "LY"), na.rm = TRUE)*100,
            "GM Budget (Dollars)" = sum((subset(`Retail$`, Source == "Budget") - subset(`Cost$`, Source == "Budget")), na.rm = TRUE),
            "GM Forecast/Actual (Dollars)" = sum((subset(`Retail$`, Source == "Forecast") - subset(`Cost$`, Source == "Forecast")), na.rm = TRUE),
            "GM LY (Dollars)" = sum((subset(`Retail$`, Source == "LY") - subset(`Cost$`, Source == "LY")), na.rm = TRUE))

# Output PCF by Brand Market ----
Output_PCF_Brand_Market <- PCF_post_proc %>% 
  unite('Brand Market', `Brand`, `Market`, sep= ' ') %>% 
  group_by(`Fiscal Quarter`, `Brand Market`) %>% 
  summarise("Forecast TY AUR of Sales" = sum(subset(`Retail$`,    Source == "Forecast"), na.rm = TRUE)/sum(subset(`Unit Sales`, Source == "Forecast"), na.rm = TRUE),
            "TY AUC of Receipts" = sum(subset(`Cost Rcpts`, Source == "Forecast"), na.rm = TRUE)/sum(subset(`Unit Rcpts`, Source == "Forecast"), na.rm = TRUE),
            "Budget AUR of Sales"= sum(subset(`Retail$`,    Source == "Budget"),na.rm = TRUE)/sum(subset(`Unit Sales`, Source == "Budget"), na.rm = TRUE),
            "Budget AUC of Receipts" = sum(subset(`Cost Rcpts`, Source == "Budget"), na.rm = TRUE)/sum(subset(`Unit Rcpts`, Source == "Budget"), na.rm = TRUE),
            "LY AUR of Sales" = sum(subset(`Retail$`,    Source == "LY"), na.rm = TRUE)/sum(subset(`Unit Sales`, Source == "LY"), na.rm = TRUE),
            "LY AUC of Receipts" = sum(subset(`Cost Rcpts`, Source == "LY"), na.rm = TRUE)/sum(subset(`Unit Rcpts`, Source == "LY"), na.rm = TRUE),
            "AUR % Change (TY vs Budget)" = as.numeric((`Forecast TY AUR of Sales`-`Budget AUR of Sales`)/`Budget AUR of Sales`)*100,
            "AUC % Change (TY vs Budget)" = as.numeric((`TY AUC of Receipts`-`Budget AUC of Receipts`)/`Budget AUC of Receipts`)*100,
            "AUR % Change (TY vs LY)" = as.numeric((`Forecast TY AUR of Sales`-`LY AUR of Sales`)/`LY AUR of Sales`)*100,
            "AUC % Change (TY vs LY)" = as.numeric((`TY AUC of Receipts`-`LY AUC of Receipts`)/`LY AUC of Receipts`)*100,
            "GM Budget" = sum((subset(`Retail$`, Source == "Budget") - subset(`Cost$`, Source == "Budget")), na.rm = TRUE)/ sum(subset(`Retail$`, Source == "Budget"), na.rm = TRUE)*100,
            "GM Forecast/Actual" = sum((subset(`Retail$`, Source == "Forecast") - subset(`Cost$`, Source == "Forecast")), na.rm = TRUE)/ sum(subset(`Retail$`, Source == "Forecast"), na.rm = TRUE)*100,
            "GM LY" = sum((subset(`Retail$`, Source == "LY") - subset(`Cost$`, Source == "LY")), na.rm = TRUE)/ sum(subset(`Retail$`, Source == "LY"), na.rm = TRUE)*100,
            "GM Budget (Dollars)" = sum((subset(`Retail$`, Source == "Budget") - subset(`Cost$`, Source == "Budget")), na.rm = TRUE),
            "GM Forecast/Actual (Dollars)" = sum((subset(`Retail$`, Source == "Forecast") - subset(`Cost$`, Source == "Forecast")), na.rm = TRUE),
            "GM LY (Dollars)" = sum((subset(`Retail$`, Source == "LY") - subset(`Cost$`, Source == "LY")), na.rm = TRUE))

# Output PCF by Region ----
Output_PCF_Region <- PCF_post_proc %>% 
  group_by(`Fiscal Quarter`, `Region`) %>% 
  summarise("Forecast TY AUR of Sales" = sum(subset(`Retail$`,    Source == "Forecast"), na.rm = TRUE)/sum(subset(`Unit Sales`, Source == "Forecast"), na.rm = TRUE),
            "TY AUC of Receipts" = sum(subset(`Cost Rcpts`, Source == "Forecast"), na.rm = TRUE)/sum(subset(`Unit Rcpts`, Source == "Forecast"), na.rm = TRUE),
            "Budget AUR of Sales"= sum(subset(`Retail$`,    Source == "Budget"),na.rm = TRUE)/sum(subset(`Unit Sales`, Source == "Budget"), na.rm = TRUE),
            "Budget AUC of Receipts" = sum(subset(`Cost Rcpts`, Source == "Budget"), na.rm = TRUE)/sum(subset(`Unit Rcpts`, Source == "Budget"), na.rm = TRUE),
            "LY AUR of Sales" = sum(subset(`Retail$`,    Source == "LY"), na.rm = TRUE)/sum(subset(`Unit Sales`, Source == "LY"), na.rm = TRUE),
            "LY AUC of Receipts" = sum(subset(`Cost Rcpts`, Source == "LY"), na.rm = TRUE)/sum(subset(`Unit Rcpts`, Source == "LY"), na.rm = TRUE),
            "AUR % Change (TY vs Budget)" = as.numeric((`Forecast TY AUR of Sales`-`Budget AUR of Sales`)/`Budget AUR of Sales`)*100,
            "AUC % Change (TY vs Budget)" = as.numeric((`TY AUC of Receipts`-`Budget AUC of Receipts`)/`Budget AUC of Receipts`)*100,
            "AUR % Change (TY vs LY)" = as.numeric((`Forecast TY AUR of Sales`-`LY AUR of Sales`)/`LY AUR of Sales`)*100,
            "AUC % Change (TY vs LY)" = as.numeric((`TY AUC of Receipts`-`LY AUC of Receipts`)/`LY AUC of Receipts`)*100,
            "GM Budget" = sum((subset(`Retail$`, Source == "Budget") - subset(`Cost$`, Source == "Budget")), na.rm = TRUE)/ sum(subset(`Retail$`, Source == "Budget"), na.rm = TRUE)*100,
            "GM Forecast/Actual" = sum((subset(`Retail$`, Source == "Forecast") - subset(`Cost$`, Source == "Forecast")), na.rm = TRUE)/ sum(subset(`Retail$`, Source == "Forecast"), na.rm = TRUE)*100,
            "GM LY" = sum((subset(`Retail$`, Source == "LY") - subset(`Cost$`, Source == "LY")), na.rm = TRUE)/ sum(subset(`Retail$`, Source == "LY"), na.rm = TRUE)*100,
            "GM Budget (Dollars)" = sum((subset(`Retail$`, Source == "Budget") - subset(`Cost$`, Source == "Budget")), na.rm = TRUE),
            "GM Forecast/Actual (Dollars)" = sum((subset(`Retail$`, Source == "Forecast") - subset(`Cost$`, Source == "Forecast")), na.rm = TRUE),
            "GM LY (Dollars)" = sum((subset(`Retail$`, Source == "LY") - subset(`Cost$`, Source == "LY")), na.rm = TRUE))


# Output PCF by BMC ----
Output_PCF_BMC <- PCF_post_proc %>% 
  group_by(`Business Unit`, `Fiscal Quarter`) %>% 
  summarise("Forecast TY AUR of Sales" = sum(subset(`Retail$`,    Source == "Forecast"), na.rm = TRUE)/sum(subset(`Unit Sales`, Source == "Forecast"), na.rm = TRUE),
            "TY AUC of Receipts" = sum(subset(`Cost Rcpts`, Source == "Forecast"), na.rm = TRUE)/sum(subset(`Unit Rcpts`, Source == "Forecast"), na.rm = TRUE),
            "Budget AUR of Sales"= sum(subset(`Retail$`,    Source == "Budget"),na.rm = TRUE)/sum(subset(`Unit Sales`, Source == "Budget"), na.rm = TRUE),
            "Budget AUC of Receipts" = sum(subset(`Cost Rcpts`, Source == "Budget"), na.rm = TRUE)/sum(subset(`Unit Rcpts`, Source == "Budget"), na.rm = TRUE),
            "LY AUR of Sales" = sum(subset(`Retail$`,    Source == "LY"), na.rm = TRUE)/sum(subset(`Unit Sales`, Source == "LY"), na.rm = TRUE),
            "LY AUC of Receipts" = sum(subset(`Cost Rcpts`, Source == "LY"), na.rm = TRUE)/sum(subset(`Unit Rcpts`, Source == "LY"), na.rm = TRUE),
            "AUR % Change (TY vs Budget)" = as.numeric((`Forecast TY AUR of Sales`-`Budget AUR of Sales`)/`Budget AUR of Sales`)*100,
            "AUC % Change (TY vs Budget)" = as.numeric((`TY AUC of Receipts`-`Budget AUC of Receipts`)/`Budget AUC of Receipts`)*100,
            "AUR % Change (TY vs LY)" = as.numeric((`Forecast TY AUR of Sales`-`LY AUR of Sales`)/`LY AUR of Sales`)*100,
            "AUC % Change (TY vs LY)" = as.numeric((`TY AUC of Receipts`-`LY AUC of Receipts`)/`LY AUC of Receipts`)*100,
            "GM Budget" = sum((subset(`Retail$`, Source == "Budget") - subset(`Cost$`, Source == "Budget")), na.rm = TRUE)/ sum(subset(`Retail$`, Source == "Budget"), na.rm = TRUE)*100,
            "GM Forecast/Actual" = sum((subset(`Retail$`, Source == "Forecast") - subset(`Cost$`, Source == "Forecast")), na.rm = TRUE)/ sum(subset(`Retail$`, Source == "Forecast"), na.rm = TRUE)*100,
            "GM LY" = sum((subset(`Retail$`, Source == "LY") - subset(`Cost$`, Source == "LY")), na.rm = TRUE)/ sum(subset(`Retail$`, Source == "LY"), na.rm = TRUE)*100,
            "GM Budget (Dollars)" = sum((subset(`Retail$`, Source == "Budget") - subset(`Cost$`, Source == "Budget")), na.rm = TRUE),
            "GM Forecast/Actual (Dollars)" = sum((subset(`Retail$`, Source == "Forecast") - subset(`Cost$`, Source == "Forecast")), na.rm = TRUE),
            "GM LY (Dollars)" = sum((subset(`Retail$`, Source == "LY") - subset(`Cost$`, Source == "LY")), na.rm = TRUE))

Market_display <- c("US", "Canada", "Europe", "China", "Hong Kong", "Japan", "Taiwan", "Mexico")

# Output PCF by Market ----
Output_PCF_Market <- PCF_post_proc %>% 
  group_by(`Fiscal Quarter`, `Market`) %>% 
  summarise("Forecast TY AUR of Sales" = sum(subset(`Retail$`,    Source == "Forecast"), na.rm = TRUE)/sum(subset(`Unit Sales`, Source == "Forecast"), na.rm = TRUE),
            "TY AUC of Receipts" = sum(subset(`Cost Rcpts`, Source == "Forecast"), na.rm = TRUE)/sum(subset(`Unit Rcpts`, Source == "Forecast"), na.rm = TRUE),
            "Budget AUR of Sales"= sum(subset(`Retail$`,    Source == "Budget"),na.rm = TRUE)/sum(subset(`Unit Sales`, Source == "Budget"), na.rm = TRUE),
            "Budget AUC of Receipts" = sum(subset(`Cost Rcpts`, Source == "Budget"), na.rm = TRUE)/sum(subset(`Unit Rcpts`, Source == "Budget"), na.rm = TRUE),
            "LY AUR of Sales" = sum(subset(`Retail$`,    Source == "LY"), na.rm = TRUE)/sum(subset(`Unit Sales`, Source == "LY"), na.rm = TRUE),
            "LY AUC of Receipts" = sum(subset(`Cost Rcpts`, Source == "LY"), na.rm = TRUE)/sum(subset(`Unit Rcpts`, Source == "LY"), na.rm = TRUE),
            "AUR % Change (TY vs Budget)" = as.numeric((`Forecast TY AUR of Sales`-`Budget AUR of Sales`)/`Budget AUR of Sales`)*100,
            "AUC % Change (TY vs Budget)" = as.numeric((`TY AUC of Receipts`-`Budget AUC of Receipts`)/`Budget AUC of Receipts`)*100,
            "AUR % Change (TY vs LY)" = as.numeric((`Forecast TY AUR of Sales`-`LY AUR of Sales`)/`LY AUR of Sales`)*100,
            "AUC % Change (TY vs LY)" = as.numeric((`TY AUC of Receipts`-`LY AUC of Receipts`)/`LY AUC of Receipts`)*100,
            "GM Budget" = sum((subset(`Retail$`, Source == "Budget") - subset(`Cost$`, Source == "Budget")), na.rm = TRUE)/ sum(subset(`Retail$`, Source == "Budget"), na.rm = TRUE)*100,
            "GM Forecast/Actual" = sum((subset(`Retail$`, Source == "Forecast") - subset(`Cost$`, Source == "Forecast")), na.rm = TRUE)/ sum(subset(`Retail$`, Source == "Forecast"), na.rm = TRUE)*100,
            "GM LY" = sum((subset(`Retail$`, Source == "LY") - subset(`Cost$`, Source == "LY")), na.rm = TRUE)/ sum(subset(`Retail$`, Source == "LY"), na.rm = TRUE)*100,
            "GM Budget (Dollars)" = sum((subset(`Retail$`, Source == "Budget") - subset(`Cost$`, Source == "Budget")), na.rm = TRUE),
            "GM Forecast/Actual (Dollars)" = sum((subset(`Retail$`, Source == "Forecast") - subset(`Cost$`, Source == "Forecast")), na.rm = TRUE),
            "GM LY (Dollars)" = sum((subset(`Retail$`, Source == "LY") - subset(`Cost$`, Source == "LY")), na.rm = TRUE)) %>% 
  right_join(as.data.frame(Market_display), by = c("Market" = "Market_display")) %>% 
  arrange(desc(`Fiscal Quarter`))


# Output PCF by Brand and Channel ----
Output_PCF_Brand_Channel <- PCF_post_proc %>% 
  group_by(`Brand`, `Channel`, `Fiscal Quarter`) %>% 
  summarise("Forecast TY AUR of Sales" = sum(subset(`Retail$`,    Source == "Forecast"), na.rm = TRUE)/sum(subset(`Unit Sales`, Source == "Forecast"), na.rm = TRUE),
            "TY AUC of Receipts" = sum(subset(`Cost Rcpts`, Source == "Forecast"), na.rm = TRUE)/sum(subset(`Unit Rcpts`, Source == "Forecast"), na.rm = TRUE),
            "Budget AUR of Sales"= sum(subset(`Retail$`,    Source == "Budget"),na.rm = TRUE)/sum(subset(`Unit Sales`, Source == "Budget"), na.rm = TRUE),
            "Budget AUC of Receipts" = sum(subset(`Cost Rcpts`, Source == "Budget"), na.rm = TRUE)/sum(subset(`Unit Rcpts`, Source == "Budget"), na.rm = TRUE),
            "LY AUR of Sales" = sum(subset(`Retail$`,    Source == "LY"), na.rm = TRUE)/sum(subset(`Unit Sales`, Source == "LY"), na.rm = TRUE),
            "LY AUC of Receipts" = sum(subset(`Cost Rcpts`, Source == "LY"), na.rm = TRUE)/sum(subset(`Unit Rcpts`, Source == "LY"), na.rm = TRUE),
            "AUR % Change (TY vs Budget)" = as.numeric((`Forecast TY AUR of Sales`-`Budget AUR of Sales`)/`Budget AUR of Sales`)*100,
            "AUC % Change (TY vs Budget)" = as.numeric((`TY AUC of Receipts`-`Budget AUC of Receipts`)/`Budget AUC of Receipts`)*100,
            "AUR % Change (TY vs LY)" = as.numeric((`Forecast TY AUR of Sales`-`LY AUR of Sales`)/`LY AUR of Sales`)*100,
            "AUC % Change (TY vs LY)" = as.numeric((`TY AUC of Receipts`-`LY AUC of Receipts`)/`LY AUC of Receipts`)*100,
            "GM Budget" = sum((subset(`Retail$`, Source == "Budget") - subset(`Cost$`, Source == "Budget")), na.rm = TRUE)/ sum(subset(`Retail$`, Source == "Budget"), na.rm = TRUE)*100,
            "GM Forecast/Actual" = sum((subset(`Retail$`, Source == "Forecast") - subset(`Cost$`, Source == "Forecast")), na.rm = TRUE)/ sum(subset(`Retail$`, Source == "Forecast"), na.rm = TRUE)*100,
            "GM LY" = sum((subset(`Retail$`, Source == "LY") - subset(`Cost$`, Source == "LY")), na.rm = TRUE)/ sum(subset(`Retail$`, Source == "LY"), na.rm = TRUE)*100,
            "GM Budget (Dollars)" = sum((subset(`Retail$`, Source == "Budget") - subset(`Cost$`, Source == "Budget")), na.rm = TRUE),
            "GM Forecast/Actual (Dollars)" = sum((subset(`Retail$`, Source == "Forecast") - subset(`Cost$`, Source == "Forecast")), na.rm = TRUE),
            "GM LY (Dollars)" = sum((subset(`Retail$`, Source == "LY") - subset(`Cost$`, Source == "LY")), na.rm = TRUE))

# Bind Gap Inc level table ----
Gap_Inc_bind <-  rbind(Output_PCF_GapInc, Output_PCF_Brand_Region)
Gap_Inc_bind$`Gap Inc` <- as.character(Gap_Inc_bind$`Gap Inc`) 
Gap_Inc_bind$`Brand Region` <- as.character(Gap_Inc_bind$`Brand Region`) 
Gap_Inc_bind <- replace_na(Gap_Inc_bind, replace = list(`Gap Inc` = "", `Brand Region` = "")) %>% 
  unite("Gap Inc",`Gap Inc`, `Brand Region`, sep="")
Gap_Inc_bind$`Gap Inc` <- as.factor(Gap_Inc_bind$`Gap Inc`) 

Gap_Inc_display <- c("Gap Inc", "ON NA", "Gap NA", "BR NA", "GO NA", "BRFS NA", "Athleta NA") 

Gap_Inc_bind <- Gap_Inc_bind %>% 
  group_by(`Fiscal Quarter`) %>% 
  subset(`Gap Inc` %in% c(Gap_Inc_display)) %>% 
  droplevels() %>%
  right_join(as.data.frame(Gap_Inc_display), by = c("Gap Inc" = "Gap_Inc_display")) %>% 
  arrange(desc(`Fiscal Quarter`))

# Brand Level bind ----
Brand_bind <-  rbind(Output_PCF_Brand_Region, Output_PCF_BMC)
Brand_bind$`Brand Region` <- as.character(Brand_bind$`Brand Region`) 
Brand_bind$`Business Unit` <- as.character(Brand_bind$`Business Unit`) 
Brand_bind <- replace_na(Brand_bind, replace = list(`Brand Region` = "", `Business Unit` = "")) %>% 
  unite("BMC", `Brand Region`, `Business Unit`, sep = "")
Brand_bind$`BMC` <- as.factor(Brand_bind$`BMC`) 

# Specify display output by Brand ----
Gap_Brand_display <- c("Gap NA", "Gap US", "Gap Canada", "Gap Online US", "Gap Online Can", "Gap Outlet US", "Gap Outlet Canada")
BR_Brand_display <- c("BR NA", "BR US", "BR Canada", "Banana Online US", "Banana Online Canada", "BRFS US", "BRFS Canada")
ON_Brand_display <- c("ON NA", "ON US", "ON Canada", "Old Navy Online US", "Old Navy Online Canada")
BRFS_Brand_display <- c("BRFS NA", "BRFS US", "BRFS Canada")
GO_Brand_display <- c("GO NA", "Gap Outlet US", "Gap Outlet Canada")
Athleta_Brand_display <-  c("Athleta NA", "Athleta Specialty", "Athleta Online")

# Output tables ----
output_Gap_Brand <- Brand_bind %>%
  group_by(`Fiscal Quarter`) %>% 
  subset(`BMC` %in% c(Gap_Brand_display)) %>% 
  droplevels() %>%
  right_join(as.data.frame(Gap_Brand_display), by = c("BMC"= "Gap_Brand_display")) %>% 
  arrange(desc(`Fiscal Quarter`))

output_BR_Brand <- Brand_bind %>%
  group_by(`Fiscal Quarter`) %>% 
  subset(`BMC` %in% c(BR_Brand_display)) %>% 
  droplevels() %>%
  right_join(as.data.frame(BR_Brand_display), by = c("BMC"= "BR_Brand_display")) %>% 
  arrange(desc(`Fiscal Quarter`))

output_ON_Brand <- Brand_bind %>%
  group_by(`Fiscal Quarter`) %>% 
  subset(`BMC` %in% c(ON_Brand_display)) %>% 
  droplevels() %>%
  right_join(as.data.frame(ON_Brand_display), by = c("BMC"= "ON_Brand_display")) %>% 
  arrange(desc(`Fiscal Quarter`))

output_BRFS_Brand <- Brand_bind %>%
  group_by(`Fiscal Quarter`) %>% 
  subset(`BMC` %in% c(BRFS_Brand_display)) %>% 
  droplevels() %>%
  right_join(as.data.frame(BRFS_Brand_display), by = c("BMC"= "BRFS_Brand_display")) %>% 
  arrange(desc(`Fiscal Quarter`))

output_GO_Brand <- Brand_bind %>%
  group_by(`Fiscal Quarter`) %>% 
  subset(`BMC` %in% c(GO_Brand_display)) %>% 
  droplevels() %>%
  right_join(as.data.frame(GO_Brand_display), by = c("BMC"= "GO_Brand_display")) %>% 
  arrange(desc(`Fiscal Quarter`))

output_Athleta_Brand <- Brand_bind %>%
  group_by(`Fiscal Quarter`) %>% 
  subset(`BMC` %in% c(Athleta_Brand_display)) %>% 
  droplevels() %>%
  right_join(as.data.frame(Athleta_Brand_display), by = c("BMC"= "Athleta_Brand_display")) %>% 
  arrange(desc(`Fiscal Quarter`))

# Output - Gap Inc AUC AUR YOY ----
Output_GapInc_YOY <- Gap_Inc_bind %>% 
  select(`Fiscal Quarter`, `AUC % Change (TY vs LY)`, `AUR % Change (TY vs LY)`) %>% 
  group_by(`Fiscal Quarter`) %>% 
  subset(`Gap Inc` %in% c(Gap_Inc_display[1]))

# Output - Brands AUC AUR YOY 
# Output_Brands_YOY <- 
 
write.xlsx(as.data.frame(Output_PCF_Market), file = paste(my_directory, "AUC AUR Workbook.xlsx", sep = .Platform$file.sep), sheetName = "Output PCF Market", showNA = FALSE) 
write.xlsx(as.data.frame(Gap_Inc_bind), file = paste(my_directory, "AUC AUR Workbook.xlsx", sep = .Platform$file.sep), sheetName = "Output Gap Inc", append = TRUE, showNA = FALSE) 
write.xlsx(as.data.frame(output_Gap_Brand), file = paste(my_directory, "AUC AUR Workbook.xlsx", sep = .Platform$file.sep), sheetName = "Output Gap Brand", append = TRUE, showNA = FALSE) 
write.xlsx(as.data.frame(output_BR_Brand), file = paste(my_directory, "AUC AUR Workbook.xlsx", sep = .Platform$file.sep), sheetName = "Output BR Brand", append = TRUE, showNA = FALSE) 
write.xlsx(as.data.frame(output_ON_Brand), file = paste(my_directory, "AUC AUR Workbook.xlsx", sep = .Platform$file.sep), sheetName = "Output ON Brand", append = TRUE, showNA = FALSE) 
write.xlsx(as.data.frame(output_BRFS_Brand), file = paste(my_directory, "AUC AUR Workbook.xlsx", sep = .Platform$file.sep), sheetName = "Output BRFS Brand", append = TRUE, showNA = FALSE) 
write.xlsx(as.data.frame(output_GO_Brand), file = paste(my_directory, "AUC AUR Workbook.xlsx", sep = .Platform$file.sep), sheetName = "Output GO Brand", append = TRUE, showNA = FALSE) 
write.xlsx(as.data.frame(output_Athleta_Brand), file = paste(my_directory, "AUC AUR Workbook.xlsx", sep = .Platform$file.sep), sheetName = "Output Athleta Brand", append = TRUE, showNA = FALSE) 

 
# Output Function (dev) ----
output_fun <- function(x, group, out_vec){
  out_table <- x %>%
  group_by(x[[group]]) %>% 
  # subset(x[[`BMC`]] %in% c(out_vec)) %>% 
  droplevels() %>%
  right_join(as.data.frame(out_vec), by = c("BMC"= names(out_vec))) %>% 
  arrange(desc(.[[group]]))
  return(out_table)
}

output_Gap_Brand <- output_fun(Brand_bind, `Fiscal Quarter`, Gap_Brand_display)

output_BR_Brand <- output_fun(Brand_bind, group = "Fiscal Quarter", BR_Brand_display)

# Depricated code ----
# PCF_Forecast[[1]] <- as.factor(PCF_Forecast[[1]])
# PCF_Forecast[[2]] <- as.factor(PCF_Forecast[[2]])
# PCF_Forecast[[3]] <- as.factor(PCF_Forecast[[3]])
# PCF_Forecast[[4]] <- as.factor(PCF_Forecast[[4]])

# Experimantal function ----

conv_fact2 <- function(x, my_table){
  for(i in seq_along(x))
    my_table[[i]] <- as.factor(my_table[[i]])
  return(my_table)
}

vec_1 <- c(1, 3, 4 )
PCF_Forecast2 <- conv_fact2(vec_1, PCF_Forecast)