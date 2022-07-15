library(tidyverse)
library(magrittr)
library(openxlsx)
library(readxl)
library(writexl)
library(reshape2)
library(skimr)


# Load data base ----
load("AS400_data_7.15.22.rds")
load("JDE_shopfloor_7.15.22.rds")


######## Saving to the database ####### ----
# AS400_data <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/PQA/raw/as400_database.xlsx")
# JDE_shopfloor <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/PQA/raw/jde_shopfloor.xlsx")
# 
# save(AS400_data, file = "AS400_data_7.15.22.rds")
# save(JDE_shopfloor, file = "JDE_shopfloor_7.15.22.rds")




#################### Reading input ######################

# Platform and Category from MicroStrategy (change the file when there are new items and MS updated) ----
cat_plat <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Automation/raw/Category and Platform and pack size.xlsx")

cat_plat[-1:-2, ] -> cat_plat
colnames(cat_plat) <- cat_plat[1, ]
cat_plat[-1, ] -> cat_plat

cat_plat %>% 
  dplyr::select(1,2,8,9) -> cat_plat


colnames(cat_plat)[1] <- "Location"
colnames(cat_plat)[2] <- "Item"
colnames(cat_plat)[3] <- "Category"
colnames(cat_plat)[4] <- "Platform"

cat_plat %>% 
  dplyr::mutate(Item = gsub("-", "", Item),
                Product = Item) %>% 
  dplyr::select(-Location) -> cat_plat

cat_plat[-which(duplicated(cat_plat$Item)),] -> cat_plat


# JDE schedule attainment ----
jde_attain <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/PQA/raw/JDE schedule attainment - June 2022.xlsx",
                         sheet = "Sheet1")

jde_attain[-1:-2, ] -> jde_attain
colnames(jde_attain) <- jde_attain[1, ]
jde_attain[-1, ] -> jde_attain

names(jde_attain) <- stringr::str_replace_all(names(jde_attain), c("%" = "rate"))
names(jde_attain) <- stringr::str_replace_all(names(jde_attain), c(" " = "_"))

str(jde_attain)


jde_attain %>% 
  dplyr::filter(UOM != "Grand Total By Branch") %>% 
  dplyr::filter(UOM != "Subtotal :") %>% 
  dplyr::filter(UOM != "UOM") -> jde_attain


jde_attain %>% 
  dplyr::mutate(length_item = nchar(Item)) %>% 
  dplyr::filter(length_item == 8) %>% 
  dplyr::select(-length_item) -> jde_attain

readr::type_convert(jde_attain) -> jde_attain

# AS400 - 7499 ----
as400_7499_loc25 <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/PQA/raw/AS400 schedule attainment - June 2022.xlsx",
                         sheet = "loc25 raw")

as400_7499_loc55 <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/PQA/raw/AS400 schedule attainment - June 2022.xlsx",
                               sheet = "loc55 raw")

as400_7499_loc86 <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/PQA/raw/AS400 schedule attainment - June 2022.xlsx",
                               sheet = "loc86 raw")


as400_7499_loc25[-1:-7, ] -> as400_7499_loc25
colnames(as400_7499_loc25) <- as400_7499_loc25[1, ]
as400_7499_loc25[-1, ] -> as400_7499_loc25

as400_7499_loc55[-1:-7, ] -> as400_7499_loc55
colnames(as400_7499_loc55) <- as400_7499_loc55[1, ]
as400_7499_loc55[-1, ] -> as400_7499_loc55

as400_7499_loc86[-1:-7, ] -> as400_7499_loc86
colnames(as400_7499_loc86) <- as400_7499_loc86[1, ]
as400_7499_loc86[-1, ] -> as400_7499_loc86


names(as400_7499_loc25) <- stringr::str_replace_all(names(as400_7499_loc25), c(" " = "_"))
names(as400_7499_loc55) <- stringr::str_replace_all(names(as400_7499_loc55), c(" " = "_"))
names(as400_7499_loc86) <- stringr::str_replace_all(names(as400_7499_loc86), c(" " = "_"))

# Loc 25
as400_7499_loc25 %>% 
  dplyr::filter(is.na(Totals)) %>% 
  dplyr::filter(!is.na(Product)) %>% 
  dplyr::mutate(length_item = nchar(Product)) %>% 
  dplyr::filter(length_item == 8) %>% 
  dplyr::select(-length_item)-> loc25_data

readr::type_convert(loc25_data) -> loc25_data
loc25_data %>% 
  dplyr::mutate(Date = lubridate::as_date(Date, format = "%m/%d/%y")) %>% 
  dplyr::mutate(Week = lubridate::week(Date)) %>% 
  dplyr::relocate(Week, .before = Date) %>% 
  dplyr::select(-Totals) %>% 
  dplyr::mutate(Location = 25) %>% 
  dplyr::relocate(Location) -> loc25_data

reshape2::dcast(loc25_data, Location + Week + Physical_Line + Product ~ ., value.var = "Scheduled_Qty", sum) %>% 
  dplyr::rename(sum_scheduled_qty = ".") -> loc_25_pivot_1
reshape2::dcast(loc25_data, Location + Week + Physical_Line + Product ~ ., value.var = "Production_Qty", sum) %>% 
  dplyr::rename(sum_production_qty = ".") %>% 
  dplyr::select(sum_production_qty) -> loc_25_pivot_2

cbind(loc_25_pivot_1, loc_25_pivot_2) -> loc_25_pivot
loc_25_pivot %>% 
  dplyr::mutate(sum_scheduled_qty = replace(sum_scheduled_qty, is.na(sum_scheduled_qty), 0)) %>% 
  dplyr::mutate(sum_production_qty = replace(sum_production_qty, is.na(sum_production_qty), 0)) -> loc_25_pivot 

# Loc 55
as400_7499_loc55 %>% 
  dplyr::filter(is.na(Totals)) %>% 
  dplyr::filter(!is.na(Product)) %>% 
  dplyr::mutate(length_item = nchar(Product)) %>% 
  dplyr::filter(length_item == 8) %>% 
  dplyr::select(-length_item)-> loc55_data

readr::type_convert(loc55_data) -> loc55_data
loc55_data %>% 
  dplyr::mutate(Date = lubridate::as_date(Date, format = "%m/%d/%y")) %>% 
  dplyr::mutate(Week = lubridate::week(Date)) %>% 
  dplyr::relocate(Week, .before = Date) %>% 
  dplyr::select(-Totals) %>% 
  dplyr::mutate(Location = 55) %>% 
  dplyr::relocate(Location) -> loc55_data

reshape2::dcast(loc55_data, Location + Week + Physical_Line + Product ~ ., value.var = "Scheduled_Qty", sum) %>% 
  dplyr::rename(sum_scheduled_qty = ".") -> loc_55_pivot_1
reshape2::dcast(loc55_data, Location + Week + Physical_Line + Product ~ ., value.var = "Production_Qty", sum) %>% 
  dplyr::rename(sum_production_qty = ".") %>% 
  dplyr::select(sum_production_qty) -> loc_55_pivot_2

cbind(loc_55_pivot_1, loc_55_pivot_2) -> loc_55_pivot
loc_55_pivot %>% 
  dplyr::mutate(sum_scheduled_qty = replace(sum_scheduled_qty, is.na(sum_scheduled_qty), 0)) %>% 
  dplyr::mutate(sum_production_qty = replace(sum_production_qty, is.na(sum_production_qty), 0)) -> loc_55_pivot

# Loc 86
as400_7499_loc86 %>% 
  dplyr::filter(is.na(Totals)) %>% 
  dplyr::filter(!is.na(Product)) %>% 
  dplyr::mutate(length_item = nchar(Product)) %>% 
  dplyr::filter(length_item == 8) %>% 
  dplyr::select(-length_item)-> loc86_data

readr::type_convert(loc86_data) -> loc86_data
loc86_data %>% 
  dplyr::mutate(Date = lubridate::as_date(Date, format = "%m/%d/%y")) %>% 
  dplyr::mutate(Week = lubridate::week(Date)) %>% 
  dplyr::relocate(Week, .before = Date) %>% 
  dplyr::select(-Totals) %>% 
  dplyr::mutate(Location = 86) %>% 
  dplyr::relocate(Location) -> loc86_data

reshape2::dcast(loc86_data, Location + Week + Physical_Line + Product ~ ., value.var = "Scheduled_Qty", sum) %>% 
  dplyr::rename(sum_scheduled_qty = ".") -> loc_86_pivot_1
reshape2::dcast(loc86_data, Location + Week + Physical_Line + Product ~ ., value.var = "Production_Qty", sum) %>% 
  dplyr::rename(sum_production_qty = ".") %>% 
  dplyr::select(sum_production_qty) -> loc_86_pivot_2

cbind(loc_86_pivot_1, loc_86_pivot_2) -> loc_86_pivot
loc_86_pivot %>% 
  dplyr::mutate(sum_scheduled_qty = replace(sum_scheduled_qty, is.na(sum_scheduled_qty), 0)) %>% 
  dplyr::mutate(sum_production_qty = replace(sum_production_qty, is.na(sum_production_qty), 0)) -> loc_86_pivot 


# combine 3 pivots 
rbind(loc_25_pivot, loc_55_pivot, loc_86_pivot) -> as400_pivot

as400_pivot %>% 
  dplyr::mutate(Date = Sys.Date(),
                FY = paste0("FY", " ",lubridate::year(Date) + 1),
                Year = lubridate::year(Date),
                Month = lubridate::month(Date) -1,
                SKU_PQA = sum_production_qty / sum_scheduled_qty,
                SKU_PQA = replace(SKU_PQA, is.infinite(SKU_PQA), 0),
                SKU_PQA = replace(SKU_PQA, is.nan(SKU_PQA), 0)) %>% 
  dplyr::relocate(Location, FY, Year, Month, Week, Physical_Line, Product) %>% 
  dplyr::select(-Date) -> as400_data

dplyr::left_join(as400_data, cat_plat %>% select(-Item), by = "Product") -> as400_data


################# move to JDE Shopfloor tab
str(jde_attain)
jde_attain %>% 
  dplyr::mutate(Requested_Date = as.integer(Requested_Date),
                Requested_Date = as.Date(Requested_Date, origin = "1899-12-30"),
                Modified_Req_Date = as.integer(Modified_Req_Date),
                Modified_Req_Date = as.Date(Modified_Req_Date, origin = "1899-12-30"),
                Date = Sys.Date(),
                FY = paste0("FY", " ",lubridate::year(Date) + 1),
                Year = lubridate::year(Date),
                Month = lubridate::month(Date) -1,
                SKU_PQA = Production_Quantity / Scheduled_Quantity,
                SKU_PQA = replace(SKU_PQA, is.infinite(SKU_PQA), 0),
                SKU_PQA = replace(SKU_PQA, is.nan(SKU_PQA), 0)) %>% 
  dplyr::filter(Work_Order_Type %in% c("WS", "WO")) -> jde_attain


jde_attain$Item -> temp_item
Label <- data.frame(substr(temp_item, nchar(temp_item)-2, nchar(temp_item)))

cbind(jde_attain, Label) -> jde_attain
colnames(jde_attain)[ncol(jde_attain)] <- "Label"

jde_attain %>% 
  dplyr::filter(Label != "BKO") %>% 
  dplyr::filter(Label != "TST") %>% 
  dplyr::filter(Label != "15498VEN") %>% 
  dplyr::filter(Label != "22079VEN") -> jde_attain


jde_attain %>% 
  dplyr::filter(!(Missed_or_Planned == "P" & Production_Quantity == 0)) -> jde_attain

dplyr::left_join(jde_attain, cat_plat %>% select(-Product), by = "Item") %>% 
  dplyr::select(-Label, -Date, -Lot_Code) %>% 
  dplyr::relocate(FY, Year, Month, SKU_PQA, Category, Platform) %>% 
  dplyr::mutate(Requested_Date = format(as.Date(Requested_Date), "%m/%d/%Y"),
                Modified_Req_Date = format(as.Date(Modified_Req_Date), "%m/%d/%Y")) %>% 
  dplyr::mutate(Month = recode(Month, "1" = "Jan", "2" = "Feb", "3" = "Mar", "4" = "Apr", "5" = "May", "6" = "Jun",
                               "7" = "Jul", "8" = "Aug", "9" = "Sep", "10" = "Oct", "11" = "Nov", "12" = "Dec")) -> jde_attain

######## Exporting to Excel ####### ----
writexl::write_xlsx(as400_data, "as400_data_7.15.22.xlsx")
writexl::write_xlsx(jde_attain, "jde_7.15.22.xlsx")



