library(tidyverse)
library(magrittr)
library(openxlsx)
library(readxl)
library(writexl)
library(reshape2)
library(skimr)
library(lubridate)


######### Important (Do this once a year) Week Calculation ############
first_day_of_the_year <- as.Date("2022-01-01")
first_monday_of_the_year <- as.Date("2022-01-03")


week_cal <- first_monday_of_the_year - first_day_of_the_year
week_cal %>% 
  as.numeric() -> week_cal


# Load data base ----

# load("C:/Users/slee/OneDrive - Ventura Foods/Stan/R Codes/Projects/PQA/venturafoods_PQA/rds/AS400_data_7.15.22.rds")
# load("C:/Users/slee/OneDrive - Ventura Foods/Stan/R Codes/Projects/PQA/venturafoods_PQA/rds/JDE_shopfloor_7.15.22.rds")


######## Saving to the database ####### ----
# AS400_data <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/PQA/raw/as400_database.xlsx")
# JDE_shopfloor <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/PQA/raw/jde_shopfloor.xlsx")
# 
# save(AS400_data, file = "AS400_data_7.15.22.rds")
# save(JDE_shopfloor, file = "JDE_shopfloor_7.15.22.rds")



#################### Reading input ######################

# Planner Manager reference ----
planner_manager <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/PQA/automation/Planner-Manager Reference.xlsx")
planner_manager %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  dplyr::rename(Planner = planner) %>% 
  dplyr::select(Planner, manager_name) -> planner_manager



# Address Book ----
address_book <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/PQA/automation/Address Book - 08.23.22.xlsx")

address_book %>% 
  janitor::clean_names() %>% 
  dplyr::rename(Planner = address_number,
                Planner_name = alpha_name) %>% 
  dplyr::select(1:2) -> address_book

# exception report ----
exception_report <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/PQA/automation/exception report 08.23.22.xlsx")

exception_report[-1:-2, ] -> exception_report
colnames(exception_report) <- exception_report[1, ]
exception_report[-1, ] -> exception_report

exception_report %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  dplyr::rename(Location = b_p, 
                Product = item_number,
                Planner = planner) %>% 
  dplyr::select(Location, Product, Planner) %>% 
  dplyr::mutate(ref = paste0(Location, "_", Product),
                Planner = replace(Planner, is.na(Planner), 0)) %>% 
  dplyr::select(ref, Planner) -> exception_report

# Macro-platform from S:drive/Rstudio ----
macro_platform <- read_excel("S:/Supply Chain Projects/RStudio/Macro-platform.xlsx")

macro_platform %>% 
  dplyr::rename(macro_platform = "Macro-Platform") -> macro_platform




# Platform and Category from MicroStrategy (change the file when there are new items and MS updated) ----

cat_plat <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/PQA/automation/BI Category and Platform and pack size.xlsx")
cat_plat[-1, ] -> cat_plat
colnames(cat_plat) <- cat_plat[1, ]
cat_plat[-1, ] -> cat_plat

cat_plat %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  dplyr::rename(Item = sku_code,
                Category = product_category_name,
                Platform = product_platform_description) %>% 
  dplyr::select(Item, Category, Platform) %>% 
  dplyr::mutate(Item = gsub("-", "", Item),
                Product = Item) -> cat_plat

cat_plat[!duplicated(cat_plat[,c("Item")]),] -> cat_plat




# JDE schedule attainment ----
jde_attain <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/PQA/automation/JDE schedule attainment - July 2022.xlsx",
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
as400_7499_loc25 <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/PQA/automation/AS400 schedule attainment - July 2022.xlsx",
                         sheet = "loc25 raw")

as400_7499_loc55 <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/PQA/automation/AS400 schedule attainment - July 2022.xlsx",
                               sheet = "loc55 raw")

as400_7499_loc86 <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/PQA/automation/AS400 schedule attainment - July 2022.xlsx",
                               sheet = "loc86 raw")


as400_7499_loc25[c(-1:-7, -nrow(as400_7499_loc25)), ] -> as400_7499_loc25
colnames(as400_7499_loc25) <- as400_7499_loc25[1, ]
as400_7499_loc25[-1, ] -> as400_7499_loc25

as400_7499_loc55[c(-1:-7, -nrow(as400_7499_loc55)), ] -> as400_7499_loc55
colnames(as400_7499_loc55) <- as400_7499_loc55[1, ]
as400_7499_loc55[-1, ] -> as400_7499_loc55

as400_7499_loc86[c(-1:-7, -nrow(as400_7499_loc86)), ] -> as400_7499_loc86
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
  dplyr::mutate(Date = lubridate::as_date(Date, format = "%m/%d/%y"),
                weekday = weekdays(Date)) %>% 
  dplyr::mutate(Week = as.integer(format(Date, "%U")) + 1) %>% 
  dplyr::relocate(Week, .before = Date) %>% 
  dplyr::select(-Totals) %>% 
  dplyr::mutate(Location = 25) %>% 
  dplyr::relocate(Location) %>% 
  dplyr::mutate(Scheduled_Qty = replace(Scheduled_Qty, is.na(Scheduled_Qty), 0)) %>% 
  dplyr::mutate(Scheduled_Pounds = replace(Scheduled_Pounds, is.na(Scheduled_Pounds), 0)) %>% 
  dplyr::mutate(Production_Qty = replace(Production_Qty, is.na(Production_Qty), 0)) %>% 
  dplyr::mutate(Production_Pounds = replace(Production_Pounds, is.na(Production_Pounds), 0)) -> loc25_data


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
  dplyr::mutate(Date = lubridate::as_date(Date, format = "%m/%d/%y"),
                weekday = weekdays(Date)) %>% 
  dplyr::mutate(Week = as.integer(format(Date, "%U")) + 1) %>% 
  dplyr::relocate(Week, .before = Date) %>% 
  dplyr::select(-Totals) %>% 
  dplyr::mutate(Location = 55) %>% 
  dplyr::relocate(Location) %>% 
  dplyr::mutate(Scheduled_Qty = replace(Scheduled_Qty, is.na(Scheduled_Qty), 0)) %>% 
  dplyr::mutate(Scheduled_Pounds = replace(Scheduled_Pounds, is.na(Scheduled_Pounds), 0)) %>% 
  dplyr::mutate(Production_Qty = replace(Production_Qty, is.na(Production_Qty), 0)) %>% 
  dplyr::mutate(Production_Pounds = replace(Production_Pounds, is.na(Production_Pounds), 0)) -> loc55_data

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
  dplyr::mutate(Week = as.integer(format(Date, "%U")) + 1) %>% 
  dplyr::relocate(Week, .before = Date) %>% 
  dplyr::select(-Totals) %>% 
  dplyr::mutate(Location = 86) %>% 
  dplyr::relocate(Location) %>% 
  dplyr::mutate(Scheduled_Qty = replace(Scheduled_Qty, is.na(Scheduled_Qty), 0)) %>% 
  dplyr::mutate(Scheduled_Pounds = replace(Scheduled_Pounds, is.na(Scheduled_Pounds), 0)) %>% 
  dplyr::mutate(Production_Qty = replace(Production_Qty, is.na(Production_Qty), 0)) %>% 
  dplyr::mutate(Production_Pounds = replace(Production_Pounds, is.na(Production_Pounds), 0)) -> loc86_data

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

dplyr::left_join(as400_data, cat_plat %>% select(-Item), by = "Product") %>% 
  dplyr::mutate(Month = recode(Month, "1" = "Jan", "2" = "Feb", "3" = "Mar", "4" = "Apr", "5" = "May", "6" = "Jun",
                               "7" = "Jul", "8" = "Aug", "9" = "Sep", "10" = "Oct", "11" = "Nov", "12" = "Dec")) -> as400_data


# filtering work
as400_data$Product -> temp_product
Label_as400 <- data.frame(substr(temp_product, nchar(temp_product)-2, nchar(temp_product)))

cbind(as400_data, Label_as400) -> as400_data
colnames(as400_data)[ncol(as400_data)] <- "Label"

as400_data %>% 
  dplyr::filter(Label != "BKO") %>% 
  dplyr::filter(Label != "BKM") %>% 
  dplyr::filter(Label != "TST") %>% 
  dplyr::filter(Product != "15498VEN") %>% 
  dplyr::filter(Product != "22079VEN") %>% 
  dplyr::select(-Label) -> as400_data



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
  dplyr::filter(Label != "BKM") %>% 
  dplyr::filter(Label != "TST") %>% 
  dplyr::filter(Item != "15498VEN") %>% 
  dplyr::filter(Item != "22079VEN") -> jde_attain


jde_attain %>% 
  dplyr::filter(!(Missed_or_Planned == "P" & Production_Quantity == 0)) -> jde_attain

dplyr::left_join(jde_attain, cat_plat %>% select(-Product), by = "Item") %>% 
  dplyr::select(-Label, -Date, -Lot_Code) %>% 
  dplyr::relocate(FY, Year, Month, SKU_PQA, Category, Platform) %>% 
  dplyr::mutate(Requested_Date = format(as.Date(Requested_Date), "%m/%d/%Y"),
                Modified_Req_Date = format(as.Date(Modified_Req_Date), "%m/%d/%Y")) %>% 
  dplyr::mutate(Month = recode(Month, "1" = "Jan", "2" = "Feb", "3" = "Mar", "4" = "Apr", "5" = "May", "6" = "Jun",
                               "7" = "Jul", "8" = "Aug", "9" = "Sep", "10" = "Oct", "11" = "Nov", "12" = "Dec")) -> jde_attain



# Final touch to make all_locations

jde_attain %>% 
  dplyr::select(Branch_Plant, FY, Year, Month, Line, Item, Scheduled_Quantity, Production_Quantity, SKU_PQA, Category, Platform) %>% 
  dplyr::rename(Location = Branch_Plant,
                Physical_Line = Line,
                Product = Item,
                sum_scheduled_qty = Scheduled_Quantity,
                sum_production_qty = Production_Quantity) %>% 
  dplyr::mutate(Week = "") %>% 
  dplyr::relocate(FY, Year, Month, Location, Week, Physical_Line, Product, sum_scheduled_qty, sum_production_qty, SKU_PQA, Category, Platform) -> jde_attain


as400_data %>% 
  dplyr::relocate(FY, Year, Month, Location, Week, Physical_Line, Product, sum_scheduled_qty, sum_production_qty, SKU_PQA, Category, Platform) -> as400_data


rbind(jde_attain, as400_data) -> all_locations



# Vlookup - Macro Platform
all_locations %>% 
  dplyr::left_join(macro_platform, by = "Platform") -> all_locations


# Planner
all_locations %>% 
  dplyr::mutate(ref = paste0(Location, "_", Product)) %>% 
  dplyr::left_join(exception_report, by = "ref") -> all_locations


# Planner Name 
all_locations %>% 
  dplyr::left_join(address_book, by = "Planner") -> all_locations


# Manager Name
planner_manager[!duplicated(planner_manager[,c("Planner")]),] -> planner_manager

all_locations %>% 
  dplyr::left_join(planner_manager, by = "Planner") -> all_locations


# testing begin

all_locations

testing_file <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/PQA/automation/Production Quantity Attainment rolling 24 months - Aug 2020 - Jul 2022.xlsx",
                           sheet = "All locations")


testing_file[-1:-3, ] -> testing_file
colnames(testing_file) <- testing_file[1, ]
testing_file[-1, ] -> testing_file

testing_file %>% 
  dplyr::filter(Year == "2022" & Month == "Jul") -> a




# testing end

###################################################################################################################################

######## Exporting to Excel ####### ----
writexl::write_xlsx(as400_data, "as400_data_7.15.22.xlsx")
writexl::write_xlsx(jde_attain, "jde_7.15.22.xlsx")



###### Save it to the database ##### ----
names(AS400_data) <- stringr::str_replace_all(names(AS400_data), c(" " = "_"))
head(AS400_data)

AS400_data %>% 
  dplyr::filter(!(Month == "Jul" & Year == 2020)) %>%   ## ---- here! change the oldest month ----
  dplyr::bind_rows(as400_data) -> AS400_data

save(AS400_data, file = "C:/Users/slee/OneDrive - Ventura Foods/Stan/R Codes/Projects/PQA/venturafoods_PQA/rds/AS400_data_7.15.22.rds")



head(JDE_shopfloor)
JDE_shopfloor %>% 
  dplyr::filter(!(Month == "Jul" & Year == 2020)) %>%   ## ---- here! change the oldest month ----
  dplyr::mutate(Year = as.numeric(Year)) %>% 
  dplyr::bind_rows(jde_attain) -> JDE_shopfloor

save(JDE_shopfloor, file = "C:/Users/slee/OneDrive - Ventura Foods/Stan/R Codes/Projects/PQA/venturafoods_PQA/rds/JDE_shopfloor_7.15.22.rds")





