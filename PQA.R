library(tidyverse)
library(magrittr)
library(openxlsx)
library(readxl)
library(writexl)
library(reshape2)
library(skimr)


#################### Reading input ######################

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


as400_7499_loc25 %>% 
  dplyr::filter(is.na(Totals)) %>% 
  dplyr::filter(!is.na(Product)) %>% 
  dplyr::filter(!is.na(Label)) -> loc25_data

readr::type_convert(loc25_data) -> loc25_data
loc25_data %>% 
  dplyr::mutate(Date = lubridate::as_date(Date, format = )) 

# currently at 14:13