#####################
####Load Packages####
#####################

library(DBI)
library(odbc)
library(tidyverse)
library(lubridate)
library(tibble)
library(readxl)

######################
####Data Retrieval####
######################

catering <- read_excel("C:/Users/c2503054/OneDrive - Coor Service Management/Documents/Dataanalyse - Praktik/Bachelorprojekt/Kantineliste.xlsx")
period <- read_excel("C:/Users/c2503054/OneDrive - Coor Service Management/Documents/Dataanalyse - Praktik/Bachelorprojekt/YearMonth.xlsx")
guest <- read_excel("C:/Users/c2503054/OneDrive - Coor Service Management/Documents/Dataanalyse - Praktik/Bachelorprojekt/Guest.xlsx")
co2 <- read_excel("C:/Users/c2503054/OneDrive - Coor Service Management/Documents/Dataanalyse - Praktik/Bachelorprojekt/CO2.xlsx")
foodwaste <- read_excel("C:/Users/c2503054/OneDrive - Coor Service Management/Documents/Dataanalyse - Praktik/Bachelorprojekt/Foodwaste.xlsx")
organic <- read_excel("C:/Users/c2503054/OneDrive - Coor Service Management/Documents/Dataanalyse - Praktik/Bachelorprojekt/Organic.xlsx")

###################################
####Data Preparation & Cleaning####
###################################

catering <- catering %>% 
  rename(ContractName = Kontraktnavn,
         FBCID = "FBC no.",
         TContractID = "T-aftale",
         Contract = Kontrakt,
         ContractType = Kontrakttype,
         Delivery = Leverance
  ) %>% 
  mutate(Delivery = if_else(Delivery == "Ja", TRUE, FALSE, missing = FALSE))

period <- period %>% 
  rename(Year = År,
         Period = YearMonth,
         MonthText = Måned_text,
         CorrectionFactor = "CO2 Korrektionsfaktor"
  ) %>% 
  mutate(Date = as.Date(paste0(Period, "01"), format = "%Y%m%d"),
         Year = as.integer(Year),
         Period = as.integer(Period)
  )

guest <- guest %>% 
  rename(TContractID = Agreement
         ) %>% 
  select(Period, TContractID, GuestAmount
         ) %>% 
  mutate(Period = as.integer(Period),
         GuestAmount = as.integer(GuestAmount)
         ) %>% 
  filter(Period <= 202512)

co2 <- co2 %>% 
  rename(TContractID = Agreement,
         RiseLevel1 = "RISE Category Level 1",
         RiseLevel2 = "RISE Category Level 2",
         RiseLevel3 = "RISE Category Level 3",
         Supplier = "F&B Supplier",
         Emission_kg = "Emission (kg)",
         Volume_kg = "Volume (kg)",
         CO2ePerKgActual = "CO2e per kg - Actual",
         Co2ePerKgInternalTarget = "CO2e per kg - Target Internal"
  ) %>% 
  filter(!is.na(Year)
  ) %>% 
  mutate(Year = as.integer(Year),
         Period = as.integer(Period),
         RiseLevel1 = na_if(RiseLevel1, "-"),
         RiseLevel2 = na_if(RiseLevel2, "-"),
         RiseLevel3 = na_if(RiseLevel3, "-"),
         Supplier = na_if(Supplier, "-"),
         Co2ePerKgInternalTarget = na_if(Co2ePerKgInternalTarget, "-"),
         Co2ePerKgInternalTarget = as.numeric(Co2ePerKgInternalTarget)
  ) %>% 
  filter(!is.na(RiseLevel1),
         Period >= 202401)

co2_missing <- co2 %>%
  anti_join(catering, by = "TContractID")

co2 <- co2 %>% 
  anti_join(co2_missing, by = "TContractID")

foodwaste <- foodwaste %>%
  rename(TContractID = Agreement,
         FoodWasteCategory = "Food Waste Category",
         FoodWasteSubcategory = "Food Wast Subcategory"
  ) %>% 
  filter(!is.na(Year)
  ) %>% 
  mutate(Year = as.integer(Year),
         Period = as.integer(Period),
         FoodWasteAmount = as.numeric(FoodWasteAmount),
         FoodWasteCategory = na_if(FoodWasteCategory, "-"),
         FoodWasteSubcategory = na_if(FoodWasteSubcategory, "-"),
         IsEatable = na_if(IsEatable, "-")
  ) %>% 
  filter(!is.na(FoodWasteAmount),
         Period <= 202512)

foodwaste_missing <- foodwaste %>%
  anti_join(catering, by = "TContractID")

foodwaste <- foodwaste %>% 
  anti_join(foodwaste_missing, by = "TContractID")

organic <- organic %>% 
  rename(Period = YearMonth,
         FBCID = "FBC No.",
         ProductCategory = "Product Category",
         ProductDescription = "Product Description",
         OrganicKg  = "Organic kg.",
         TotalKg = "Included kg.",
         OrganicPercentage = "Organic percentage",
         Target = Minimumskrav
  ) %>%
  filter(!is.na(FBCID) & FBCID != ""
         ) %>% 
  mutate(Year = as.integer(Year),
         Period = as.integer(Period))

organic_missing <- organic %>%
  anti_join(catering, by = "FBCID")

organic <- organic %>% 
  anti_join(organic_missing, by = "FBCID")

#################################
####Empty List for Clean Data####
#################################

# Load dataframes into a list
tables <- list(
  catering = catering,
  period = period,
  guest = guest,
  co2 = co2,
  foodwaste = foodwaste,
  organic = organic
)

#############################
####Connect to SQL Server####
#############################

# Connect to database
con <- dbConnect(
  odbc(),
  Driver   = "ODBC Driver 17 for SQL Server",
  Server   = "coorsqlhotelp01",
  Database = "DK_Catering",
  Trusted_Connection = "Yes"
)

##############################################
####Transfer Data & Update Existing Tables####
##############################################

# Loop list with dataframes and load SQL tables with new data

for (tbl in names(tables)) {
  dbWriteTable(
    con,
    name   = tbl,              # tabellenavnet i SQL
    value  = tables[[tbl]],    # dataframe i R
    append = TRUE,             # tilføj data
    row.names = FALSE          # typisk en god idé
  )
}
