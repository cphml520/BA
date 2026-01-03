######################
####Read Packages#####
######################

library(readxl)
library(tidyverse)
library(factoextra)
library(plotly)
library(car)
library(ggplot2)

#######################
####Load dataframes####
#######################

catering <- read_excel("C:/Dokumenter/DataAnalyse/Bachelorprojekt/Kantineliste.xlsx")
guest <- read_excel("C:/Dokumenter/DataAnalyse/Bachelorprojekt/guest.xlsx")
co2 <- read_excel("C:/Dokumenter/DataAnalyse/Bachelorprojekt/co2.xlsx")
foodwaste <- read_excel("C:/Dokumenter/DataAnalyse/Bachelorprojekt/foodwaste.xlsx")
organic <- read_excel("C:/Dokumenter/DataAnalyse/Bachelorprojekt/Organic.xlsx")

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
  mutate(Delivery = if_else(Delivery == "Ja", TRUE, FALSE, missing = FALSE)
         ) %>% 
  filter(Status == "Aktiv" & ContractType == "Catering")

guest <- guest %>% 
  rename(TContractID = Agreement
         ) %>% 
  filter(Period >= 202401 & Period <= 202509
         ) %>% 
  mutate(GuestAmount = as.numeric(GuestAmount),
         Period = as.integer(Period))

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
         Fødevaregrupper = case_when(
           RiseLevel1 %in% c("Meat") ~ "Meat",
           RiseLevel1 %in% c("Seafood") ~ "Seafood",
           RiseLevel1 %in% c("Dairy products", "Eggs") ~ "Dairy",
           RiseLevel1 %in% c("Fruit and berries", "Nuts, seeds or pips",
                             "Vegetables, root crop and mushroom", 
                             "Vegetables, root vegetables and mushrooms",
                             "Grocery products and ingredients") ~ "Fruit & Vegetables",
           RiseLevel1 %in% c("Cereals and cereal products") ~ "Cornproducts",
           RiseLevel1 %in% c("Sugar and sweet foods", "Sugar and sweets") ~ "Sugar & Sweets",
           RiseLevel1 %in% c("Beverage (not milk)", "Drinks (not milk)") ~ "Drinks",
           RiseLevel1 %in% c("Fats and oils") ~ "Fats and oils",
           TRUE ~ RiseLevel1
         )
  ) %>% 
  filter(!is.na(RiseLevel1) & Period >= 202401 & Period <= 202509,
         RiseLevel3 != "Coffee, Slow coffee, supplier specific emission factor",
         Volume_kg >= 0)



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
         IsEatable = na_if(IsEatable, "-"),
         FoodWaste_kg = FoodWasteAmount / 1000
  ) %>% 
  filter(!is.na(FoodWasteAmount) & Period >= 202401 & Period <= 202509, 
         FoodWasteAmount < 200000)

foodwaste_missing <- foodwaste %>%
  anti_join(catering, by = "TContractID")

foodwaste <- foodwaste %>% 
  anti_join(foodwaste_missing, by = "TContractID")

organic <- organic %>% 
  rename(Period = YearMonth,
         FBCID = "FBC No.",
         ProductCategory = "Product Category",
         ProductDescription = "Product Description",
         Included_kg = "Included kg.",
         Organic_kg = "Organic kg.",
         OrganicPercentage = "Organic percentage"
         ) %>% 
  mutate(
    Organic_kg = replace_na(Organic_kg, 0),
    OrganicPercentage = replace_na(OrganicPercentage, 0)
    ) %>% 
  filter(Period >= 202401 & Period <= 202509)

organic_missing <- organic %>%
  anti_join(catering, by = "FBCID")

organic <- organic %>% 
  anti_join(organic_missing, by = "FBCID")

organic <- organic %>% 
  left_join(
    catering %>% select(FBCID, TContractID),
    by = "FBCID"
  )

###########################
####Feature Engineering####
###########################

CO2_monthly <- co2 %>%
  group_by(TContractID, Period) %>%
  summarise(
    total_CO2 = sum(Emission_kg),
    total_volume = sum(Volume_kg)
  )

Guests_monthly <- guest %>%
  group_by(TContractID, Period) %>%
  summarise(guests = sum(GuestAmount))

Organic_monthly <- organic %>%
  group_by(TContractID, Period) %>%
  summarise(
    organic_pct = sum(Organic_kg) / 
      sum(Included_kg))

Foodwaste_monthly <- foodwaste %>%
  group_by(TContractID, Period) %>%
  summarise(total_foodwaste = sum(FoodWaste_kg))

Product_mix <- co2 %>%
  group_by(TContractID, Period, Fødevaregrupper) %>%
  summarise(volume = sum(Volume_kg)) %>%
  group_by(TContractID, Period) %>%
  mutate(pct = volume / sum(volume)) %>%
  select(-volume) %>%
  pivot_wider(
    names_from = Fødevaregrupper,
    values_from = pct,
    values_fill = 0
  )

Product_mix <- Product_mix[,-c(12:14)]

Master <- CO2_monthly %>%
  left_join(Guests_monthly, by = c("TContractID", "Period")) %>%
  left_join(Organic_monthly, by = c("TContractID", "Period")) %>%
  left_join(Foodwaste_monthly, by = c("TContractID", "Period")) %>% 
  left_join(Product_mix, by = c("TContractID", "Period"))

Master <- Master %>%
  mutate(
    CO2_per_guest = total_CO2 / guests,
    volume_per_guest = total_volume / guests,
    foodwaste_per_guest = total_foodwaste / guests
  )

##########################
####Lineære regression####
##########################

RegressionModel <- lm(
  CO2_per_guest ~ organic_pct + Cornproducts + Dairy +
  `Fats and oils` + `Food dishes and ingredients` +
  `Fruit & Vegetables` + Meat + Seafood + `Sugar & Sweets` + 
  volume_per_guest + foodwaste_per_guest, 
  data = Master
)

summary(RegressionModel)

vif(RegressionModel)

RegressionModel2 <- lm(
  CO2_per_guest ~ Dairy + Meat + Seafood +
    `Sugar & Sweets` + volume_per_guest + foodwaste_per_guest, 
  data = Master
)

summary(RegressionModel2)

RegressionModel3 <- lm(
  CO2_per_guest ~ Dairy + Meat + 
    `Sugar & Sweets` + volume_per_guest + foodwaste_per_guest, 
  data = Master
)

summary(RegressionModel3)

#######################
####Cluster analyse####
#######################

set.seed(123)

kantine_cluster <- Master %>%
  group_by(TContractID) %>%
  summarise(
    CO2_per_guest = mean(CO2_per_guest, na.rm = TRUE),
    volume_per_guest = mean(volume_per_guest, na.rm = TRUE),
    foodwaste_per_guest = mean(foodwaste_per_guest, na.rm = TRUE),
    organic_pct = mean(organic_pct, na.rm = TRUE),
    Meat = mean(Meat, na.rm = TRUE),
    Seafood = mean(Seafood, na.rm = TRUE),
    Dairy = mean(Dairy, na.rm = TRUE),
    `Fruit & Vegetables` = mean(`Fruit & Vegetables`, na.rm = TRUE),
    Cornproducts = mean(Cornproducts, na.rm = TRUE),
    `Sugar & Sweets` = mean(`Sugar & Sweets`, na.rm = TRUE),
    `Fats and oils` = mean(`Fats and oils`, na.rm = TRUE)
  ) %>%
  ungroup()

kantine_cluster <- kantine_cluster %>%
  mutate(across(where(is.numeric), ~ ifelse(is.na(.) | is.nan(.) | is.infinite(.), 0, .)))

cluster_scaled <- scale(kantine_cluster %>% select(-TContractID))

fviz_nbclust(cluster_scaled, kmeans, method = "wss") +
  ggtitle("Elbow-method")

fviz_nbclust(cluster_scaled, kmeans, method = "silhouette") +
  ggtitle("Silhouette-method")

kmeans_model <- kmeans(cluster_scaled, centers = 3, nstart = 25)


kantine_cluster$cluster <- as.factor(kmeans_model$cluster)

fviz_cluster(kmeans_model, data = cluster_scaled,
             geom = "point",
             ellipse.type = "convex",
             main = "Kantinerne opdeler sig i tre tydelige klynger på baggrund af CO2-udledning, fødevare & madspild")

cluster_summary <- kantine_cluster %>%
  group_by(cluster) %>%
  summarise(across(where(is.numeric), mean))

print(cluster_summary)

kan_cluster <- kantine_cluster %>%
  mutate(cluster_label = case_when(
    CO2_per_guest < median(CO2_per_guest) & `Fruit & Vegetables` > median(`Fruit & Vegetables`,) ~ "Green",
    Meat > median(Meat) ~ "Protein-heavy",
    foodwaste_per_guest > median(foodwaste_per_guest) ~ "High-waste",
    TRUE ~ "Balanced"
  ))

##################################
####Visualisering af variabler####
##################################

Organic_plot <- organic %>%
  group_by(TContractID) %>%
  summarise(
    organic_pct = sum(Organic_kg) / 
      sum(Included_kg)*100)

ggplot(Organic_plot, aes(x = organic_pct)) +
  geom_histogram(
    bins = 30,
    fill = "darkgreen",
    color = "white",
    alpha = 0.8) +
  geom_vline(
    aes(xintercept = mean(organic_pct, na.rm = TRUE)),
    color = "red",
    linetype = "dashed",
    linewidth = 1) +
  scale_x_continuous(
    breaks = seq(0, 100, by = 25),
    limits = c(0, 100)) +
  labs(
    title = "Der opleves betydelig variation i økologiprocent på tværs af Coors kantiner",
    x = "Økologiprocent",
    y = "Antal kantiner") +
  theme_minimal(base_size = 13)

sd(Organic_plot$organic_pct)

ggplot(Master, aes(x = total_volume, y = CO2_per_guest)) +
  geom_point(alpha = 0.5, color = "darkgreen") +
  geom_smooth(method = "lm", se = TRUE, color = "black", linetype = "dashed") +
  labs(
    title = "CO2-udledningen stiger med råvareforbruget, men sammenhængen er svag",
    x = "Samlet mængde råvarer (kg)",
    y = "CO2-udledning pr. gæst (kg)"
  ) +
  theme_minimal(base_size = 13)

cor(Master$total_volume, Master$CO2_per_guest, use = "complete.obs")
