# Load necessary library
library(tidyverse)

# Set working directory
setwd("X:/My Documents/Test esg")

# Function to read and process data in parts
read_and_process_data <- function(file, sheet, range1, range2, range3) {
  part1 <- readxl::read_excel(file, sheet = sheet, range = range1) %>%
    mutate(Years = row_number()) %>%
    pivot_longer(cols = -Years, names_to = "Stock_ID", values_to = "Value")
  
  part2 <- readxl::read_excel(file, sheet = sheet, range = range2) %>%
    mutate(Years = row_number()) %>%
    pivot_longer(cols = -Years, names_to = "Stock_ID", values_to = "Value")
  
  part3 <- readxl::read_excel(file, sheet = sheet, range = range3) %>%
    mutate(Years = row_number()) %>%
    pivot_longer(cols = -Years, names_to = "Stock_ID", values_to = "Value")
  
  bind_rows(part1, part2, part3) %>%
    arrange(Stock_ID, Years)
}

# Read and process all datasets
data_env <- read_and_process_data("Test_esg_1.xlsx", "envscore", "B1:GJJ22", "GJK1:NTR22", "NTS1:UKM22")
data_scope1 <- read_and_process_data("Test_esg_1.xlsx", "scope1", "B1:GJJ22", "GJK1:NTR22", "NTS1:UKM22")
data_scope2 <- read_and_process_data("Test_esg_1.xlsx", "scope2", "B1:GJJ22", "GJK1:NTR22", "NTS1:UKM22")
data_scope3 <- read_and_process_data("Test_esg_1.xlsx", "scope3", "B1:GJJ22", "GJK1:NTR22", "NTS1:UKM22")
data_tr <- read_and_process_data("Test_esg_1.xlsx", "Total Return", "B1:GJJ22", "GJK1:NTR22", "NTS1:UKM22")
data_mc <- read_and_process_data("Test_esg_1.xlsx", "Market Capitalization", "B1:GJJ22", "GJK1:NTR22", "NTS1:UKM22")

# Save processed datasets
saveRDS(data_env, "Data_SC.rds")
saveRDS(data_scope1, "Data_SCA.rds")
saveRDS(data_scope2, "Data_SCB.rds")
saveRDS(data_scope3, "Data_SCC.rds")
saveRDS(data_tr, "Data_TR.rds")
saveRDS(data_mc, "Data_MC.rds")

# Merge all datasets
data_final <- data_env %>%
  rename(SC = Value) %>%
  inner_join(data_scope1 %>% rename(SCA = Value), by = c("Stock_ID", "Years")) %>%
  inner_join(data_scope2 %>% rename(SCB = Value), by = c("Stock_ID", "Years")) %>%
  inner_join(data_scope3 %>% rename(SCC = Value), by = c("Stock_ID", "Years")) %>%
  inner_join(data_tr %>% rename(TR = Value), by = c("Stock_ID", "Years")) %>%
  inner_join(data_mc %>% rename(MC = Value), by = c("Stock_ID", "Years"))

saveRDS(data_final, "Data.rds")

# Perform Panel Regression Analysis
data_final <- readRDS("Data.rds") %>%
  mutate(LTA = log(MC))  # Assuming MC is Total Assets

model1 <- lm(ROE ~ SC + LTA, data = data_final)
model2 <- lm(ROE ~ SC + LTA + factor(Years), data = data_final)
model3 <- lm(ROE ~ SC + LTA + factor(Years), data = data_final, weights = ~robust)

summary(model1)
summary(model2)
summary(model3)

# Save the results
saveRDS(list(model1 = model1, model2 = model2, model3 = model3), "regression_models.rds")
