rm(list = ls())
cat("\014") 
options(scipen = 999)


if (!require("pacman")) install.packages("pacman")


pacman::p_load(tidyverse,
               rlang,
               dplyr,
               readxl,
               openxlsx,
               lubridate,
               data.validator,
               randomcoloR,
               mgsub,
               data.table,
               rstatix,
               stringi
)
params = list(
  kobo_tool = "inputs/YEM2203_HH_survey_tool_SBA_01252023.xlsx",
  survey_sheet = "survey",
  choices_sheet = "choices",
  data_file = "inputs/YEM2203_HH_SBA_-_all_versions_-_False_-_2023-03-23-05-31-45.xlsx",
  first_language_label_column = "label::English (en)",
  second_language_label_column = "label::Arabic (ar)"

)

source("src/util.R")
###### Loadting data, questions, ansewrs, checks######

tool <- questions <- read_xlsx(path = params$kobo_tool,
                               sheet = params$survey_sheet,
                               guess_max = 500000) %>% filter(!is.na(name))


choices <- read_xlsx(path = params$kobo_tool,
                     sheet = params$choices_sheet,
                     guess_max = 5000000) %>% filter(!is.na(name))


if(str_detect(params$data_file,pattern = "\\.csv$")) {
  df <- read.csv(params$data_file,na.strings=c("","NA","#N/A","N/A"),stringsAsFactors = F) %>% setNames(gsub("\\/",".",names(.))) %>% mutate(uuid = `_uuid`)
} else {
  df <- read_xlsx(params$data_file,
                  guess_max = 50000,
                  na = c("NA","#N/A",""," ","N/A")
  ) %>% setNames(gsub("\\/",".",names(.))) %>% mutate(
    uuid = `_uuid`
  )
}


sample_data <- df %>% select(uuid,C1_main_water_source,C2_water_source_availability,A5_characteristic,weight)

answers <- unique(na.omit(df$C1_main_water_source))

df_total_c1 <- reshape_data_total(sample_data,"C1_main_water_source")
df_total_c1_pivoted <- combine_and_pivot(df_total_c1,"C1_main_water_source")

View(df_total_c1_pivoted)
### sample data grouped grouped 
grouping_columns <- c("A5_characteristic", "C2_water_source_availability")

output <- write_grouped_values(sample_data,grouping_columns,val="C1_main_water_source")
wb <- output

saveWorkbook(wb, "outputs//output.xlsx", overwrite = TRUE)
