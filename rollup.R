library(dplyr)
library(readxl)
library(tidyr)

all_df = data.frame()

japan_path <- 'data/Variant_Study - Shell Table_Aggregate Data Request_Japan_FF_editsForR.xlsx'
japan_sheets <- excel_sheets(japan_path)
for(sheet in japan_sheets) {
  this_sheet_df <- read_excel(japan_path, sheet = sheet, skip = 0)
  outcome_name <- strsplit(sheet, '-')[[1]][2]
  outcome_name <- sub("1$", "", outcome_name) # remove final "1" if present
  total_missing <- sum(this_sheet_df$missing, na.rm = TRUE)
  total_all <- sum(this_sheet_df$total, na.rm = TRUE)
  percent_missing = total_missing / total_all
  all_df <- rbind(all_df,
                  data.frame(site = 'Japan',
                             outcome = outcome_name,
                             percent_missing = percent_missing))
}

philly_path <- 'data/COVID-19 Variant Study_5_31_23_AZB_FF_EditsForR.xlsx'
philly_sheets <- excel_sheets(philly_path)
for(sheet in philly_sheets) {
  this_sheet_df <- read_excel(philly_path, sheet = sheet, skip = 1)
  outcome_name <- sheet
  outcome_name <- sub("1$", "", outcome_name) # remove final "1" if present
  total_missing <- sum(this_sheet_df$missing, na.rm = TRUE)
  total_all <- sum(this_sheet_df$total, na.rm = TRUE)
  percent_missing = total_missing / total_all
  all_df <- rbind(all_df,
                  data.frame(site = 'Philly',
                             outcome = outcome_name,
                             percent_missing = percent_missing))
}

france_path <- 'data/TABLE_outcomes_france.xlsx'
france_sheets <- excel_sheets(france_path)
for(sheet in france_sheets) {
  this_sheet_df <- read_excel(france_path, sheet = sheet, skip = 1)
  outcome_name <- strsplit(sheet, '- ')[[1]][2]
  outcome_name <- sub("1$", "", outcome_name) # remove final "1" if present
  total_missing <- sum(this_sheet_df$Missing, na.rm = TRUE)
  total_all <- sum(this_sheet_df$Total, na.rm = TRUE)
  percent_missing = total_missing / total_all
  all_df <- rbind(all_df,
                  data.frame(site = 'France',
                             outcome = outcome_name,
                             percent_missing = percent_missing))
}


switzerland_path <- 'data/TABLE_outcomes_switzerland.xlsx'
switzerland_sheets <- excel_sheets(switzerland_path)
for(sheet in switzerland_sheets) {
  this_sheet_df <- read_excel(switzerland_path, sheet = sheet, skip = 1)
  # Because last tab in Switzerland has no data
  if (nrow(this_sheet_df) == 0) {
    next
  }
  outcome_name <- strsplit(sheet, '- ')[[1]][2]
  outcome_name <- sub("1$", "", outcome_name) # remove final "1" if present
  total_missing <- sum(this_sheet_df$Missing, na.rm = TRUE)
  total_all <- sum(this_sheet_df$Total, na.rm = TRUE)
  percent_missing = total_missing / total_all
  all_df <- rbind(all_df,
                  data.frame(site = 'Switzerland',
                             outcome = outcome_name,
                             percent_missing = percent_missing))
}

missingness_df <- all_df %>%
  mutate(percent_missing = round(percent_missing, 2)) %>%
  pivot_wider(names_from = outcome,
              values_from = percent_missing)


