library(dplyr)
library(tidyr)
library(openxlsx)

# Source the styles sheet
source("styles.R")

# CSV file obtained from NISRA data portal https://data.nisra.gov.uk/table/MYE01T09

# Data read in as csv
data <- read.csv("mid-year-population-estimates-2022.csv") %>%
  # Total rows removed to avoid double counting
  filter(Sex != "All" & Local.Government.District != "Northern Ireland") %>%
  # Create a new grouped age variable
  mutate(`Age group` = case_when(Age < 25 ~ "Under 25",
                                 Age < 35 ~ "25-34",
                                 Age < 45 ~ "35-44",
                                 Age < 55 ~ "45-54",
                                 TRUE ~ "55 and over"),
         # Apply ordering of grouped age variable
         `Age group` = factor(`Age group`, levels = c("Under 25", "25-34", "35-44", "45-54", "55 and over"))) %>%
  # Use select to keep and rename some variables
  select(Year,
         `Local Government District` = Local.Government.District,
         Sex = Sex.Label,
         `Age group`,
         VALUE)

# Pivot data using new age group variable
by_age <- data %>%
  group_by(Year, `Age group`) %>%
  summarise(VALUE = sum(VALUE)) %>%
  pivot_wider(names_from = Year, values_from = VALUE)

# Pivot data using lgd variable
by_lgd <- data %>%
  group_by(Year, `Local Government District`) %>%
  summarise(VALUE = sum(VALUE)) %>%
  pivot_wider(names_from = Year, values_from = VALUE)

# Create a new workbook, give it the title "Population Estimates" and the subject "Demography Statistics" 
wb <- createWorkbook(title = "Population Estimates",
                     subject = "Demography Statistics")

# Name sheet 1 and add to workbook
sheet1 <- "Population Summary"
addWorksheet(wb, sheet1)

# Set default font option
modifyBaseFont(wb, fontSize = 12, fontName = "Arial")

# Instead of explicitly writing content to cells in particular row numbers,
# we will use a variable with the name "cursor" and instruct it to move down the page
# as we write content to our Excel sheet. Doing this allows you to insert content without
# having to update all row numbers throughout the code.
# Starting on row 1:
cursor <- 1

# Write a title for the sheet
writeData(wb, sheet1,
          "2022 Mid Year Population Summaries",
          startRow = cursor)

# Change the title to size 14 and bold
addStyle(wb, sheet1, style_page_title, rows = cursor, cols = 1)

# After writing text we will move cursor on to the next row
cursor <- cursor + 1

# Write a title for the table
writeData(wb, sheet1,
          "Population by Age Group",
          startRow = cursor)

# Change the title to bold
addStyle(wb, sheet1, style_table_title, rows = cursor, cols = 1)

# After writing text we will move cursor on to the next row
cursor <- cursor + 1

# Write the data frame by_age out as a table
writeDataTable(wb, sheet1,
               by_age,
               startRow = cursor,
               tableStyle = "none",
               withFilter = FALSE,
               headerStyle = style_table_header,
               tableName = "pop_by_age")

# Change the first row heading back to aligned left
addStyle(wb, sheet1, style_table_title, rows = cursor, cols = 1)
# Change the figures to have comma formatting
addStyle(wb, sheet1, style_table_figures, rows = (cursor + 1):(cursor + nrow(by_age) + 1), cols = 2:ncol(by_age), gridExpand = TRUE)

# Move the cursor below the by_age table using the nrow() property of the data frame:
cursor <- cursor + nrow(by_age) + 2

# Now repeat the steps above to add the LGD table to the same sheet

# Write a title for the table
writeData(wb, sheet1,
          "Population Local Government District",
          startRow = cursor)

# Change the title to bold
addStyle(wb, sheet1, style_table_title, rows = cursor, cols = 1)

# After writing text we will move cursor on to the next row
cursor <- cursor + 1

# Write the data frame by_age out as a table
writeDataTable(wb, sheet1,
               by_lgd,
               startRow = cursor,
               tableStyle = "none",
               withFilter = FALSE,
               headerStyle = style_table_header,
               tableName = "pop_by_lgd")

# Change the first row heading back to aligned left
addStyle(wb, sheet1, style_table_title, rows = cursor, cols = 1)
# Change the figures to have comma formatting
addStyle(wb, sheet1, style_table_figures, rows = (cursor + 1):(cursor + nrow(by_lgd) + 1), cols = 2:ncol(by_age), gridExpand = TRUE)

# Notice that LGD name doesn't fit on the sheet and increase the width of the first column to 31. 
setColWidths(wb, sheet1,
             cols = 1,
             widths = 31)

# See if you can follow the steps above and create a data frame by_sex using the 'Sex' field in the data
# and add this data_frame to the worksheet below the by_lgd table.
 
# See if you re-order the tables in the worksheet. By moving the code for the by_age table so that it appears
# last on the worksheet.

saveWorkbook(wb, "mid-year-population-summary-2022-further-exercise.xlsx", overwrite = TRUE)
openXL("mid-year-population-summary-2022-further-exercise.xlsx")
