library(readxl)
library(dplyr)
library(stringr)
library(purrr)
library(lubridate)
library(writexl)

# helper: safe date extraction 
get_report_date <- function(x) {
  x <- x[[1]]
  if (inherits(x, "Date")) return(x)
  if (inherits(x, "POSIXt")) return(as.Date(x))
  # numeric Excel date
  suppressWarnings(as.Date(as.numeric(x), origin = "1899-12-30"))
}

# folder with the Balance Sheet and Income Statement data
in_dir  <- "./Input Data/Credits and Deposits"

### --- Format 1 (2021 09 to Latest) --- ###
{
files_format_1 <- list.files(
  path       = in_dir,
  pattern    = "^bs_q_[0-9]{6}_a2_bg\\.xlsx$",
  full.names = TRUE
)

process_format_1 <- function(path) {
  
  library(readxl)
  library(tidyverse)
  
  sheets <- excel_sheets(path) |> setdiff(c("МПФ 1 Методологически пояснения","МПФ 1 Методологически пояснeния"))
  
  map_dfr(sheets, function(sh) {
    dat <- read_xlsx(path, sheet = sh, col_names = FALSE)
    
    bank_name     <- as.character(dat[1, 2, drop = TRUE])
    report_date   <- get_report_date(dat[2, 2, drop = TRUE])
    
    debt_securities = dat[8:13,c(1,2,5)] %>%
      set_names(c("description", "total", "interest_income_expense")) %>%
      mutate(category = "Debt Securities")
    
    credits_and_advances = dat[18:26,c(1,2,5)] %>%
      set_names(c("description", "total", "interest_income_expense")) %>%
      mutate(category = "Credits and Advances")
    
    deposits = dat[31:37,c(1,2,5)] %>%
      set_names(c("description", "total", "interest_income_expense")) %>%
      mutate(category = "Deposits")
    
    all_data = debt_securities %>%
      bind_rows(credits_and_advances) %>%
      bind_rows(deposits) %>%
      mutate(bank_name = bank_name
             , report_date = report_date
             , excel_sheet_code = sh
      )
    
    
  })
}

bs_long_format_1 <- map_dfr(files_format_1, process_format_1)

}









