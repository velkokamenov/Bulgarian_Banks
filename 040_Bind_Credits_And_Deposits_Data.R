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
  
  # character in DDMMYYYY format
  if (is.character(x) && grepl("^[0-9]{8}$", x)) {
    return(as.Date(x, format = "%d%m%Y"))
  }
  
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

cd_long_format_1 <- map_dfr(files_format_1, process_format_1)

}

### --- Format 2 (2015 03 to 2021 06) --- ###
{
files_format_2 <- list.files(
  path       = in_dir,
  pattern    = "^bs_q_((2015(0[3-9]|1[0-2]))|(201[6-9][0-1][0-9])|(2020(0[1-9]|1[0-2]))|(2021(0[1-6])))_a2_bg\\.xls$",
  full.names = TRUE
)

process_format_2 <- function(path) {
  
  library(readxl)
  library(tidyverse)
  
  sheets <- excel_sheets(path) |> setdiff(c("МПФ 1 Методологически пояснения","МПФ 1 Методологически пояснeния"))
  
  map_dfr(sheets, function(sh) {
    dat <- read_xls(path, sheet = sh, col_names = FALSE)
    
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

cd_long_format_2 <- map_dfr(files_format_2, process_format_2)

}

### --- Format 3 (2008 12 to 2014 12) --- ###
{
files_format_3 <- list.files(
  path       = in_dir,
  pattern    = "^bs_q_((2008(1[2]))|(2009(0[1-9]|1[0-2]))|(201[0-3](0[1-9]|1[0-2]))|(2014(0[1-9]|1[0-2])))_a[2-4]_bg\\.xls$",
  full.names = TRUE
)

process_format_3 <- function(path) {
  
  library(readxl)
  library(tidyverse)
  
  sheets <- excel_sheets(path) |> setdiff(c("Sheet2","Sheet3"))
  
  map_dfr(sheets, function(sh) {
    dat <- read_xls(path, sheet = sh, col_names = FALSE)
    
    header1 = coalesce(as.character(dat[1,1]),"NA")
    header2 = coalesce(as.character(dat[1,2]),"NA")
    header3 = coalesce(as.character(dat[2,2]),"NA")
    
    if (header1 == "40. Капиталови, дългови и компенсаторни инструменти,  инвестиции в асоциирани, дъщерни и съвместни предприятия (осчетоводени, прилагайки капиталовия метод - включително репутация), кредити и аванси, привлечени средства и справочна информация.") {
      
      bank_name     <- as.character(dat[12, 2, drop = TRUE])
      report_date   <- get_report_date(dat[11, 7, drop = TRUE])
      
      debt_securities = dat[17:31,c(2,3,7)] %>%
        set_names(c("description", "total", "interest_income_expense")) %>%
        mutate(category = "Debt Securities") %>%
        mutate(bank_name = bank_name
               , report_date = report_date
               , excel_sheet_code = sh
        )
      
    }
    
    if (header2 == "40Б1 – Кредити и аванси (брутни) по видове валути, информация за специфичните провизии за загуби от обезценка и приходите от  лихви") {
      
      bank_name     <- as.character(dat[9, 2, drop = TRUE])
      report_date   <- get_report_date(dat[8, 8, drop = TRUE])
      
      credits_and_advances = dat[14:21,c(2,3,8)] %>%
        set_names(c("description", "total", "interest_income_expense")) %>%
        mutate(category = "Credits and Advances") %>%
        mutate(bank_name = bank_name
               , report_date = report_date
               , excel_sheet_code = sh
        )
      
    }
    
    if (header3 == "40В1. Привлечени средства по видове валути и информация за разходите за лихви") {
      
      bank_name     <- as.character(dat[5, 2, drop = TRUE])
      report_date   <- get_report_date(dat[4, 8, drop = TRUE])
      
      deposits = dat[12:25,c(2,3,7)] %>%
        set_names(c("description", "total", "interest_income_expense")) %>%
        mutate(category = "Deposits") %>%
        mutate(bank_name = bank_name
               , report_date = report_date
               , excel_sheet_code = sh
        )
      
    }
    
  })
}

cd_long_format_3 <- map_dfr(files_format_3, process_format_3)
}

### --- Format 4 (2007 03 to 2008 09) --- ###