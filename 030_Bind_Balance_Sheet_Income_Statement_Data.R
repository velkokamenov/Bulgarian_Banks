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
in_dir  <- "./Input Data/Balance Sheet and Income Statement"

### --- Format 1 (2004 03 to 2006 12) --- ###
{
files_format_1 <- list.files(
  path       = in_dir,
  pattern    = "^bcb_q_old_bi_[0-9]{2}(2004|2005|2006)_a1_bg\\.xls$",
  full.names = TRUE
)

# process one workbook (all sheets) ----------------------------------------
process_format_1 <- function(path) {
  
  library(readxl)
  library(tidyverse)
  library(lubridate)
  
  sheets <- excel_sheets(path) |> setdiff(c("Sheet1", "Title"))
  
  map_dfr(sheets, function(sh) {
    
    dat <- read_xls(path, sheet = sh, col_names = FALSE)
    
    if (is.na(dat[1,1])) {
      dat <- dat[-1, ]
    }
    
    bank_name     <- as.character(dat[1, 2, drop = TRUE])
    
    year  <- as.integer(dat[1,5])
    month <- as.integer(dat[1,4])
    
    print(paste0(year,"-",month,"-",sh))
    
    report_date <- ymd(sprintf("%04d-%02d-01", year, month)) %>%
      ceiling_date("month") - days(1)
    
    if (!sh %in% c('190','199','145','250','350','898')) {
    
    assets <- dat[7:28, c(2,4)] %>%
      set_names(c("description", "value")) %>%
      mutate(category = "Assets")
    
    liabilities_equity = dat[30:51,c(2,4)] %>%
      set_names(c("description", "value")) %>%
      mutate(category = "Liabilities & Equity")
    
    income_statement = dat[57:86,c(2,4)] %>%
      set_names(c("description", "value")) %>%
      mutate(category = "Income Statement")
    
    all_data = assets %>%
      bind_rows(liabilities_equity) %>%
      bind_rows(income_statement) %>%
      mutate(bank_name = bank_name
             , report_date = report_date
             , excel_sheet_code = sh
      )
    
    } else
    {
      
      assets <- dat[7:27, c(2,4)] %>%
        set_names(c("description", "value")) %>%
        mutate(category = "Assets")
      
      liabilities_equity = dat[29:47,c(2,4)] %>%
        set_names(c("description", "value")) %>%
        mutate(category = "Liabilities & Equity")
      
      income_statement = dat[53:80,c(2,4)] %>%
        set_names(c("description", "value")) %>%
        mutate(category = "Income Statement")
      
      all_data = assets %>%
        bind_rows(liabilities_equity) %>%
        bind_rows(income_statement) %>%
        mutate(bank_name = bank_name
               , report_date = report_date
               , excel_sheet_code = sh
        )
      
    }
    
  })
}

bs_long_format_1 <- map_dfr(files_format_1, process_format_1)
}

### --- Format 2 (2007 03 to 2014 12) --- ###
{
files_format_2 <- list.files(
  path       = in_dir,
  pattern    = "^(bcb_q_income_|bs_q_((2008(1[2]))|(2009(0[1-9]|1[0-2]))|(201[0-4](0[1-9]|1[0-2]))))",
  full.names = TRUE
)

# process one workbook (all sheets) ----------------------------------------
process_format_2 <- function(path) {
  
  library(readxl)
  library(tidyverse)
  library(lubridate)
  
  sheets <- excel_sheets(path) |> setdiff(c("Sheet1", "Title"))
  
  map_dfr(sheets, function(sh) {
    
    dat <- read_xls(path, sheet = sh, col_names = FALSE, .name_repair = "minimal")
    
    if (dat[5,1] == "Шифър" && sh %in% c('660') && as.character(dat[1, 2, drop = TRUE]) != "ЕЙЧ ВИ БИ БАНК БИОХИМ") {
      
      bank_name <- as.character(dat[1, 2, drop = TRUE])
      
      year  <- as.integer(dat[2,3])
      month <- as.integer(dat[2,2])
      
    } else if (dat[5,1] == "Шифър") {
      
      bank_name <- as.character(dat[1, 2, drop = TRUE])
      
      year  <- as.integer(dat[2,4])
      month <- as.integer(dat[2,3])
      
    } else if (is.na(dat[2,2])) {
      
      bank_name <- as.character(dat[1, 1, drop = TRUE])
      
      year <- str_extract(dat[2,3], "(?<=\\.)[0-9]{4}") |> as.integer()
      month <- str_extract(dat[2,3], "^[0-9]{1,2}") |> as.integer()
      
      
    } else if (sh %in% c('660')) {
      
      bank_name <- as.character(dat[1, 1, drop = TRUE])
      
      year  <- as.integer(dat[2,2])
      month <- as.integer(dat[2,1])
      
    } else {
      
      bank_name <- as.character(dat[1, 1, drop = TRUE])
      
      year  <- as.integer(dat[2,3])
      month <- as.integer(dat[2,2])
      
    }
    
    report_date <- ymd(sprintf("%04d-%02d-01", year, month)) %>%
      ceiling_date("month") - days(1)
    
    if (year <= 2007) {
    
    assets <- dat[6:20, c(2,3)] %>%
      set_names(c("description", "value")) %>%
      mutate(category = "Assets")
    
    liabilities = dat[23:35,c(2,3)] %>%
      set_names(c("description", "value")) %>%
      mutate(category = "Liabilities")
      
    equity = dat[38:48,c(2,3)] %>%
      set_names(c("description", "value")) %>%
      mutate(category = "Equity")
      
    income_statement = dat[52:80,c(2,3)] %>%
      set_names(c("description", "value")) %>%
      mutate(category = "Income Statement")
    
    } else if (dat[5,1] == "Шифър") {
      
      assets <- dat[6:20, c(2,3)] %>%
        set_names(c("description", "value")) %>%
        mutate(category = "Assets")
      
      liabilities = dat[23:35,c(2,3)] %>%
        set_names(c("description", "value")) %>%
        mutate(category = "Liabilities")
      
      equity = dat[38:48,c(2,3)] %>%
        set_names(c("description", "value")) %>%
        mutate(category = "Equity")
      
      income_statement = dat[52:80,c(2,3)] %>%
        set_names(c("description", "value")) %>%
        mutate(category = "Income Statement") 
      
    } else {
      
      assets <- dat[6:20, c(1,2)] %>%
        set_names(c("description", "value")) %>%
        mutate(category = "Assets")
      
      liabilities = dat[23:35,c(1,2)] %>%
        set_names(c("description", "value")) %>%
        mutate(category = "Liabilities")
      
      equity = dat[38:48,c(1,2)] %>%
        set_names(c("description", "value")) %>%
        mutate(category = "Equity")
      
      income_statement = dat[52:80,c(1,2)] %>%
        set_names(c("description", "value")) %>%
        mutate(category = "Income Statement") 
      
    }
      
    all_data = assets %>%
      bind_rows(liabilities) %>%
      bind_rows(equity) %>%
      bind_rows(income_statement) %>%
      mutate(bank_name = bank_name
               , report_date = report_date
               , excel_sheet_code = sh
        )
 
  })
}

bs_long_format_2 <- map_dfr(files_format_2, process_format_2)
}

### --- Format 3 (2015 03 to 2020 03) --- ###
{
files_format_3 <- list.files(
  path       = in_dir,
  pattern    = "^bs_q_((2015(0[3-9]|1[0-2]))|(201[6-9](0[1-9]|1[0-2]))|(2020(0[1-3])))_a1_bg\\.xls$",
  full.names = TRUE
)

# process one workbook (all sheets) ----------------------------------------
process_format_3 <- function(path) {
  
  library(readxl)
  library(tidyverse)
  library(lubridate)
  
  sheets <- excel_sheets(path) |> setdiff(c("Sheet1", "Title"))
  
  map_dfr(sheets, function(sh) {
    
    dat <- read_xls(path, sheet = sh, col_names = FALSE, .name_repair = "minimal")
    
    bank_name <- as.character(dat[1, 1, drop = TRUE])
    report_date   <- get_report_date(dat[2, 3, drop = TRUE])
      
    assets <- dat[9:23, c(1,2,3)] %>%
      set_names(c("code","description", "value")) %>%
      mutate(category = "Assets")
    
    liabilities = dat[29:39,c(1,2,3)] %>%
      set_names(c("code","description", "value")) %>%
      mutate(category = "Liabilities")
    
    equity = dat[45:58,c(1,2,3)] %>%
      set_names(c("code","description", "value")) %>%
      mutate(category = "Equity")
    
    income_statement = dat[64:94,c(1,2,3)] %>%
      set_names(c("code","description", "value")) %>%
      mutate(category = "Income Statement")
    
    all_data = assets %>%
      bind_rows(liabilities) %>%
      bind_rows(equity) %>%
      bind_rows(income_statement) %>%
      mutate(bank_name = bank_name
             , report_date = report_date
             , excel_sheet_code = sh
      )
    
  })
}

bs_long_format_3 <- map_dfr(files_format_3, process_format_3)
}

### --- Format 4 (2020 06 to 2021 06) --- ###
{
files_format_4 <- list.files(
  path       = in_dir,
  pattern    = "^bs_q_((2020(0[6-9]|1[0-2]))|(2021(0[1-6])))_a1_bg\\.xls$",
  full.names = TRUE
)

# process one workbook (all sheets) ----------------------------------------
process_format_4 <- function(path) {
  
  library(readxl)
  library(tidyverse)
  library(lubridate)
  
  sheets <- excel_sheets(path) |> setdiff(c("Sheet1", "Title"))
  
  map_dfr(sheets, function(sh) {
    
    dat <- read_xls(path, sheet = sh, col_names = FALSE, .name_repair = "minimal")
    
    bank_name <- as.character(dat[1, 1, drop = TRUE])
    report_date   <- get_report_date(dat[2, 3, drop = TRUE])
    
    assets <- dat[9:23, c(1,2,3)] %>%
      set_names(c("code","description", "value")) %>%
      mutate(category = "Assets")
    
    liabilities = dat[29:39,c(1,2,3)] %>%
      set_names(c("code","description", "value")) %>%
      mutate(category = "Liabilities")
    
    equity = dat[45:58,c(1,2,3)] %>%
      set_names(c("code","description", "value")) %>%
      mutate(category = "Equity")
    
    if (path == "./Input Data/Balance Sheet and Income Statement/bs_q_202106_a1_bg.xls") {
    
    income_statement = dat[64:96,c(1,2,3)] %>%
      set_names(c("code","description", "value")) %>%
      mutate(category = "Income Statement")
    
    } else {
      
      income_statement = dat[64:95,c(1,2,3)] %>%
        set_names(c("code","description", "value")) %>%
        mutate(category = "Income Statement")
      
    }
    
    all_data = assets %>%
      bind_rows(liabilities) %>%
      bind_rows(equity) %>%
      bind_rows(income_statement) %>%
      mutate(bank_name = bank_name
             , report_date = report_date
             , excel_sheet_code = sh
      )
    
  })
}

bs_long_format_4 <- map_dfr(files_format_4, process_format_4)
}

### --- Format 5 (2021 09 to Latest) --- ###
{
files_format_5 <- list.files(
  path       = in_dir,
  pattern    = "^bs_q_[0-9]{6}_a1_bg\\.xlsx$",
  full.names = TRUE
)

# process one workbook (all sheets) ----------------------------------------
process_format_5 <- function(path) {
  
  library(readxl)
  library(tidyverse)
  
  sheets <- readxl::excel_sheets(path)
  
  map_dfr(sheets, function(sh) {
    dat <- read_xlsx(path, sheet = sh, col_names = FALSE)
    
    bank_name     <- as.character(dat[1, 1, drop = TRUE])
    report_date   <- get_report_date(dat[2, 3, drop = TRUE])
    
    assets = dat[9:23,] %>%
      set_names(c("code", "description", "value")) %>%
      mutate(category = "Assets")
    
    liabilities = dat[29:39,] %>%
      set_names(c("code", "description", "value")) %>%
      mutate(category = "Liabilities")
    
    equity = dat[45:58,] %>%
      set_names(c("code", "description", "value")) %>%
      mutate(category = "Equity")
    
    income_statement = dat[64:96,] %>%
      set_names(c("code", "description", "value")) %>%
      mutate(category = "Income Statement")
    
    all_data = assets %>%
      bind_rows(liabilities) %>%
      bind_rows(equity) %>%
      bind_rows(income_statement) %>%
      mutate(bank_name = bank_name
             , report_date = report_date
             , excel_sheet_code = sh
      )
    
    
  })
}

bs_long_format_5 <- map_dfr(files_format_5, process_format_5)
  
}

# Format Final File
Bank_Names_Nomenclature = read_xlsx("./Input Data/Bank_Names_Nomenclature.xlsx")

All_Balance_Sheet_Income_Data = bs_long_format_1 %>%
  bind_rows(bs_long_format_2) %>%
  bind_rows(bs_long_format_3) %>%
  bind_rows(bs_long_format_4) %>%
  bind_rows(bs_long_format_5) %>%
  left_join(Bank_Names_Nomenclature %>% mutate(excel_sheet_code = as.character(excel_sheet_code)), by = "excel_sheet_code") %>%
  mutate(description = toupper(description)
         , description = case_when(description %in% c("ОБЩО КАПИТАЛ","ОБЩО СОБСТВЕН КАПИТАЛ") ~ "ОБЩО СОБСТВЕН КАПИТАЛ"
                                   , description %in% c("ОБЩО ПАСИВИ И КАПИТАЛ","ОБЩО СОБСТВЕН КАПИТАЛ И ОБЩО ПАСИВИ","ОБЩО ПАСИВИ, МАЛЦИНСТВЕНО УЧАСТИЕ И КАПИТАЛ") ~ "ОБЩО СОБСТВЕН КАПИТАЛ И ОБЩО ПАСИВИ"
                                   , description %in% c("ПРИХОДИ ОТ ЛИХВИ","ПРИХОД ОТ ЛИХВИ","ПРИХОДИ ОТ ЛИХВИ И ДИВИДЕНТИ") ~ "ПРИХОДИ ОТ ЛИХВИ"
                                   , description %in% c("ПРИХОДИ ОТ ТАКСИ И КОМИСИОННИ","ПРИХОДИ ОТ ТАКСИ И КОМИСИОНИ") ~ "ПРИХОДИ ОТ ТАКСИ И КОМИСИОНИ"
                                   , description %in% c("ПЕЧАЛБА ИЛИ (-) ЗАГУБА ЗА ГОДИНАТА","ПЕЧАЛБА ИЛИ ЗАГУБА, ПРИНАДЛЕЖАЩА НА АКЦИОНЕРИТЕ НА МАЙКАТА","НЕТНА ПЕЧАЛБА/ЗАГУБА") ~ "ПЕЧАЛБА ИЛИ (-) ЗАГУБА ЗА ГОДИНАТА"
                                   , description %in% c("(РАЗХОДИ ЗА ЛИХВИ)","РАЗХОД ЗА ЛИХВИ","РАЗХОДИ ЗА ЛИХВИ","(ЛИХВЕНИ РАЗХОДИ)") ~ "(РАЗХОДИ ЗА ЛИХВИ)"
                                   , description %in% c("(РАЗХОДИ ЗА ТАКСИ И КОМИСИОНИ)","РАЗХОДИ ЗА ТАКСИ И КОМИСИОННИ") ~ "(РАЗХОДИ ЗА ТАКСИ И КОМИСИОНИ)"
                                   , description %in% c("(АДМИНИСТРАТИВНИ РАЗХОДИ)","АДМИНИСТРАТИВНИ РАЗХОДИ") ~ "(АДМИНИСТРАТИВНИ РАЗХОДИ)"
                                   , description %in% c("(ОБЕЗЦЕНКА ИЛИ (-) ОБРАТНО ВЪЗСТАНОВЯВАНЕ НА ОБЕЗЦЕНКА НА ФИНАНСОВИ АКТИВИ, КОИТО НЕ СЕ ОТЧИТАТ ПО СПРАВЕДЛИВА СТОЙНОСТ В ПЕЧАЛБАТА ИЛИ ЗАГУБАТА)"
                                                        ,"(ОБЕЗЦЕНКА ИЛИ (-) ОБРАТНО ВЪЗСТАНОВЯВАНЕ НА ОБЕЗЦЕНКА НА ИНВЕСТИЦИИ В ДЪЩЕРНИ ДРУЖЕСТВА, СЪВМЕСТНИ ПРЕДПРИЯТИЯ И АСОЦИИРАНИ ПРЕДПРИЯТИЯ)"
                                                        , "(ОБЕЗЦЕНКА ИЛИ (-) ОБРАТНО ВЪЗСТАНОВЯВАНЕ НА ОБЕЗЦЕНКА НА НЕФИНАНСОВИ АКТИВИ)"
                                                        , "(ОБЕЗЦЕНКА ИЛИ (-) КОРЕКЦИЯ НА ОБЕЗЦЕНКА НА ФИНАНСОВИ АКТИВИ, КОИТО НЕ СЕ ОТЧИТАТ ПО СПРАВЕДЛИВА СТОЙНОСТ В ПЕЧАЛБАТА ИЛИ ЗАГУБАТА)"
                                                        , "(ОБЕЗЦЕНКА ИЛИ (-) СТОРНИРАНЕ НА ОБЕЗЦЕНКА НА ИНВЕСТИЦИИ В ДЪЩЕРНИ ПРЕДПРИЯТИЯ, СЪВМЕСТНИ ПРЕДПРИЯТИЯ И АСОЦИИРАНИ ПРЕДПРИЯТИЯ)"
                                                        , "(ОБЕЗЦЕНКА ИЛИ (-) СТОРНИРАНЕ НА ОБЕЗЦЕНКА НА НЕФИНАНСОВИ АКТИВИ)"
                                                        , "ОБЕЗЦЕНКА"
                                                        , "(ОБЕЗЦЕНКА ИЛИ (-) ОБРАТНО ВЪЗСТАНОВЕНА ОБЕЗЦЕНКА НА ФИНАНСОВИ АКТИВИ, КОИТО НЕ СЕ ОТЧИТАТ ПО СПРАВЕДЛИВА СТОЙНОСТ В ПЕЧАЛБАТА ИЛИ ЗАГУБАТА)"
                                                        , "(ОБЕЗЦЕНКА ИЛИ (-) ОБРАТНО ВЪЗСТАНОВЕНА ОБЕЗЦЕНКА НА ИНВЕСТИЦИИ В ДЪЩЕРНИ ПРЕДПРИЯТИЯ, СЪВМЕСТНИ ПРЕДПРИЯТИЯ И АСОЦИИРАНИ ПРЕДПРИЯТИЯ)"
                                                        , "(ОБЕЗЦЕНКА ИЛИ (-) ОБРАТНО ВЪЗСТАНОВЕНА ОБЕЗЦЕНКА НА НЕФИНАНСОВИ АКТИВИ)"
                                                        , "НЕТНИ КРЕДИТНИ ПРОВИЗИИ") ~ "ОБЕЗЦЕНКА"
                                   , description %in% c("ОБЩО НЕТЕН ОПЕРАТИВЕН ДОХОД","НЕТЕН ОБЩ ОПЕРАТИВЕН ПРИХОД") ~ "ОБЩО НЕТЕН ОПЕРАТИВЕН ДОХОД"
                                   , description %in% c("ПЕЧАЛБА ИЛИ (-) ЗАГУБА ПРЕДИ ДАНЪЧНО ОБЛАГАНЕ ОТ ТЕКУЩИ ДЕЙНОСТИ","ОБЩО ПЕЧАЛБА ИЛИ ЗАГУБА ОТ ПРОДЪЛЖАВАЩИ (НЕПРЕУСТАНОВЕНИ) ДЕЙНОСТИ ПРЕДИ ДАНЪЦИ","ПЕЧАЛБА/ЗАГУБА ПРЕДИ ВАЛУТНА ПРЕОЦ., ИЗВЪНРЕДНИ ПР./РАЗХ.И ДАНЪЦИ") ~ "ПЕЧАЛБА ИЛИ (-) ЗАГУБА ПРЕДИ ДАНЪЧНО ОБЛАГАНЕ ОТ ТЕКУЩИ ДЕЙНОСТИ"
                                   , description %in% c("(ДАНЪЧНИ РАЗХОДИ ИЛИ (-) ПРИХОДИ, СВЪРЗАНИ С ПЕЧАЛБАТА ИЛИ ЗАГУБАТА ОТ ТЕКУЩИ ДЕЙНОСТИ)","ДАНЪЧЕН РАЗХОД (ПРИХОД) СВЪРЗАН С ПЕЧАЛБАТА ИЛИ ЗАГУБАТА ОТ  ПРОДЪЛЖАВАЩИ (НЕПРЕУСТАНОВЕНИ) ДЕЙНОСТИ","ДАНЪЦИ") ~ "(ДАНЪЧНИ РАЗХОДИ ИЛИ (-) ПРИХОДИ, СВЪРЗАНИ С ПЕЧАЛБАТА ИЛИ ЗАГУБАТА ОТ ТЕКУЩИ ДЕЙНОСТИ)"
                                   , description %in% c("(ДРУГИ ОПЕРАТИВНИ РАЗХОДИ)","ДРУГИ ОПЕРАТИВНИ РАЗХОДИ") ~ "(ДРУГИ ОПЕРАТИВНИ РАЗХОДИ)"
                                   , description %in% c("НЕТНИ ПЕЧАЛБИ ИЛИ (-) ЗАГУБИ ОТ КУРСОВИ РАЗЛИКИ","НЕТНИ КУРСОВИ РАЗЛИКИ [ПЕЧАЛБА (-) ЗАГУБА]","НЕТНИ ВАЛУТНИ РАЗЛИКИ","ПЕЧАЛБА/ЗАГУБА  ОТ ВАЛУТНА ПРЕОЦЕНКА") ~ "НЕТНИ ПЕЧАЛБИ ИЛИ (-) ЗАГУБИ ОТ КУРСОВИ РАЗЛИКИ"
                                   , description %in% c("НЕТНИ ПЕЧАЛБИ ИЛИ (-) ЗАГУБИ ОТ ФИНАНСОВИ АКТИВИ И ПАСИВИ, ДЪРЖАНИ ЗА ТЪРГУВАНЕ","НЕТНИ ПЕЧАЛБИ (ЗАГУБИ) ОТ ФИНАНСОВИ АКТИВИ И ПАСИВИ ДЪРЖАНИ ЗА ТЪРГУВАНЕ") ~ "НЕТНИ ПЕЧАЛБИ ИЛИ (-) ЗАГУБИ ОТ ФИНАНСОВИ АКТИВИ И ПАСИВИ, ДЪРЖАНИ ЗА ТЪРГУВАНЕ"
                                   , TRUE ~ description
                                   )
         )

write_xlsx(All_Balance_Sheet_Income_Data,"./Output Data/030_All_Balance_Sheet_Income_Data.xlsx")



