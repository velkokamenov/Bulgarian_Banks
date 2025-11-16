library(httr)

years_cd    <- 2004:2025
quarters_cd <- c(3, 6, 9, 12)

dest_dir_cd <- "./Input Data/Credits and Deposits"
if (!dir.exists(dest_dir_cd)) dir.create(dest_dir_cd, recursive = TRUE)

base_url <- "https://www.bnb.bg/bnbweb/groups/public/documents/bnb_download"

# which codes (a2/a3/a4) exist for a given year / month
codes_for <- function(year, month) {
  # 2004–2006: no data
  if (year < 2007) return(character(0))
  
  # 2007 – 2008-09: a2, a3, a4 (bcb_* pattern)
  if (year < 2008 || (year == 2008 && month <= 9)) {
    return(c("a2", "a3", "a4"))
  }
  
  # 2008-12 – 2014-12: a2, a3, a4 (bs_q_*.xls)
  if (year <= 2014) {
    return(c("a2", "a3", "a4"))
  }
  
  # 2015 onwards: only a2
  "a2"
}

build_fname_cd <- function(year, month, code) {
  
  # 2007–2008-09 → bcb_* pattern, .xls
  if (year >= 2007 && (year < 2008 || (year == 2008 && month <= 9))) {
    mid <- switch(
      code,
      "a2" = "securities",
      "a3" = "loans_adv",
      "a4" = "attr_funds",
      stop("Unexpected code for 2007–2008-09: ", code)
    )
    return(sprintf("bcb_q_%s_%02d%d_%s_bg.xls", mid, month, year, code))
  }
  
  # 2008-12 – 2014-12 → bs_q_YYYYMM_a2/a3/a4_bg.xls
  if (year >= 2008 && year <= 2014) {
    return(sprintf("bs_q_%d%02d_%s_bg.xls", year, month, code))
  }
  
  # 2015 – 2021-06 → bs_q_YYYYMM_a2_bg.xls (only a2)
  # 2021-09 onwards → bs_q_YYYYMM_a2_bg.xlsx (only a2)
  if (year >= 2015) {
    ext <- if (year > 2021 || (year == 2021 && month >= 9)) "xlsx" else "xls"
    return(sprintf("bs_q_%d%02d_a2_bg.%s", year, month, ext))
  }
  
  stop("Year/month not covered in build_fname_cd: ", year, "-", month)
}

for (y in years_cd) {
  for (m in quarters_cd) {
    
    codes <- codes_for(y, m)
    if (length(codes) == 0) next
    
    for (code in codes) {
      
      fname_cd     <- build_fname_cd(y, m, code)
      file_url_cd  <- paste0(base_url, "/", fname_cd)
      dest_file_cd <- file.path(dest_dir_cd, fname_cd)
      
      if (file.exists(dest_file_cd)) next
      
      ok_cd <- FALSE
      try({
        resp_cd <- httr::HEAD(file_url_cd)
        ok_cd   <- (httr::status_code(resp_cd) == 200)
      }, silent = TRUE)
      
      if (!ok_cd) {
        message("Skip (Credits & Deposits): ", fname_cd)
        next
      }
      
      message("Downloading (Credits & Deposits): ", fname_cd)
      
      tryCatch({
        download.file(
          url      = file_url_cd,
          destfile = dest_file_cd,
          mode     = "wb",
          quiet    = TRUE
        )
      }, error = function(e) {
        message("Failed (Credits & Deposits): ", fname_cd, " - ", e$message)
      })
    }
  }
}

