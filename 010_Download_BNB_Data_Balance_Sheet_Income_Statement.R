library(httr)

years    <- 2004:2025
quarters <- c(3, 6, 9, 12)

dest_dir <- "./Input Data/Balance Sheet and Income Statement"
if (!dir.exists(dest_dir)) dir.create(dest_dir, recursive = TRUE)

base_url <- "https://www.bnb.bg/bnbweb/groups/public/documents/bnb_download"

build_fname <- function(year, month) {
  
  # Case 1: ≤ 2006  → old format
  if (year <= 2006) {
    return(sprintf("bcb_q_old_bi_%02d%d_a1_bg.xls", month, year))
  }
  
  # Case 2: 2007–2008-09 → income format
  if (year < 2008 || (year == 2008 && month <= 9)) {
    return(sprintf("bcb_q_income_%02d%d_a1_bg.xls", month, year))
  }
  
  # Case 3: 2008-12 to 2020 → bs_q + xls
  if (year <= 2020) {
    return(sprintf("bs_q_%d%02d_a1_bg.xls", year, month))
  }
  
  # Case 4: 2021+ → bs_q + xlsx
  sprintf("bs_q_%d%02d_a1_bg.xlsx", year, month)
}

for (y in years) {
  for (m in quarters) {
    
    fname     <- build_fname(y, m)
    file_url  <- paste0(base_url, "/", fname)
    dest_file <- file.path(dest_dir, fname)
    
    if (file.exists(dest_file)) next
    
    ok <- FALSE
    try({
      resp <- HEAD(file_url)
      ok   <- (status_code(resp) == 200)
    }, silent = TRUE)
    
    if (!ok) {
      message("Skip: ", fname)
      next
    }
    
    message("Downloading: ", fname)
    
    tryCatch({
      download.file(
        url      = file_url,
        destfile = dest_file,
        mode     = "wb",
        quiet    = TRUE
      )
    }, error = function(e) message("Failed: ", fname))
  }
}


