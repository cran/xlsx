#
#
#

test.export <- function(outdir="C:/Temp/")
{
  require(xlsx)

  cat("####################################################\n")
  cat("Test high level export ... \n")
  cat("####################################################\n")
      

  x <- data.frame(mon=month.abb[1:10], day=1:10, year=2000:2009,
    date=seq(as.Date("2009-01-01"), by="1 month", length.out=10),
    bool=ifelse(1:10 %% 2, TRUE, FALSE))

  file <- paste(outdir, "test_export.xlsx", sep="")
  cat("Write an xlsx file with char, int, double, date, bool columns ...\n")
  write.xlsx(x, file)
  cat("Wrote file ", file, "\n\n")

  cat("Test the append argument by adding another sheet \n")
  file <- paste(outdir, "test_export.xlsx", sep="")
  write.xlsx(USArrests, file, sheetName="usarrests", append=TRUE)
  cat("Wrote file ", file, "\n\n")
  
  
}










