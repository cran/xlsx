# One sheet extraction.  Similar to read.csv. 
#
#
read.xlsx <- function(file, sheetIndex, sheetName=NULL, rowIndex=NULL,
  as.data.frame=TRUE, header=TRUE, colClasses=NA, keepFormulas=FALSE)
{
  if (is.null(sheetName) & missing(sheetIndex))
    stop("Please provide a sheet name OR a sheet index.")

  wb <- loadWorkbook(file)
  sheets <- getSheets(wb)

  if (is.null(sheetName)){
    sheet <- sheets[[sheetIndex]]
  } else {
    sheet <- sheets[[sheetName]]
  }
    
  rows  <- getRows(sheet, rowIndex)  
  cells <- getCells(rows)
  res <- lapply(cells, getCellValue, keepFormulas=keepFormulas)

  if (as.data.frame){
    ind <- lapply(strsplit(names(res), "\\."), as.numeric)
    namesIndM <- do.call(rbind, ind)
    
    row.names <- sort(as.numeric(unique(namesIndM[,1])))
    col.names <- paste("V", sort(unique(namesIndM[,2])), sep="")  
    cols <- length(col.names)

    VV <- matrix(list(NA), nrow=length(row.names), ncol=cols,
      dimnames=list(NULL, col.names))
    VV[namesIndM] <- res
    
    if (header){  # first row of cells that you want
      colnames(VV) <- VV[1,]
      VV <- VV[-1,]
    }
    
    res <- vector("list", length=cols)
    names(res) <- colnames(VV)
    for (ic in seq_len(cols)){
      aux <- unlist(VV[,ic], use.names=FALSE)
      ind <- min(which(!is.na(aux)))  
      if (class(aux[ind])=="numeric"){
        # test first not NA cell if it's a date/datetime
        dateUtil <- .jnew("org/apache/poi/ss/usermodel/DateUtil")
        cell <- cells[[paste(ind+header, ".", ic, sep="")]]
        isDatetime <- dateUtil$isCellDateFormatted(cell)
        if (isDatetime){
          if (identical(aux, round(aux))){ # you have dates
            aux <- as.Date(aux-25569, origin="1970-01-01") 
          } else {  # Excel does not know timezones?!
            aux <- as.POSIXct((aux-25569)*86400, tz="GMT", origin="1970-01-01")
          }
        }
      }
      if (!is.na(colClasses[ic]))
        class(aux) <- colClasses[ic]  # if it gets specified
      res[[ic]] <- aux
    }

    res <- data.frame(res)
  }

  res
}




# old function, for a quick and dirty read.xls
##     # probably should not have it!
##     extractor <- .jnew("org/apache/poi/xssf/extractor/XSSFExcelExtractor", file)
##     txt <- .jcall(extractor, "S", "getText")
##     #TODO: release the extractor!!!  
    
##     txt <- strsplit(txt, "\n")[[1]]
##     aux <- strsplit(txt, "\t")
##     ncols <- sapply(aux, length)  # empty cells are skipped!
##     indSheets <- which(ncols==1)  # sheet feeds start with sheet name

##     # extract the sheet you want
##     ind <- grep(paste("\\b", sheet, "\\b", sep=""), aux)
##     if (length(ind)==1){  # found only one entry
##       bux <- aux[(ind[1]+1):(indSheets[min(which(indSheets > ind[1]))]-1)]
##     }
##     if (length(bux)==0){
##       warning("Something went wrong!")
##       res <- txt
##     } else {
##       res <- do.call(rbind, bux[-1])
##       colnames(res) <- bux[[1]]      
##     }
 
