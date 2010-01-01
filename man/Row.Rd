\name{Row}
\alias{Row}
\alias{createRow}
\alias{getRows}
\title{Functions to manipulate rows of a worksheet.}
\description{
Functions to manipulate rows of a worksheet. 
}
\usage{

createRow(sheet, rowIndex=1:5)

getRows(sheet, rowIndex=NULL)

}
\arguments{
  \item{sheet}{a workbook object as returned by \code{createSheet} or
    subsetting \code{getSheets}.}
  \item{rowIndex}{a numeric vector specifying the index of rows to
    create.  For \code{getRows}, a \code{NULL} value will return all
    non empty rows.}
}

\value{
  A list of java object references each  pointing to a row.  The list is
  named with the row number. 
}
\author{Adrian Dragulescu}
\seealso{To extract the cells from a given row, see \code{\link{Cell}}.}
\examples{

file <- system.file("tests", "test_import.xlsx", package = "xlsx")

wb <- loadWorkbook(file)  
sheets <- getSheets(wb)

sheet <- sheets[[2]]  
rows  <- getRows(sheet)  # get all the rows

# see all the available java methods that you can call
.jmethods(rows[[1]])

# for example
rows[[1]]$getRowNum()   # zero based index in Java


}

