\name{Cell}
\alias{Cell}
\alias{createCell}
\alias{getCells}
\alias{setCellValue}
\alias{getCellValue}
\title{Functions to manipulate cells.}
\description{
Functions to manipulate cells. 
}
\usage{

createCell(row, colIndex=1:5)

getCells(row, simplify=TRUE)

setCellValue(cell, value, richTextString=FALSE)

getCellValue(cell, keepFormulas=FALSE)

}
\arguments{
  \item{row}{a list of row objects. See \code{Row}.}
  \item{colIndex}{a numeric vector specifying the index of columns to
    create.}
  \item{simplify}{a logical value.  If \code{TRUE}, the result will be
    unlisted.}
  \item{value}{an R variable of length one.}
  \item{richTextString}{a logical value indicating if the value should
    be inserted into the Excel cell as rich text.}
  \item{keepFormulas}{a logical value.  If \code{TRUE} the formulas will 
    be returned as characters instead of being explicitly evaluated.}
  \item{cell}{an object of type \code{Cell}.}
}

\value{
  
  \code{createCell} creates a matrix of lists, each element of the list
  being a java object reference to an object of type Cell representing
  an empty cell.  The dimnames of this matrix are taken from the names
  of the rows and the \code{colIndex} variable.
    
  \code{getCells} returns a list of java object references for all the
  cells in the row. 
  
  \code{getCellValue} returns the value in the cell as an R object.
  Type conversions are done behind the scene.  This function is not
  vectorized. 
}
\details{
  \code{setCellValue} writes the content of an R variable into the cell.
  \code{Date} and \code{POSIXct} objects are passed in as numerical
  values.  To format them as dates in Excel see \code{\link{CellStyle}}.
  This function is not vectorized. 
}

\author{Adrian Dragulescu}
\seealso{To format cells, see \code{\link{CellStyle}}.}
\examples{

file <- system.file("tests", "test_import.xlsx", package = "xlsx")

wb <- loadWorkbook(file)  
sheets <- getSheets(wb)

sheet <- sheets[[2]]      # get second sheet
rows  <- getRows(sheet)   # get all the rows

cells <- getCells(rows)   # returns all non empty cells

values <- lapply(cells, getCellValue) # extract the values

}

