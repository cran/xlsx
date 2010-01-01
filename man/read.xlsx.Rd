\name{read.xlsx}
\alias{read.xlsx}
\title{ Read the contents of a worksheet into an R data.frame}
\description{
Read the contents of a worksheet into an R \code{data.frame}.
}
\usage{
read.xlsx(file, sheetIndex, sheetName=NULL, rowIndex=NULL,
  as.data.frame=TRUE, header=TRUE, colClasses=NA, keepFormulas=FALSE)
}
\arguments{ 
  \item{file}{the \emph{absolute} path to the file which the data are to
    be read from.}
  \item{sheetIndex}{a number representing the sheet index in the workbook.}
  \item{sheetName}{a character string with the sheet name.}
  \item{rowIndex}{a numeric vector indicating the rows you want to
    extract.  If \code{NULL}, all the rows found will be extracted.}
  \item{as.data.frame}{a logical value indicating if the result should
    be coerced into a \code{data.frame}.  If \code{FALSE}, the result is
    a list with one element for each non empty cell.  The name for each
    element of this list is formed by pasting together the row index
    with the colum index separated by \emph{.}.}     
  \item{header}{a logical value indicating whether the first row
    corresponding to the first element of the \code{rowIndex} vector
    contains the names of the variables.}
  \item{colClasses}{a vector of classes to be assumed for the columns.
    Recycled as necessary, or if the character vector is named,
    unspecified values are taken to be \code{NA}.}
  \item{keepFormulas}{a logical value indicating if Excel formulas
    should be shown as text in \R and not evaluated before bringing them
    in.}
}

\details{
  
This function provides a high level API for reading data from an Excel
2007 worksheet.  It calls several low level functions in the process.
Its goal is to provide the conveniency of \code{\link[utils]{read.csv}}
by borrowing from its signature. 

The function pulls the value of each non empty cell in the worksheet
into a vector of type \code{list} by preserving the data type.  If
\code{as.data.frame=TRUE}, this vector of lists is then formatted into a
rectangular shape.  Special care is needed for worksheets with ragged
data.   

The class type of the variable corresponding to one column in the
worksheet is taken from the class of the first non empty cell in that
column.  If you need to impose a specific class type on a variable, use
the \code{colClasses} argument.  

Excel internally stores dates and datetimes as numeric values, and does
not keep track of time zones and DST.  When a datetime column is 
brought into \R, it is converted to \code{POSIXct} class with a
\emph{GMT} timezone.  Occasional rounding errors may appear and the \R
and Excel string representation my differ by one second.
   
}
\value{
  A data.frame containing the data as downloaded from the worksheet.  The
  rownames of the dataframe are the unique list id's used internally by
  Google.  You need these unique id's if you want to overwrite the rows
  of the spreadsheet. 
}
\author{ Adrian Dragulescu }
\seealso{\code{\link{write.xlsx}} for writing \code{xlsx} documents.
  See also \code{\link{saveWorkbook}} for reading a workbook into R for
  further processing and \code{\link{saveWorkbook}} for saving a
  workbook to file.}
\examples{
\dontrun{

file <- system.file("tests", "test_import.xlsx", package = "xlsx")
res <- read.xlsx(file, 1)  # read first sheet
head(res)
#          NA. Population Income Illiteracy Life.Exp Murder HS.Grad Frost   Area
# 1    Alabama       3615   3624        2.1    69.05   15.1    41.3    20  50708
# 2     Alaska        365   6315        1.5    69.31   11.3    66.7   152 566432
# 3    Arizona       2212   4530        1.8    70.55    7.8    58.1    15 113417
# 4   Arkansas       2110   3378        1.9    70.66   10.1    39.9    65  51945
# 5 California      21198   5114        1.1    71.71   10.3    62.6    20 156361
# 6   Colorado       2541   4884        0.7    72.06    6.8    63.9   166 103766
# >


}
}

