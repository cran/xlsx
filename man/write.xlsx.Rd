\name{write.xlsx}
\alias{write.xlsx}
\title{Write a data.frame to an Excel 2007 workbook.}
\description{
Write a \code{data.frame} to an Excel 2007 workbook.
}
\usage{
write.xlsx(x, file, sheetName="Sheet 1", formatTemplate=NULL,
  col.names=TRUE, row.names=TRUE)
}
\arguments{
  \item{x}{a \code{data.frame} to write to the workbook.}
  \item{file}{the \emph{absolute} path to the file which the data are to
    be read from.}
  \item{sheetName}{a character string with the sheet name.}
  \item{formatTemplate}{Not implemented now.  A way for specifying
    the formatting of the table.} 
  \item{col.names}{a logical value indicating if the column names of
    \code{x} are to be written along with \code{x} to the file.}
  \item{row.names}{a logical value indicating whether the row names of 
    \code{x} are to be written along with \code{x} to the file.}
}

\details{
  
This function provides a high level API for writing a \code{data.frame}
to an Excel 2007 worksheet.  It calls several low level functions in the
process.  Its goal is to provide the conveniency of
\code{\link[utils]{write.csv}} by borrowing from its signature.

Internally, there is a double loop in \R over all the elements of the
\code{data.frame} so performance for very large \code{data.frame} may be
an issue.  Please report if you experience slow performance.  Dates and
POSIXct classes are formatted separately after the insertion.  This also
adds to processing time.

}
\author{ Adrian Dragulescu }
\seealso{\code{\link{read.xlsx}} for reading \code{xlsx} documents.
  See also \code{\link{saveWorkbook}} for reading a workbook into R for
  further processing and \code{\link{saveWorkbook}} for saving a
  workbook to file.}
\examples{
\dontrun{

file <- paste(tempdir(), "/usarrests.xlsx", sep="")
res <- write.xlsx(USArrests, file)  

}
}

