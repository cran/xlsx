
# Run some simple automated tests for the package
#
#


require(RUnit); require(xlsx)

dirUT <- paste(.Library, "/xlsx/tests", sep="")

source(paste(dirUT, "/test.workbook.R", sep=""))
test.lowlevel()

source(paste(dirUT, "/test.workbook.R", sep=""))
test.workbook()

source(paste(dirUT, "/test.formats.R", sep=""))
test.formats()

source(paste(dirUT, "/test.import.R", sep=""))
test.import()











