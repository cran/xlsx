
# Run some simple automated tests for the package
#
#


require(RUnit); require(xlsx)

dirUT <- paste(.Library, "/xlsx/tests", sep="")
#dirUT <- "H:/user/R/Adrian/findataweb/temp/xlsx/trunk/inst/tests"

source(paste(dirUT, "/test.comments.R", sep=""))
test.comments()

source(paste(dirUT, "/test.dataFormats.R", sep=""))
test.dataFormats()

source(paste(dirUT, "/test.export.R", sep=""))
test.export()

source(paste(dirUT, "/test.formats.R", sep=""))
test.formats()

source(paste(dirUT, "/test.import.R", sep=""))
test.import()

#source(paste(dirUT, "/test.lowlevel.R", sep=""))
#test.lowlevel()

source(paste(dirUT, "/test.otherEffects.R", sep=""))
test.otherEffects()

source(paste(dirUT, "/test.picture.R", sep=""))
test.picture()

source(paste(dirUT, "/test.workbook.R", sep=""))
test.workbook()










