## ----setup, echo=FALSE, message=FALSE------------------------------------
knitr::opts_chunk$set(echo=TRUE, collapse=T, comment='#>')
library(RefManageR)
bib <- ReadBib('xlsx.bib')
BibOptions(check.entries = FALSE, style = "markdown", cite.style = "numeric",bib.style = "numeric")

## ----results = "asis", echo = FALSE--------------------------------------
PrintBibliography(bib, .opts = list(check.entries = FALSE, sorting = "ynt"))

