% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/excel_subscript.R
\name{excel_subscript}
\alias{excel_subscript}
\title{Adds subscripts to text in excel}
\usage{
excel_subscript(strings)
}
\arguments{
\item{strings}{sharedStrings from an openxlsx workbook object}
}
\value{
strings for an openxlsx workbook object
}
\description{
Adds subscripts to text in excel
}
\examples{
wb <- openxlsx::createWorkbook()

openxlsx::addWorksheet(wb, "sheet")

excel_subscript(strings = wb$sharedStrings)
}
