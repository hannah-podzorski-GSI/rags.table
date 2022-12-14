% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/final_touch.R
\name{final_touch}
\alias{final_touch}
\title{Applies page setup to excel table}
\usage{
final_touch(wb, sheet, title_length, start_row = 1, end_col)
}
\arguments{
\item{wb}{an openxlsx workbook object}

\item{sheet}{name of sheet in workbook object}

\item{title_length}{number of rows to include on each printed page
(typical all rows through the column headers)}

\item{start_row}{number for starting row}

\item{end_col}{number for the last column}
}
\value{
Updated excel workbook object.
}
\description{
Applies page setup to excel table
}
\examples{
wb <- openxlsx::createWorkbook()

openxlsx::addWorksheet(wb, "sheet")

final_touch(wb = wb, sheet = "sheet",
title_length = 5,
start_row = 2, end_col = 17)
}
