% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/fmt_meta.R
\name{fmt_meta}
\alias{fmt_meta}
\title{Adds Meta information to an excel file}
\usage{
fmt_meta(
  wb,
  sheet,
  meta_text,
  d,
  start_row = 1,
  start_col = 1,
  meta_end_col = 3
)
}
\arguments{
\item{wb}{an openxlsx workbook object}

\item{sheet}{name of sheet in workbook object}

\item{meta_text}{dataframe of text to display in meta info of workbook object}

\item{d}{dataframe of data to be included in table}

\item{start_row}{number for starting row}

\item{start_col}{number for starting column}

\item{meta_end_col}{number for the last column of meta info box}
}
\value{
A list with the updated workbook object (wb) and number of the last
row in the updated workbook object (length)
}
\description{
Adds Meta information to an excel file
}
\examples{
wb <- openxlsx::createWorkbook()

openxlsx::addWorksheet(wb, "sheet")

fmt_meta(wb = wb, sheet = "sheet",
meta_text = meta_text, d = sample_data,
start_row = 2, start_col = 1, meta_end_col = 3)
}
