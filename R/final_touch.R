#' Applies page setup to excel table
#'
#' @param wb an openxlsx workbook object
#' @param sheet name of sheet in workbook object
#' @param title_length number of rows to include on each printed page
#' (typical all rows through the column headers)
#' @param start_row number for starting row
#' @param end_col number for the last column
#'
#' @return Updated excel workbook object.
#' @export
#'
#' @examples
#' wb <- openxlsx::createWorkbook()
#'
#' openxlsx::addWorksheet(wb, "sheet")
#'
#' final_touch(wb = wb, sheet = "sheet",
#' title_length = 5,
#' start_row = 2, end_col = 17)
final_touch <- function(wb, sheet, title_length, start_row = 1, end_col){
  # Adding blank columns to help with row sizing
  blank <- data.frame(V1 = replicate(start_row, "help_fix_height"))

  openxlsx::writeData(wb, sheet, blank,
                      startCol = end_col + 1,
                      startRow = start_row, colNames = F)

  # Adding super and sub scripts
  wb$sharedStrings <- excel_subscript(wb$sharedStrings)
  wb$sharedStrings <- excel_superscript(wb$sharedStrings)
  wb$sharedStrings <- excel_row_height(wb$sharedStrings)

  # Page Setup
  openxlsx::pageSetup(wb, sheet, orientation = "landscape",
                      fitToWidth = T, printTitleRows = 1:title_length)

  return(wb)
}
