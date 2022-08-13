#' Adds a header to an excel file
#'
#' @param wb an openxlsx workbook object
#' @param sheet name of sheet in workbook object
#' @param header_text text string to display in header of workbook object
#' @param start_row number for starting row
#' @param start_col number for starting column
#' @param end_col number for the last column
#'
#' @return A list with the updated workbook object (wb) and number of the last
#' row in the updated workbook object (length)
#' @export
#'
#' @examples
#' wb <- openxlsx::createWorkbook()
#'
#' openxlsx::addWorksheet(wb, "sheet")
#'
#' fmt_header(wb = wb, sheet = "sheet",
#' header_text = "DRAFT - Privileged and Confidential",
#' start_row = 1, start_col = 1, end_col = 17)
fmt_header <- function(wb, sheet, header_text, start_row = 1, start_col = 1, end_col){

  # Add to Header to workbook object ------------
  openxlsx::writeData(wb, sheet = sheet, x = header_text,
                      startCol = start_col, startRow = start_row, colNames = F)

  # Merge Cells --------------
  openxlsx::mergeCells(wb, sheet, cols = 1:end_col, rows = start_row)

  # Add Style ------------
  openxlsx::addStyle(wb, sheet, header_style(), rows = start_row, cols = start_col, gridExpand = TRUE)

  return(list(wb = wb,
              length = start_row))
}
