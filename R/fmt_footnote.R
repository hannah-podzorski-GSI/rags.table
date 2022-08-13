#' Adds footnotes to an excel file
#'
#' @param wb an openxlsx workbook object
#' @param sheet name of sheet in workbook object
#' @param footnote_text dataframe of text to display in footnote section of
#' workbook object
#' @param start_row number for starting row
#' @param start_col number for starting column
#' @param end_col number for the last column
#'
#' @return A list with the updated workbook object (wb) and number of the last
#' row in the updated workbook object (length).
#' @export
#'
#' @examples
#' wb <- openxlsx::createWorkbook()
#'
#' openxlsx::addWorksheet(wb, "sheet")
#'
#' fmt_footnote(wb = wb, sheet = "sheet",
#' footnote_text = foot_text,
#' start_row = 2, start_col = 1, end_col = 17)
fmt_footnote <- function(wb, sheet, footnote_text, start_row = 1, start_col = 1, end_col){
  # Add Text ---------------
  openxlsx::writeData(wb, sheet = sheet, x = "Footnotes",
                      startCol = start_col, startRow = start_row,
                      colNames = F)

  openxlsx::writeData(wb, sheet = sheet, x = footnote_text$text,
                      startCol = start_col, startRow = start_row + 1,
                      colNames = F)

  rows <- dim(footnote_text)[1]

  # Merge Cells -------------
  purrr::walk(start_row:(start_row + rows),
              ~openxlsx::mergeCells(wb, sheet,
                                    cols = start_col:end_col, rows = .x))

  # Style --------------------
  openxlsx::addStyle(wb, sheet, footer_style_bold(),
                     rows = start_row,
                     cols = start_col,
                     gridExpand = T, stack = T)

  openxlsx::addStyle(wb, sheet, footer_style(),
                     rows = start_row :(start_row + rows),
                     cols = start_col,
                     gridExpand = T, stack = T)

  return(list(wb = wb,
              length = start_row + rows))

}
