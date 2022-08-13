#' Adds abbreviations to an excel file
#'
#' @param wb an openxlsx workbook object
#' @param sheet name of sheet in workbook object
#' @param abbr_text dataframe of text to display in abbreviations section of
#' workbook object. Include "MMMMM_#" to merge # number of cells to the right.
#' @param start_row number for starting row
#' @param start_col number for starting column
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
#' fmt_abbr(wb = wb, sheet = "sheet",
#' abbr_text = abbr_text,
#' start_row = 2, start_col = 1)
fmt_abbr <- function(wb, sheet, abbr_text, start_row = 1, start_col = 1){
  a_text <- abbr_text %>% dplyr::select(-table)
  # Add Text ---------------
  openxlsx::writeData(wb, sheet = sheet, x = a_text,
                      startCol = start_col,
                      startRow = start_row, colNames = F)
  rows <- dim(a_text)[1]

  # Merge Cells ------------------
  for(i in 1:dim(a_text)[1]){
    m <- which(grepl("MMMMM_", a_text[i,]) | grepl("BBBBB_", a_text[i,]))
    if(length(m) != 0){
      purrr::walk(m,
                  ~openxlsx::mergeCells(wb, sheet,
                                        cols = (.x - 1):
                                          (.x + as.numeric(gsub("MMMMM_|BBBBB_",
                                                                "", a_text[i, .x])) - 2),
                                        rows = (i + start_row - 1)))
    }
  }

  # Style ---------------------
  openxlsx::addStyle(wb, sheet, abbr_style(),
                     rows = start_row:(start_row - 1 + rows),
                     cols = start_col:dim(a_text)[2],
                     gridExpand = T, stack = T)

  return(list(wb = wb,
              length = start_row - 1 + rows))
}
