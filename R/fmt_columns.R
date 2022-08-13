#' Adds column headers to an excel file
#'
#' @param wb an openxlsx workbook object
#' @param sheet name of sheet in workbook object
#' @param col_names dataframe of text to display in column headers of
#' workbook object
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
#' fmt_columns(wb = wb, sheet = "sheet",
#' col_names = col_names,
#' start_row = 2, start_col = 1, end_col = 17)
fmt_columns <- function(wb, sheet, col_names, start_row = 1, start_col = 1,
                        end_col){
  # First row of col_names should be numeric width of columns
  c_names <- col_names %>% dplyr::select(-table)
  col_widths <- c_names[1,]
  c_names <- c_names[-1,]

  # Adding Column Widths ----------------
  openxlsx::setColWidths(wb, sheet = sheet,
                         cols = start_col:(end_col + 1),
                         widths = c(col_widths, 1))

  rows <- length(c_names$col_1)

  # Add Text --------------------------------
  openxlsx::writeData(wb, sheet = sheet, x = c_names,
                      startCol = start_col, startRow = start_row,
                      colNames = F,
                      borders = c("surrounding"),
                      borderStyle = getOption("openxlsx.borderStyle", "thin"))

  # Merge Cells --------------------------
  for(i in 1:dim(c_names)[1]){
    m <- which(grepl("MMMMM_", c_names[i,]) | grepl("BBBBB_", c_names[i,]))
    if(length(m) != 0){
      purrr::walk(m,
                  ~openxlsx::mergeCells(wb, sheet,
                                        cols = (.x - 1):
                                          (.x + as.numeric(gsub("MMMMM_|BBBBB_",
                                                                "", c_names[i, .x])) - 2),
                                        rows = (i + start_row - 1)))
    }
  }

  # Style -----------------------
  openxlsx::addStyle(wb, sheet = sheet, title_style(),
                     rows = start_row:(start_row - 1 + rows),
                     cols = start_col:end_col, gridExpand = T, stack = T)

  # Add Border ---------------------
  openxlsx::addStyle(wb, sheet = sheet, border_style_bottom_double(),
                     rows = (start_row - 1 + rows),
                     cols = start_col:end_col, gridExpand = T, stack = T)
  openxlsx::addStyle(wb, sheet = sheet, border_style_left(),
                     rows = start_row:(start_row - 1 + rows),
                     cols = start_col:end_col, gridExpand = T, stack = T)

  for(i in 1:dim(c_names)[1]){
    m <- which(grepl("BBBBB_", c_names[i,]))
    if(length(m) != 0){
      purrr::walk(m,
                  ~openxlsx::addStyle(wb, sheet, border_style_bottom(),
                                      cols = (.x - 1):
                                        (.x + as.numeric(gsub("MMMMM_|BBBBB_", "", c_names[i, .x])) - 2),
                                      rows = (i + start_row - 1),
                                      gridExpand = T, stack = T))
    }
  }

  return(list(wb = wb,
              length = start_row - 1 + rows))
}
