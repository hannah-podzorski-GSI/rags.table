#' Adds Meta information to an excel file
#'
#' @param wb an openxlsx workbook object
#' @param sheet name of sheet in workbook object
#' @param meta_text dataframe of text to display in meta info of workbook object
#' @param d dataframe of data to be included in table
#' @param start_row number for starting row
#' @param start_col number for starting column
#' @param meta_end_col number for the last column of meta info box
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
#' fmt_meta(wb = wb, sheet = "sheet",
#' meta_text = meta_text, d = sample_data,
#' start_row = 2, start_col = 1, meta_end_col = 3)
fmt_meta <- function(wb, sheet, meta_text, d, start_row = 1, start_col = 1,
                     meta_end_col = 3){

  # Creating Text for Meta -----------
  txt <- meta_text %>% tidyr::separate_rows(XXXXX_sub, sep = ", ")

  for(i in which(!is.na(txt$XXXXX_sub))){ #for the rows with text in the XXXXX_sub
    var <- txt$XXXXX_sub[i]
    txt[i, "XXXXX_sub"] <- unique(d %>%
                                    dplyr::pull(!!var)) # replace with unique value for data file
  }

  txt <- txt %>% dplyr::group_by(line, text) %>%
    dplyr::summarise(XXXXX_sub = toString(unique(XXXXX_sub))) %>%
    dplyr::ungroup() %>% dplyr::rowwise() %>%
    dplyr::mutate(text = gsub("XXXXX", XXXXX_sub, text))

  rows <- length(txt$text) # number of rows to be added

  # Add to Meta Info to Workbook ----------
  openxlsx::writeData(wb, sheet, x = txt %>% dplyr::select(text),
                      startCol = start_col, startRow = start_row, colNames = F)

  # Merge Cells -----------------
  purrr::walk(start_row:(start_row - 1 + rows),
              ~openxlsx::mergeCells(wb, sheet,
                                    cols = 1:meta_end_col, rows = .x))

  # Add Style ------------------
  openxlsx::addStyle(wb, sheet, md_style(),
                     rows = start_row:(start_row - 1 + rows),
                     cols = start_col:meta_end_col, gridExpand = T)

  # Add Border ------------------
  openxlsx::addStyle(wb, sheet, border_style_box_double_sides(),
                     rows = start_row:(start_row - 1 + rows),
                     cols = start_col:meta_end_col,
                     stack = T, gridExpand = T)

  openxlsx::addStyle(wb, sheet, border_style_box_double_top(),
                     rows = start_row,
                     cols = start_col:meta_end_col,
                     stack = T, gridExpand = T)

  openxlsx::addStyle(wb, sheet, border_style_box_double_bottom(),
                     rows = (start_row - 1 + rows),
                     cols = start_col:meta_end_col,
                     stack = T, gridExpand = T)

  return(list(wb = wb,
              length = rows + start_row - 1))

}
