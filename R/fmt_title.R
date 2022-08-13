#' Adds title to an excel file
#'
#' @param wb an openxlsx workbook object
#' @param sheet name of sheet in workbook object
#' @param title_text  dataframe of text to display in title of workbook object
#' @param d dataframe of data to be included in table
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
#' fmt_title(wb = wb, sheet = "sheet",
#' title_text = title_text, d = sample_data,
#' start_row = 2, start_col = 1, end_col = 17)
fmt_title <- function(wb, sheet, title_text, d,
                      start_row = 1, start_col = 1, end_col){

  # Creating Text for Title ---------
  for(i in which(!is.na(title_text$XXXXX_sub))){ # for the rows with text in the XXXXX_sub
    title_text[i, "XXXXX_sub"] <- unique(d %>%
                                           dplyr::pull(!!title_text$XXXXX_sub[i])) # replace with unique value for data file
  }

  txt <- title_text %>% dplyr::rowwise() %>%
    dplyr::mutate(text = gsub("XXXXX", XXXXX_sub, text),
                  text = dplyr::case_when(caps == "upper" ~ toupper(text), # all caps if 'caps' == "upper"
                                          caps == "lower" ~ stringr::str_to_title(text), #c aps first letter of each work if "lower"
                                          T ~ text))

  rows <- length(txt$text) # number of rows to be added

  # Add to Title to Workbook ------
  openxlsx::writeData(wb, sheet = sheet, x = txt %>% dplyr::select(text),
                      startCol = start_col, startRow = start_row, colNames = F)

  # Merge Cells ------------------
  purrr::walk(start_row:(start_row + rows - 1),
              ~openxlsx::mergeCells(wb, sheet,
                                    cols = start_col:end_col,
                                    rows = .x))

  # Add Style ----------
  purrr::walk(start_row - 1 + which(txt$bold == T),
              ~openxlsx::addStyle(wb, sheet, title1_style(),
                                  rows = .x, cols = start_col,
                                  gridExpand = TRUE))

  purrr::walk(start_row - 1 + which(txt$bold == F),
              ~openxlsx::addStyle(wb, sheet, title2_style(),
                                  rows = .x, cols = start_col,
                                  gridExpand = TRUE))

  return(list(wb = wb,
              length = rows + start_row - 1))
}
