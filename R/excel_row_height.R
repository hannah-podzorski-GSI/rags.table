#' Fixes row height if sub/superscripts are present in table
#'
#' @param strings sharedStrings from an openxlsx workbook object
#'
#' @return strings for an openxlsx workbook object
#' @export
#'
#' @examples
#' wb <- openxlsx::createWorkbook()
#'
#' openxlsx::addWorksheet(wb, "sheet")
#'
#' excel_row_height(strings = wb$sharedStrings)
excel_row_height <- function(strings){
  for(i in grep("help_fix_height", strings)){
    strings[[i]] <- "<si><r><t xml:space=\"preserve\"></t></r><r><rPr><vertAlign val=\"superscript\"/><sz val =\"10\"/><color rgb =\"FFFFFF\"/><rFont val =\"Arial\"/></rPr><t xml:space=\"preserve\">1</t></r><r><rPr><sz val =\"10\"/><color rgb =\"FFFFFF\"/><rFont val =\"Arial\"/></rPr><t xml:space=\"preserve\"></t></r></si>"
  }
  return(strings)

}
