#' Adds subscripts to text in excel
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
#' excel_subscript(strings = wb$sharedStrings)
excel_subscript <- function(strings){
  for(i in grep("\\__[A-z0-9\\s]*", strings)){
    # insert additioanl formating in shared string
    strings[[i]] <- gsub("<si>", "<si><r>", gsub("</si>", "</r></si>", strings[[i]]))

    # if the character string !stop! is included in the string then format text inbetween __ and !stop!
    # Otherwise format all text from __ to end of string
    if(grepl("!stop!",  strings[[i]])){
      strings[[i]] <- gsub("\\__[A-z0-9\\s]*", "</t></r><r><rPr><vertAlign val=\"subscript\"/><sz val =\"8\"/><rFont val =\"Arial\"/></rPr><t xml:space=\"preserve\">", strings[[i]])
      strings[[i]] <- gsub("!stop!","</t></r><r><rPr><sz val =\"10\"/><rFont val =\"Arial\"/></rPr><t xml:space=\"preserve\">", strings[[i]])
    }else{
      strings[[i]] <- gsub("\\__", "</t></r><r><rPr><vertAlign val=\"subscript\"/><sz val =\"10\"/><rFont val =\"Arial\"/></rPr><t xml:space=\"preserve\">", strings[[i]])

      n_tot <- nchar(strings[[i]])
      strings[[i]] <- paste0(substr(strings[[i]], start = 1, stop = (n_tot-13)),"</t></r><r><t xml:space=\"preserve\">",
                             substr(strings[[i]], start = n_tot-12, stop = n_tot))
    }

  }
  return(strings)
}
