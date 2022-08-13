#' Adds superscripts to text in excel
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
#' excel_superscript(strings = wb$sharedStrings)
excel_superscript <- function(strings){
  for(i in grep("\\^[A-z0-9\\s]*", strings)){
    # insert additioanl formating in shared string
    strings[[i]] <- gsub("<si>", "<si><r>", gsub("</si>", "</r></si>", strings[[i]]))

    # if the character string !stop! is included in the string then format text inbetween ^ and !stop!
    # Otherwise format one character after ^
    if(grepl("!stop!",  strings[[i]])){
      strings[[i]] <- gsub("\\^[A-z0-9\\s]*", "</t></r><r><rPr><vertAlign val=\"superscript\"/><sz val =\"10\"/><rFont val =\"Arial\"/></rPr><t xml:space=\"preserve\">", strings[[i]])
      # strings[[i]] <- gsub("!stop!","</t></r><r><t xml:space=\"preserve\">", strings[[i]])
      strings[[i]] <- gsub("!stop!","</t></r><r><rPr><sz val =\"10\"/><rFont val =\"Arial\"/></rPr><t xml:space=\"preserve\">", strings[[i]])
    }else{
      n <- nchar(gsub("\\^.*","",strings[[i]])) + 2
      n_tot <- nchar(strings[[i]])

      strings[[i]] <- paste0(substr(strings[[i]], start = 1, stop = n),"</t></r><r><rPr><sz val =\"10\"/><rFont val =\"Arial\"/></rPr><t xml:space=\"preserve\">",
                             substr(strings[[i]], start = n+1, stop = n_tot))
      strings[[i]] <-gsub("\\^", "</t></r><r><rPr><vertAlign val=\"superscript\"/><sz val =\"10\"/><rFont val =\"Arial\"/></rPr><t xml:space=\"preserve\">", strings[[i]])
    }
  }

  return(strings)

}
