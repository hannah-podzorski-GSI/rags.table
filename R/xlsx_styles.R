#' Custom Cell Styles for openxlsx
#'
header_style <- function(){
  openxlsx::createStyle(
    fontName = "Arial",
    fontSize = 10,
    fontColour = "black",
    halign = "left",
    valign = "center")
}

#' @rdname add
title1_style <- function(){
  openxlsx::createStyle(
    fontName = "Arial",
    fontSize = 12,
    fontColour = "black",
    halign = "center",
    valign = "center",
    textDecoration = "bold")
}

#' @rdname add
title2_style <- function(){
  openxlsx::createStyle(
    fontName = "Arial",
    fontSize = 11,
    fontColour = "black",
    halign = "center",
    valign = "center")
}

#' @rdname add
title_style <- function(){
  openxlsx::createStyle(
    fontName = "Arial",
    fontSize = 12,
    fontColour = "black",
    halign = "center",
    valign = "center",
    textDecoration = "bold",
    fgFill = "#D9D9D9")
}

#' @rdname add
md_style <- function(){
  openxlsx::createStyle(
    fontName = "Arial",
    fontSize = 10,
    fontColour = "black",
    halign = "left",
    valign = "center",
    wrapText = F)
}

#' @rdname add
d_style <- function(){
  openxlsx::createStyle(
    fontName = "Arial",
    fontSize = 10,
    fontColour = "black",
    halign = "left",
    valign = "center",
    wrapText = T)
}

#' @rdname add
d_style_bold <- function(){
  openxlsx::createStyle(
    fontName = "Arial",
    fontSize = 10,
    fontColour = "black",
    halign = "left",
    valign = "center",
    wrapText = T,
    textDecoration = "bold")
}

#' @rdname add
d_style_center <- function(){
  openxlsx::createStyle(
    fontName = "Arial",
    fontSize = 10,
    fontColour = "black",
    halign = "center",
    valign = "center",
    wrapText = T)
}

#' @rdname add
d_style_right <- function(){
  openxlsx::createStyle(
    fontName = "Arial",
    fontSize = 10,
    fontColour = "black",
    halign = "right",
    valign = "center",
    wrapText = T)
}

#' @rdname add
fill_style <- function(){
  openxlsx::createStyle(
    fontName = "Arial",
    fontSize = 12,
    fontColour = "black",
    halign = "center",
    valign = "center",
    fgFill = "#D9D9D9")
}

#' @rdname add
border_style_right <- function(){
  openxlsx::createStyle(
    border = "right",
    borderStyle = c("thin"))
}

#' @rdname add
border_style_bottom <- function(){
  openxlsx::createStyle(
    border = "bottom",
    borderStyle = c("thin"))
}

#' @rdname add
border_style_left <- function(){
  openxlsx::createStyle(
    border = "left",
    borderStyle = c("thin"))
}

#' @rdname add
border_style_top <- function(){
  openxlsx::createStyle(
    border = "top",
    borderStyle = c("thin"))
}

#' @rdname add
border_style_bottom_double <- function(){
  openxlsx::createStyle(
    border = "bottom",
    borderStyle = c("double"))
}

#' @rdname add
border_style_all_thick <- function(){
  openxlsx::createStyle(
    border = c("bottom", "top", "left", "right"),
    borderStyle = c("thick", "thick", "thick", "thick"))
}

#' @rdname add
border_style_bottom_thick <- function(){
  openxlsx::createStyle(
    border = c("bottom"),
    borderStyle = c("thick"))
}

#' @rdname add
border_style_top_bottom_thick <- function(){
  openxlsx::createStyle(
    border = c("top", "bottom"),
    borderStyle = c("thick", "thick"))
}

#' @rdname add
border_style_side_thick <- function(){
  openxlsx::createStyle(
    border = c("left","right"),
    borderStyle = c("thick", "thick"))
}

#' @rdname add
border_style_box_double_sides <- function(){
  openxlsx::createStyle(
    border = c("left", "right"),
    borderStyle = c("double", "double"))
}

#' @rdname add
border_style_box_double_top <- function(){
  openxlsx::createStyle(
    border = c("top"),
    borderStyle = c("double"))
}

#' @rdname add
border_style_box_double_bottom <- function(){
  openxlsx::createStyle(
    border = c("bottom"),
    borderStyle = c("double"))
}

#' @rdname add
footer_style <- function(){
  openxlsx::createStyle(
    fontName = "Arial",
    fontSize = 10,
    fontColour = "black",
    halign = "left",
    wrapText = T)
}

#' @rdname add
abbr_style <- function(){
  openxlsx::createStyle(
    fontName = "Arial",
    fontSize = 10,
    fontColour = "black",
    halign = "left",
    wrapText = F)
}

#' @rdname add
footer_style_bold <- function(){
  openxlsx::createStyle(
    fontName = "Arial",
    fontSize = 10,
    fontColour = "black",
    textDecoration = "bold",
    halign = "left")
}

