#' Templates for creating tables
#'
#' @param file_path text string of file path to save template
era_copc_temp <- function(file_path){
  file.copy(system.file("extdata", "ERA_COPC_Template.xlsx",
                        package = "rags.table"),
            file_path)
}
