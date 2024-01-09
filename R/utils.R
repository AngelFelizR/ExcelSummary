#' @importFrom data.table := %chin%
#' @importFrom stats as.formula

create_formula <- function(y, x){

  paste0(paste0(paste0("`", y, "`"), collapse = " + "),
         " ~ ",
         paste0(paste0("`", x, "`"), collapse = " + ") ) |>
    stats::as.formula()

}

`:=` <- function(...) data.table::`:=`(...)

`%chin%` <- function(...) data.table::`%chin%`(...)
