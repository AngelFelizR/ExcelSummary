#' @importFrom data.table := %chin%
#' @importFrom stats as.formula
#' @import openxlsx2
#' @importFrom stringr str_match str_split_fixed

# Transform data ----

create_formula <- function(y, x){

  paste0(paste0(paste0("`", y, "`"), collapse = " + "),
         " ~ ",
         paste0(paste0("`", x, "`"), collapse = " + ") ) |>
    stats::as.formula()

}

`:=` <- function(...) data.table::`:=`(...)

`%chin%` <- function(...) data.table::`%chin%`(...)

`%like%` <- function(...) data.table::`%like%`(...)

# Excel ----


st_header_matrix <- function(DT, sep = "_"){

  # We need to make sure that all titles has
  # the same number of sep before applying the split
  check_n_sep <-
    data.table::data.table(col_names = names(DT)
    )[, n_missing_levels :=
        str_count(col_names, sep) |>
        (`+`)(1L) |>
        (\(x) max(x) - x)()]

  check_n_sep[, .(new_col_name =
                    fifelse(n_missing_levels != 0L,
                            stringr::str_match(col_names,"^((\\w| |%)+)_|^((\\w| |%)+)$") |>
                              (\(x) data.table::fcoalesce(x[,2], x[,4]))() |>
                              rep(n_missing_levels) |>
                              paste0(collapse = sep) |>
                              paste0("_",col_names),
                            col_names) ),
              by = "col_names"
  ][, new_col_name] |>

    # We can split and transpose the matrix
    stringr::str_split_fixed(pattern = "_",
                             n = max(check_n_sep$n_missing_levels) + 1L) |>
    t()

}


wb_merge_header <- function(wb,
                            sheet = openxlsx2::current_sheet(),
                            header_matrix,
                            start_row,
                            start_col,
                            custom_h_align,
                            color_fill = openxlsx2::wb_color(hex = "FF1F497D"),
                            color_font = openxlsx2::wb_color(name = "white"),
                            color_border_bottom = openxlsx2::wb_color(name = "white"),
                            bold_font = TRUE,
                            h_align = "center",
                            v_align = "center",
                            wrap_text = TRUE,
                            avoid_merge = c("Number", "Value")){

  # Adding Header Matrix to Excel
  wb <-
    openxlsx2::wb_add_data(wb, x = header_matrix,
                           startCol = start_col,
                           startRow = start_row,
                           colNames = FALSE,
                           na.strings = "")


  # Bottom borders for avoid merge values

  wb <-
    which(header_matrix[nrow(header_matrix),] %chin% avoid_merge) |>
    (\(x) x + start_col - 1L)() |>
    range() |>
    openxlsx2::wb_dims(row = nrow(header_matrix) + start_row - 1L) |>
    paste0(collapse = ":") |>
    openxlsx2::wb_add_border(wb = wb,
                             sheet = sheet,
                             left_border = NULL,
                             right_border = NULL,
                             top_border = NULL)


  # We need to defining the cells to change

  merge_instructions <-

    # Listing all values that need to be merged
    as.vector(header_matrix) |>
    unique() |>
    setdiff(y = avoid_merge) |>

    # For each value we need to define and start point and positions to move
    lapply(M = header_matrix,
           FUN = \(check_value, M){
             data.table::as.data.table(which(M == check_value, arr.ind = TRUE)
             )[,.(check_value,
                  # Start point
                  row = row - 1L,
                  col = col - 1L,
                  # Total of positions
                  down = data.table::uniqueN(row) - 1L,
                  right = data.table::uniqueN(col) - 1L)
             ][1L]
           }) |>
    data.table::rbindlist() |>

    # Translate merge movements into Excel dimensions
    (\(DT) DT[, .(check_value,
                  value_row_start = row + start_row,
                  value_row_end = row + start_row + down,
                  value_col_start = col + start_col,
                  value_col_end = col + start_col + right)
    ][, dims := paste0(openxlsx2::wb_dims(value_row_start, value_col_start),
                       ":",
                       openxlsx2::wb_dims(value_row_end, value_col_end))]  )()

  data.table::setkey(merge_instructions, check_value)

  # Defining dims to border

  # We need to highlight the values that are:
  # - Present in first row
  # - Not present last row
  white_bottom_headers <-
    merge_instructions[value_row_start == start_row & value_row_end < max(value_row_end),
                       check_value]

  first_col_value <- header_matrix[1L, 1L]
  last_col_value <- header_matrix[1L, ncol(header_matrix)]

  # We need to border at top and bottom columns to be merged in a single column
  # after taking out the first and last columns
  top_bottom_headers <-
    merge_instructions[value_row_start == start_row & value_row_end == max(value_row_end),
                       check_value] |>
    setdiff(c(first_col_value, last_col_value))

  for (variable in merge_instructions$check_value) {

    # Merging instructions
    rows <- merge_instructions[list(variable), c(value_row_start:value_row_end)]
    cols <- merge_instructions[list(variable), c(value_col_start:value_col_end)]

    # Dims to format
    dims <- merge_instructions[.(variable), dims]

    # Merging values and applying fill and font color
    wb <-
      openxlsx2::wb_merge_cells(wb, sheet = sheet, rows = rows, cols = cols) |>
      openxlsx2::wb_add_fill(sheet = sheet, dims = dims,
                             color = color_fill) |>
      openxlsx2::wb_add_font(sheet = sheet, dims = dims,
                             color = color_font,
                             bold = fifelse(bold_font, "single", "") )

    # Defining borders all cells but avoid_merge values
    if(variable %chin% first_col_value){
      wb <-
        openxlsx2::wb_add_border(wb,
                                 sheet = sheet,
                                 dims = dims,
                                 right_border = NULL)
    }

    if(variable %chin% last_col_value){
      wb <-
        openxlsx2::wb_add_border(wb,
                                 sheet = sheet,
                                 dims = dims,
                                 left_border = NULL)
    }

    if(variable %chin% white_bottom_headers){
      wb <-
        openxlsx2::wb_add_border(wb, sheet = sheet, dims = dims,
                                 bottom_color = color_border_bottom,
                                 left_border = NULL,
                                 right_border = NULL,
                                 top_border = "thin")
    }

    if(variable %chin% top_bottom_headers){
      wb <-
        openxlsx2::wb_add_border(wb,
                                 sheet = sheet,
                                 dims = dims,
                                 right_border = NULL,
                                 left_border = NULL)
    }


    # Know we can apply the custom horizontal align and wrap the text
    wb <-
      if(!missing(custom_h_align) && variable %chin% names(custom_h_align)){
        openxlsx2::wb_add_cell_style(wb,
                                     sheet = sheet,
                                     dims = dims,
                                     horizontal = custom_h_align[variable],
                                     vertical = v_align,
                                     wrapText = if(wrap_text) "1" else NULL )
      }else{
        openxlsx2::wb_add_cell_style(wb,
                                     sheet = sheet,
                                     dims = dims,
                                     horizontal = h_align,
                                     vertical = v_align,
                                     wrapText = if(wrap_text) "1" else NULL )
      }

  }



  # Formatting avoided columns

  last_header_row <- nrow(header_matrix)
  avoid_positions <- which(header_matrix[last_header_row, ] %chin% avoid_merge) + start_col - 1L

  avoid_dims <- openxlsx2::wb_dims(start_row + last_header_row - 1L, range(avoid_positions)) |> paste0(collapse = ":")

  wb <-
    openxlsx2::wb_add_fill(wb,
                           sheet = sheet,
                           dims = avoid_dims,
                           color = color_fill) |>
    openxlsx2::wb_add_font(sheet = sheet,
                           dims = avoid_dims,
                           color = color_font,
                           bold = fifelse(bold_font, "single", "") ) |>
    openxlsx2::wb_add_cell_style(sheet = sheet,
                                 dims = avoid_dims,
                                 horizontal = h_align,
                                 vertical = v_align)

  return(wb)

}


apply_num_format <- function(wb,
                             DT_names,
                             col_pattern,
                             start_col,
                             start_row,
                             last_row,
                             col_width,
                             col_h_align,
                             col_v_align,
                             col_numfmt){

  # We don't know if we need to start a different column, not just A
  col_positions <- which(DT_names %like% col_pattern) + start_col - 1L

  # We need to make sure that numbers have enough space to show up
  wb <- openxlsx2::wb_set_col_widths(wb, cols = col_positions, widths = col_width)

  # Defining ranges to edit at detail level
  value_dims <-
    paste0(openxlsx2::wb_dims(start_row + 1L, col_positions),
           ":",
           openxlsx2::wb_dims(last_row, col_positions))


  # Applying formatting on each column at detail and total levels
  for(dim_i in seq_along(col_positions)){

    wb <-
      openxlsx2::wb_add_cell_style(wb,
                                   dims = value_dims[dim_i],
                                   horizontal = col_h_align,
                                   vertical = col_v_align) |>
      openxlsx2::wb_add_numfmt(dims = value_dims[dim_i],
                               numfmt = col_numfmt)
  }

  return(wb)

}


