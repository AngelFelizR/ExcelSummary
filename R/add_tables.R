
#' @import openxlsx2
#' @import data.table

#' @title Defines a custom format
#'
#' @param numfmt a string value
#' @param width a numeric value
#' @param h_align a string value
#' @param v_align a string value
#' @param wrap_text a Boolean value
#'
#' @return a named \code{list} based on column properties
#' @export
#'
col_format <- function(numfmt,
                       width,
                       h_align,
                       v_align = "center",
                       wrap_text = FALSE){

  list(numfmt = numfmt,
       width = width,
       h_align = h_align,
       v_align = v_align,
       wrap_text = wrap_text)
}


#' @title  Add sheet with a summary table
#'
#' @description This function add a summery table to an existing Excel document.
#'
#' @param wb a workbook.
#' @param DT a reshaped \code{data.table}.
#' @param sheet_title a string to define the title to define the sheet.
#' @param highlight_col a string to define which have a pattern to highlight as summary row.
#' @param highlight_row_pattern the pattern to find in the \code{highlight_col} variable.
#' @param custom_col_format a named \code{list} based on DT var names with the results of calling the \code{col_format} function.
#' @param header_custom_h_align align content horizontal ('left', 'center', 'right')
#' @param sheet_name a string to define the sheet name.
#' @param values_types a
#' @param sheet_title_row a
#' @param sheet_title_size a
#' @param start_row a
#' @param start_col a
#' @param zoom a
#' @param header_has_levels a
#' @param header_color_fill a
#' @param header_color_font a
#' @param header_color_border_bottom a
#' @param header_bold_font a
#' @param header_h_align a
#' @param header_v_align a
#' @param header_wrap_text a
#' @param header_avoid_merge a
#' @param num_col_pattern a
#' @param num_col_numfmt a
#' @param num_col_width a
#' @param num_col_h_align a
#' @param num_col_v_align a
#' @param value_col_pattern a
#' @param value_col_numfmt a
#' @param value_col_width a
#' @param value_col_h_align a
#' @param value_col_v_align a
#' @param highlight_row_fill a
#'
#' @return a workbook object
#' @export
#'
st_add_summary <- function(wb,
                           DT,
                           sheet_title,
                           highlight_col,
                           highlight_row_pattern,
                           custom_col_format,
                           header_custom_h_align,
                           sheet_name = sheet_title,
                           values_types = c("Number", "Value"),
                           sheet_title_row = 1L,
                           sheet_title_size = 22,
                           start_row = 4L,
                           start_col = 1L,
                           zoom = 100,
                           header_has_levels = TRUE,
                           header_color_fill = openxlsx2::wb_color(hex = "FF1F497D"),
                           header_color_font = openxlsx2::wb_color(name = "white"),
                           header_color_border_bottom = openxlsx2::wb_color(name = "white"),
                           header_bold_font = TRUE,
                           header_h_align = "center",
                           header_v_align = "center",
                           header_wrap_text = TRUE,
                           header_avoid_merge = values_types,
                           num_col_pattern = "_Number$",
                           num_col_numfmt = "#,##0",
                           num_col_width = 10,
                           num_col_h_align = "center",
                           num_col_v_align = "center",
                           value_col_pattern = "_Value$",
                           value_col_numfmt = "$#,##0",
                           value_col_width = 12,
                           value_col_h_align = "center",
                           value_col_v_align = "center",
                           highlight_row_fill = openxlsx2::wb_color(hex = "FFD9D9D9")){


  # Defining reference points

  DT_names <- names(DT)
  n_cols <- length(DT_names)

  # Changing start point of dims to apply borders

  if(header_has_levels){

    header_matrix <- st_header_matrix(DT)
    start_row_original <- start_row
    start_row <- start_row + nrow(header_matrix) - 1L

  }else{

    # DEFINING HEADER BORDERS

    # It's important to split the header row to apply borders
    header_first_col <- openxlsx2::wb_dims(start_row, start_col)
    header_last_col <- openxlsx2::wb_dims(start_row, start_col + n_cols - 1L)
    header_mid_col <- paste0(
      openxlsx2::wb_dims(start_row, start_col + 1L), ":",
      openxlsx2::wb_dims(start_row, start_col + n_cols - 2L)
    )

    # We use this dims to apply most of the formats
    header_row <- paste0(header_first_col, ":", header_last_col)


    # DEFINING BODY BORDERS

    # To define border for left side
    body_top_left <- openxlsx2::wb_dims(start_row + 1L, start_col)
    body_mid_left <- paste0(
      openxlsx2::wb_dims(start_row + 2L, start_col),":",
      openxlsx2::wb_dims(last_row - 1L, start_col)
    )
    body_bottom_left <- openxlsx2::wb_dims(last_row, start_col)

    # To define border for bottom mid
    body_bottom_mid <- paste0(
      openxlsx2::wb_dims(last_row, start_col + 1L), ":",
      openxlsx2::wb_dims(last_row, start_col + n_cols - 2L)
    )

    # To define border for right side
    body_top_right <- openxlsx2::wb_dims(start_row + 1L, start_col + n_cols - 1L)
    body_mid_right <- paste0(
      openxlsx2::wb_dims(start_row + 2L, start_col + n_cols - 1L),":",
      openxlsx2::wb_dims(last_row - 1L, start_col + n_cols - 1L)
    )
    body_bottom_right <- openxlsx2::wb_dims(last_row, start_col + n_cols - 1L)

  }

  last_row <- start_row + nrow(DT)


  # Adding a sheet, sheet title and base data

  wb <-
    # New sheet
    openxlsx2::wb_add_worksheet(wb,
                                sheet = sheet_name,
                                gridLines = FALSE,
                                zoom = zoom) |>
    openxlsx2::wb_page_setup(orientation = "landscape",
                             paperSize = 1,
                             fitToWidth = TRUE,
                             fitToHeight = TRUE,
                             left = 0.25 ,
                             right = 0.08,
                             top = 0.6,
                             bottom = 0.08,
                             header = 0.08,
                             footer = 0.08) |>

    # Adding and formatting sheet title
    openxlsx2::wb_add_data(x = sheet_title) |>
    openxlsx2::wb_add_font(size = sheet_title_size, bold = "single") |>
    openxlsx2::wb_add_border(dims = paste0(openxlsx2::wb_dims(sheet_title_row, start_col),
                                ":",
                                openxlsx2::wb_dims(sheet_title_row, ncol(DT))),
                  left_border = NULL,
                  right_border = NULL,
                  top_border = NULL) |>

    # Adding data to report
    openxlsx2::wb_add_data(x = DT,
                           start_row = start_row,
                na.strings = "")


  # Formatting headers

  if(header_has_levels){

    wb <-
      wb_merge_header(wb,
                      header_matrix = header_matrix,
                      start_row = start_row_original,
                      start_col = start_col,
                      custom_h_align = header_custom_h_align,
                      color_fill = header_color_fill,
                      color_font = header_color_font,
                      color_border_bottom = header_color_border_bottom,
                      bold_font = header_bold_font,
                      h_align = header_h_align,
                      v_align = header_v_align,
                      wrap_text = header_wrap_text,
                      avoid_merge = header_avoid_merge)

  }else{

    wb <-
      # General formatting
      openxlsx2::wb_add_fill(wb,
                             dims = header_row,
                             color = header_color_fill) |>
      openxlsx2::wb_add_font(dims = header_row,
                             color = header_color_font,
                             bold = fifelse(header_bold_font, "single", "") ) |>
      openxlsx2::wb_add_cell_style(dims = header_row,
                                   horizontal = header_h_align,
                                   vertical = header_v_align,
                                   wrapText = if(header_wrap_text) "1" else NULL ) |>
      # Applying header borders
      openxlsx2::wb_add_border(dims = header_first_col,
                    left_border = "thin",
                    right_border = NULL,
                    bottom_border = "thin",
                    top_border = "thin") |>
      openxlsx2::wb_add_border(dims = header_last_col,
                    left_border = NULL,
                    right_border = "thin",
                    bottom_border = "thin",
                    top_border = "thin") |>
      openxlsx2::wb_add_border(dims = header_mid_col,
                    left_border = NULL,
                    right_border = NULL,
                    bottom_border = "thin",
                    top_border = "thin")

    # Now we need to overwrite some cell styles
    if(!missing(header_custom_h_align)){
      for(col_i in names(header_custom_h_align)){
        wb <-
          openxlsx2::wb_add_cell_style(wb,
                                       dims = openxlsx2::wb_dims(
                                         start_row,
                                         start_col - 1L + which(DT_names == col_i)),
                            horizontal = header_custom_h_align[col_i],
                            vertical = header_v_align,
                            wrapText = if(header_wrap_text) "1" else NULL )

      }
    }


  }


  # Formatting column by column

  if("Number" %chin% values_types){
    wb <-
      apply_num_format(wb = wb,
                       DT_names = DT_names,
                       col_pattern = num_col_pattern,
                       start_col = start_col,
                       start_row = start_row,
                       last_row = last_row,
                       col_width = num_col_width,
                       col_h_align = num_col_h_align,
                       col_v_align = num_col_v_align,
                       col_numfmt = num_col_numfmt)
  }


  if("Value" %chin% values_types){

    wb <-
      apply_num_format(wb = wb,
                       DT_names = DT_names,
                       col_pattern = value_col_pattern,
                       start_col = start_col,
                       start_row = start_row,
                       last_row = last_row,
                       col_width = value_col_width,
                       col_h_align = value_col_h_align,
                       col_v_align = value_col_v_align,
                       col_numfmt = value_col_numfmt)

  }


  if(!missing(custom_col_format)){

    for(col_i in names(custom_col_format)){

      custom_col_position <- which(DT_names == col_i) + start_col - 1L

      wb$set_col_widths(cols = custom_col_position,
                        widths = custom_col_format[[col_i]][["width"]])

      custom_dim <-
        openxlsx2::wb_dims(
          rows = c((start_row + 1L):last_row),
          from_col = custom_col_position
        )

      wb <-
        openxlsx2::wb_add_cell_style(wb, dims = custom_dim,
                          horizontal = custom_col_format[[col_i]][["h_align"]],
                          vertical  = custom_col_format[[col_i]][["v_align"]],
                          wrapText = if(custom_col_format[[col_i]][["wrap_text"]]) "1" else NULL) |>
        openxlsx2::wb_add_numfmt(dims = custom_dim,
                      numfmt = custom_col_format[[col_i]][["numfmt"]])

    }


  }


  # Formatting column by column

  if(!missing(highlight_col)){

    highlight_dims <-
      c(which(DT[[highlight_col]] %like% highlight_row_pattern) + start_row,
        last_row) |>
      (\(x) openxlsx2::wb_dims(cols = start_col:(ncol(DT) + start_col -1L),
                               rows = x[1L]:x[2L]) )()

    for(highlight_i in highlight_dims){

      wb <-
        openxlsx2::wb_add_fill(wb, dims = highlight_i,
                    color = highlight_row_fill) |>
        openxlsx2::wb_add_font(dims = highlight_i,
                    bold = "single")

    }

  }else{

    last_row_dims <-
      openxlsx2::wb_dims(from_row = last_row,
                         cols = start_col:( ncol(DT) + start_col -1L))


    wb <-
      openxlsx2::wb_add_fill(wb, dims = last_row_dims,
                  color = highlight_row_fill) |>
      openxlsx2::wb_add_font(dims = last_row_dims,
                  bold = "single")

  }


  # Adding borders

  if(header_has_levels){

    # Defining border rules

    table_borders <-
      data.table::data.table(original_name = DT_names,
                 last_group = header_matrix[nrow(header_matrix) - 1L,]
      )[,`:=`(last_count = 1:.N,
              last_max_count = .N),
        by = "last_group"
      ][, `:=`(last_group = NULL,
               last_count = NULL,
               last_max_count = NULL,
               left_border = data.table::fifelse(last_count == 1L | .I == 1L, 2L, 1L),
               right_border = data.table::fifelse(last_count == last_max_count | .I == .N, 2L, 1L))][]

    data.table::setkey(table_borders, original_name)


    for(col_i in DT_names){

      col_left_border <- list(NULL,"thin")[[table_borders[.(col_i), left_border]]]
      col_right_border <- list(NULL,"thin")[[table_borders[.(col_i), right_border]]]
      border_col_position <- which(DT_names == col_i) + start_col - 1L

      border_body_dim <-
        openxlsx2::wb_dims(from_col = border_col_position,
                           rows = (start_row + 1L):(last_row - 1L))

      border_last_dim <- openxlsx2::wb_dims(last_row, border_col_position)

      wb <-
        openxlsx2::wb_add_border(wb,
                      dims = border_body_dim,
                      left_border = col_left_border,
                      right_border = col_right_border,
                      bottom_border = NULL,
                      top_border = NULL) |>
        openxlsx2::wb_add_border(dims = border_last_dim,
                      left_border = col_left_border,
                      right_border = col_right_border,
                      bottom_border = "thin",
                      top_border = "thin")

    }

  }else{

    wb <-
      openxlsx2::wb_add_border(wb,
                               dims = body_top_left,
                               top_border = "thin",
                               left_border = "thin",
                               right_border = NULL,
                               bottom_border = NULL) |>
      openxlsx2::wb_add_border(dims = body_mid_left,
                               top_border = NULL,
                               left_border = "thin",
                               right_border = NULL,
                               bottom_border = NULL) |>
      openxlsx2::wb_add_border(dims = body_bottom_left,
                               top_border = "thin",
                               left_border = "thin",
                               right_border = NULL,
                               bottom_border = "thin") |>
      openxlsx2::wb_add_border(dims = body_bottom_mid,
                               top_border = "thin",
                               left_border = NULL,
                               right_border = NULL,
                               bottom_border = "thin") |>
      openxlsx2::wb_add_border(dims = body_top_right,
                               top_border = "thin",
                               left_border = NULL,
                               right_border = "thin",
                               bottom_border = NULL) |>
      openxlsx2::wb_add_border(dims = body_mid_right,
                               top_border = NULL,
                               left_border = NULL,
                               right_border = "thin",
                               bottom_border = NULL) |>
      openxlsx2::wb_add_border(dims = body_bottom_right,
                               top_border = "thin",
                               left_border = NULL,
                               right_border = "thin",
                               bottom_border = "thin")

  }


  return(wb)

}


#' @title Add a table data in workbook
#'
#' @description  Add a table data in workbook
#'
#' @param wb a
#' @param DT a
#' @param sheet_name a
#' @param custom_col_format a
#' @param sheet_title a
#' @param sheet_title_row a
#' @param sheet_title_size a
#' @param ... a
#'
#' @return a workbook object
#' @export

st_add_table <- function(wb,
                         DT,
                         sheet_name,
                         custom_col_format,
                         sheet_title,
                         sheet_title_row = 1L,
                         sheet_title_size = 22L,
                         ...){

  num_cols <- ncol(DT)
  num_rows <- nrow(DT)

  wb <- openxlsx2::wb_add_worksheet(
    wb,
    sheet = sheet_name,
    gridLines = FALSE
  )

  # Adding and formatting sheet title
  if(!missing(sheet_title)){

    wb <-
      openxlsx2::wb_add_data(wb, x = sheet_title) |>
      openxlsx2::wb_add_font(size = sheet_title_size, bold = "single") |>
      openxlsx2::wb_add_border(dims = paste0(openxlsx2::wb_dims(sheet_title_row, 1L),
                                  ":",
                                  openxlsx2::wb_dims(sheet_title_row, num_cols)),
                    left_border = NULL,
                    right_border = NULL,
                    top_border = NULL)

  }


  wb <-
    openxlsx2::wb_add_data_table(wb,
                      x = DT,
                      startRow = fifelse(!missing(sheet_title), 4L, 1L),
                      na.strings = "",
                      ...) |>
    openxlsx2::wb_set_col_widths(cols = 1:num_cols, widths = "auto")


  for(col_i in seq_len(num_cols)){

    wb <- openxlsx2::wb_add_cell_style(wb,
                                       dims = paste0(wb_dims(2L, col_i),
                                                     ":",
                                                     wb_dims(num_rows+1L, col_i)),
                                       vertical =  "center")
  }


  if(!missing(custom_col_format)){

    DT_names <- names(DT)

    cols_to_check <-
      names(custom_col_format) |>
      (\(x) x[x %chin% DT_names])()

    for(col_i in cols_to_check){

      custom_col_position <- which(DT_names == col_i)

      wb <- openxlsx2::wb_set_col_widths(wb,
                                         cols = custom_col_position,
                                         widths = custom_col_format[[col_i]][["width"]])

      custom_dim <-
        paste0(openxlsx2::wb_dims(2, custom_col_position),
               ":",
               openxlsx2::wb_dims(nrow(DT), custom_col_position))


      wb <-
        openxlsx2::wb_add_cell_style(wb, dims = custom_dim,
                          horizontal = custom_col_format[[col_i]][["h_align"]],
                          wrapText = if(custom_col_format[[col_i]][["wrap_text"]]) "1" else NULL) |>
        openxlsx2::wb_add_numfmt(dims = custom_dim,
                      numfmt = custom_col_format[[col_i]][["numfmt"]])


    }

  }

  return(wb)
}

