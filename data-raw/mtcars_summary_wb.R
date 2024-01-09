
# mtcars_headers ----

mtcars_headers <-
  summary_reshape(mtcars,
                  vars_to_summary = c("Value" = "mpg",
                                      Number = "vs"),
                  row_var1 = "cyl",
                  col_var = "am")

mtcars_headers_wb <-
  st_add_summary(openxlsx2::wb_workbook("Angel Feliz"),
                 mtcars_headers,
                 sheet_title = "example",
                 value_col_pattern = "_(Value|Number)$",
                 values_types = "Value",
                 header_has_levels = TRUE,
                 header_avoid_merge = c("Value", "Number"))


# mtcars_row_highlight ----

mtcars_row_highlight <-
  summary_reshape(mtcars,
                  vars_to_summary = c("Value" = "mpg"),
                  row_var1 = "cyl",
                  row_var2 = "vs",
                  col_var = "am") |>
  arrange_table(row_var1 = "cyl",
                row_var2 = "vs")


st_add_summary(openxlsx2::wb_workbook("Angel Feliz"),
               mtcars_row_highlight,
               sheet_title = "example",
               highlight_col = "vs",
               highlight_row_pattern = "^cyl",
               value_col_pattern = "_(Value|Number)$",
               values_types = "Value",
               header_has_levels = TRUE,
               header_avoid_merge = c("Value", "Number")) |>
  openxlsx2::wb_open()


use_data(mtcars_headers_wb, internal = TRUE, overwrite = TRUE)
