
mtcars_summary <-
  summary_reshape(mtcars,
                  vars_to_summary = c("Value" = "mpg",
                                      Number = "vs"),
                  row_var1 = "cyl",
                  col_var = "am")

mtcars_summary_wb <-
  st_add_summary(openxlsx2::wb_workbook("Angel Feliz"),
                 dt_summary,
                 sheet_title = "example",
                 value_col_pattern = "_(Value|Number)$",
                 values_types = "Value",
                 header_has_levels = TRUE,
                 header_avoid_merge = c("Value", "Number"))

use_data(mtcars_summary_wb, internal = TRUE, overwrite = TRUE)
