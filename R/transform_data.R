
#' @import data.table

#' @title Long-to-wide reshaping tool
#'
#' @description Long-to-wide reshaping tool
#'
#' @param DT a data.table
#' @param names_from string vars to be used as columns
#' @param id_cols string vars to be used as rows
#' @param summary_col rows to be used as totals
#' @param values_from rows to be used as values
#' @param ... it goes to \code{dcast.data.table}
#'
#' @return A reshaped data.table
#'
custom_dcast <- function(DT,
                         names_from,
                         id_cols,
                         summary_col,
                         values_from,
                         ...){

  # Validating initial data
  original_col_names <- names(DT)
  stopifnot("DT must has names_from" = all(names_from %chin% original_col_names))
  stopifnot("DT must has id_cols" = all(id_cols %chin% original_col_names))

  # Validation: We need to have a column named as Number or Value
  stopifnot("DT must has the Number or Value columns" = length(values_from) > 0L)

  # We need to spread the data set
  DT_wider <-
    dcast(DT,
          formula = create_formula(id_cols, names_from),
          value.var = values_from,
          ...)


  # We just need to change col names when
  # having 2 columns with values
  if(length(values_from) > 1L){

    DT_wider_names <- names(DT_wider)

    for(value_i in values_from){

      start_pattern <- paste0("^", value_i,"_")
      cols_to_change <- grep(x = DT_wider_names, pattern = start_pattern, value = TRUE)

      data.table::setnames(DT_wider,
                           cols_to_change,
                           sub(x = cols_to_change,
                               pattern = start_pattern,
                               replacement = "") |>
                             paste0("_", value_i))

    }

    # Defining col-order
    data.table::setcolorder(DT_wider,
                            c(id_cols, sort(DT_wider_names)) |> unique())

  }else{

    value_col_order <-
      names(DT_wider) |>
      setdiff(y = c(summary_col, id_cols)) |>
      sort()

    # Defining col-order
    data.table::setcolorder(DT_wider, c(id_cols, value_col_order, summary_col))

  }


  # We want to keep the id_vars for other functions
  setattr(DT_wider, "id_vars", id_cols)

  # Exporting Results
  return(DT_wider)

}




#' Creates a tidy summary
#'
#' @description  Create all the values to be displayed in the summary table in tidy format.
#'
#' If after creating the summary you need to add more columns to the final table,
#' it's recommended to use this tidy format before casting the final table.
#'
#' @param DT a data.table or data.frame with the data to summarize.
#' @param vars_to_count a named string vector with the columns we want to apply an unique count. The names will be used to identify the values after casting.
#' @param vars_to_summary a named string vector with the columns we want to summary using the function defined in the \code{value_fun} argument. The names will be used to identify the values after casting.
#' @param row_var1 a string representing a categorical column to be represented on rows after casting the table.
#' @param col_var a string representing a categorical column to be represented on columns after casting the table.
#' @param row_var2 a string representing a categorical column to be represented as sub-rows of the column indicated in the \code{row_var1} argument.
#' @param split_var a string representing a categorical column used as a reference to create more than want summary table of the same data.
#' @param row_var1_prefix a prefix to to identify a summary row.
#' @param value_fun defines a simple summary function for the variables defined in the \code{vars_to_summary} argument. By default it calculates the sum of variables.
#' @param value_total_fun defines the functions to be used to calculate the Total by row and column.
#' @param col_title a title to describe the variable used to create the columns.
#' @param total_title_by_row a title to describe the column created to have the total of each row.
#' @param total_title_by_col a title to describe the row created to have the total of each column.
#' @param col_title_sep a separator needed to define hierarchy between \code{col_title} and \code{col_var}.
#' @param reshape_summary if \code{TRUE} reshape the data according prior arguments.
#' @param ... it goes to \code{dcast.data.table}
#'
#' @return A \code{data.table}
#'
#' @export
#'
summary_reshape <- function(
    # The data to use
    DT,

    # Numeric variables
    vars_to_count = NULL,
    vars_to_summary = NULL,

    # String variables
    row_var1,
    col_var = NULL,
    row_var2 = NULL,
    split_var = NULL,

    # Argument with defaults
    value_fun = \(x) sum(x, na.rm = TRUE),
    value_total_fun = value_fun,
    row_var1_prefix = if(is.null(row_var2)) NULL else col_var,
    col_title = col_var,
    total_title_by_row = "Total",
    total_title_by_col = "Grand Total",
    col_title_sep = "_",
    reshape_summary = TRUE,
    ...
) {

  # Validating data format
  stopifnot("DT must be a data.table or data.frame object" = is.data.frame(DT))

  # Validating variables to
  stopifnot("Select a variable to create the summary" = length(c(vars_to_count, vars_to_summary)) > 0L)
  stopifnot("vars_to_count must have names" = is.null(vars_to_count) || !is.null(attr(vars_to_count, "names")))
  stopifnot("vars_to_summary must have names" = is.null(vars_to_summary) || !is.null(attr(vars_to_summary, "names")))

  # All string are in the DT
  stopifnot("One of the variables is not present in DT" =
              c(vars_to_count,
                vars_to_summary,
                row_var1,
                col_var,
                row_var2,
                split_var) %chin% names(DT))

  # Defining values to summarize
  vars_to_report <- c("Value","Number")[c(!is.null(vars_to_summary),
                                          !is.null(vars_to_count))]

  # Making sure that DT is in correct format
  if(data.table::is.data.table(DT)){

    # We don't want side effects
    DT <- data.table::copy(DT)

  }else{

    # We need a data.table to continue the process
    DT <- data.table::as.data.table(DT)

  }

  # Changing NULL for NA to tracking what is missing after
  # combining the values to defini the group column
  if(is.null(col_var)) col_var <- NA_character_
  if(is.null(split_var)) split_var <- NA_character_
  if(is.null(row_var2)) row_var2 <- NA_character_
  if(is.null(col_title)) col_title <- NA_character_

  # Defining variables to group by and saving
  id_vars_names <-
    c(split_var = split_var,
      col_var = col_var,
      row_var1 = row_var1,
      row_var2 = row_var2) |>
    (\(x) x[!is.na(x)])()

  id_vars <- unname(id_vars_names)

  # Add title to col_var
  if(!is.na(col_var)){

    DT[, (col_var) := paste0(col_title, col_title_sep, var),
       env = list(var = col_var)]

  }

  # Defining j expressions for values

  expr_value <- NULL
  expr_total_value <- NULL

  if(!is.null(vars_to_summary)){

    expr_value <-
      lapply(vars_to_summary,
             \(x) data.table::substitute2(expr = value_fun(var), env = list(var = x)))


    vars_to_summary_names <- names(vars_to_summary)
    data.table::setattr(vars_to_summary_names, "names", vars_to_summary_names)

    expr_total_value <-
      lapply(vars_to_summary_names,
             \(x) data.table::substitute2(expr = value_total_fun(var), env = list(var = x)))

  }


  # Defining j expressions for numbers

  expr_number <- NULL
  expr_total_number <- NULL

  if(!is.null(vars_to_count)){

    expr_number <-
      lapply(vars_to_count,
             \(x) data.table::substitute2(expr = data.table::uniqueN(var), env = list(var = x)))

    vars_to_count_names <- names(vars_to_count)
    data.table::setattr(vars_to_count_names, "names", vars_to_count_names)

    expr_total_number <-
      lapply(vars_to_count_names,
             \(x) data.table::substitute2(expr = sum(var), env = list(var = x)))

  }


  # Joining expressions

  expr_initial <- append(expr_value, expr_number)
  expr_cumulative <- append(expr_total_value, expr_total_number)


  # Defining the most detailed summary
  output_summary <-
    DT[, j_expr,
       by = id_vars,
       env = list(j_expr = expr_initial)]


  # Adding summary by row_var1
  if(!is.na(row_var2)){

    id_vars_rowvar2 <- id_vars[id_vars != row_var2]

    output_summary <-
      output_summary[, j_expr,
                     by = id_vars_rowvar2,
                     env = list(j_expr = expr_cumulative)
      ][, (row_var2) := var, env = list(var = row_var1)
      ][, rbind(output_summary, .SD)]

  }


  # Adding total per row (including row_var1 summary)
  if(!is.na(col_var)){

    id_vars_colvar <- id_vars[id_vars != col_var]

    output_summary <-
      output_summary[, j_expr,
                     by = id_vars_colvar,
                     env = list(j_expr = expr_cumulative)
      ][, (col_var) := total_title_by_row
      ][, rbind(output_summary, .SD)]

  }

  # Adding total per col
  output_summary <-
    if(!is.na(row_var2)){

      id_vars_rowvar <- id_vars[!id_vars %chin% c(row_var1, row_var2)]

      output_summary[var1 == var2,
                     j_expr,
                     by = id_vars_rowvar,
                     env = list(j_expr = expr_cumulative,
                                var1 = row_var1,
                                var2 = row_var2)
      ][, (row_var1) := total_title_by_col
      ][, (row_var2) := total_title_by_col
      ][, rbind(output_summary, .SD)]

    }else{

      id_vars_rowvar1 <- id_vars[id_vars != row_var1]

      output_summary[, j_expr,
                     by = id_vars_rowvar1,
                     env = list(j_expr = expr_cumulative)
      ][, (row_var1) := total_title_by_col
      ][, rbind(output_summary, .SD)]
    }

  # Creating a matrix with the report
  if(reshape_summary){

    output_summary <-
      custom_dcast(DT = output_summary,
                   names_from = id_vars_names["col_var"],
                   id_cols = id_vars_names[names(id_vars_names) != "col_var"],
                   summary_col = total_title_by_row,
                   values_from = c(names(vars_to_count),
                                   names(vars_to_summary)),
                   ...)

  }



  return(output_summary)

}



#' @title Arrange the reshaped table
#'
#' @param DT a reshaped data.table
#' @param row_var1 a string representing a categorical column to be represented on rows after casting the table.
#' @param row_var2 a string representing a categorical column to be represented as sub-rows of the column indicated in the `row_var1` argument.
#' @param split_var a string representing a categorical column used as a reference to create more than want summary table of the same data.
#' @param row_var1_prefix defines the prefix to be used in summary rows related to  `row_var1` argument.
#' @param total_title_by_row a title to describe the column created to have the total of each row.
#' @param total_title_by_col a title to describe the row created to have the total of each column.
#'
#' @return A `data.table`
#' @export

arrange_table <- function(DT,
                          row_var1,
                          row_var2 = NULL,
                          split_var = NULL,
                          row_var1_prefix = if(!is.null(row_var2)) paste0(row_var1,": ") else NULL,
                          total_title_by_row = "Total",
                          total_title_by_col = "Grand Total") {

  # Defining valid variables
  vars_list <-
    list(row_var1_d = row_var1,
         row_var2_d = row_var2,
         split_var_d = split_var,
         total_title_by_row_d = total_title_by_row) |>
    (\(x) x[!is.null(x)])()

  # Arranging the table
  dt_arranged <-
    DT[do.call("order",
               list(if(!is.null(split_var)) split_var_d else NULL,
                    if(!is.null(row_var2)) (row_var2_d == total_title_by_col) else (row_var1_d == total_title_by_col),
                    if(!is.null(row_var2)) row_var1_d else -total_title_by_row_d,
                    if(!is.null(row_var2)) -(row_var2_d == row_var1_d) else NULL,
                    if(!is.null(row_var2)) -total_title_by_row_d else NULL) |>
                 (\(x) x[!is.null(x)])()),
       env = vars_list]

  if(!is.null(row_var1_prefix)){

    # Adding prefix
    dt_arranged[row_var2_d == row_var1_d &
                  row_var2_d != total_title_by_col,
                (row_var2) := paste0(row_var1_prefix, row_var1_d),
                env = vars_list]

    # Removing original column
    dt_arranged[, (row_var1) := NULL][]

  }

  return(dt_arranged)
}


#' Creates a list based on summary
#'
#' @description   Divide the rows of the data.table into different elements
#' of a list but keeping the names provided in the \code{split_var}.
#'
#' @param DT a reshaped data.table
#' @param split_var a string making the reference to the column we need to use a reference to split the \code{data.table}.
#' @param lookup_v a named vector used to change the groups used to split the rows.
#' @param name_suffix adding a suffix to the table name to be show as the title of each sheet.
#' @param decreasing it can revert the order of the \code{data.tables} in the resulting list.
#'
#' @return a \code{list} of \code{data.table}
#' @export
#'
split_summary <- function(DT,
                          split_var,
                          lookup_v = NULL,
                          name_suffix = NULL,
                          decreasing = FALSE){

  # We don't want side effects
  DT <- data.table::copy(DT)

  # If the table is empty
  if(nrow(DT) == 0L) return(vector("list"))

  # Changing the split criteria
  if(is.null(lookup_v)){
    DT[, split_value := split_var_value,
       env = list(split_var_value = split_var)]
  }else{
    DT[, split_value := lookup_v[split_var_value],
       env = list(split_var_value = split_var)]
  }

  # Setting a key
  data.table::setkey(DT, split_value)

  # Listing the iterations needed
  values_to_split <- DT[, unique(split_value)] |> sort(decreasing = decreasing)

  # Defining a name for each split
  table_names <-
    if(is.null(name_suffix)){
      values_to_split
    }else{
      paste0(values_to_split, name_suffix)
    }

  # Applying the split
  split_list <-
    lapply(seq_along(values_to_split),
           function(x){

             # Using key filtering to select the data
             DT[list(values_to_split[x]),
                nomatch = NULL,
                # Remove columns used to split
                j = !c("split_value", split_var),
                with = FALSE] |>

               # Assigning table name to each table
               data.table::setattr("table_name", table_names[x])

           })

  data.table::setattr(split_list, "names", table_names)

  return(split_list)

}


