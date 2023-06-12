#' Extract information from a cel in an excel sheet
#'
#' @param sheet the excel sheet as data.frame - read in with readxl::read_excel
#' @param indicator text in a cell to search for
#' @param off_n_rows the number of rows relative to the indicator from which to extract the information. Defaults to 0 which is suitable for all new grading forms.
#' @param off_n_cols the number of columns relative to the indicator from which to extract the information. Defaults to 2 which is suitable for all new grading forms.
#'
#' @return a string

get_info <- function(sheet,
                     indicator,
                     off_n_rows = 0,
                     off_n_cols = 2) {
  # FIND Index
  index <-
    which(sheet == indicator,
          arr.ind = TRUE
    ) # return row / col

  # FIND information
  # Use the offset and index to find the correct cell / value
  # Loop over all instances (e.g., 'Nakijkdatum' may occur multiple times if multiple tentamenkansen were used).
  info <-
    sapply(
      X = 1:nrow(index),
      FUN = function(i) {
        sheet[
          index[i, "row"] + off_n_rows,
          index[i, "col"] + off_n_cols
        ]
      }
    )


  # Clean
  info <- info[!is.na(info)]

  return(info)
}
