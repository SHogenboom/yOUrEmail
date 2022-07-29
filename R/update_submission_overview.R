#' Update the Submission Overview
#'
#' Add the most recent submission
#' Restructure such that the overview shows all 'TO DO''s at the top,
#' with the shortest 'grade_before_date' taking priority
#'
#' @param submission_info information from the most recent submission
#' @param submission_overview file path - local to the `grading` project -
#' to the submission overview excel file.
#'
update_submission_overview <- function(submission_info,
                                       submission_overview) {

  #### LOAD CURRENT OVERVIEW ####
  # Load the existing submission_overview or create one if it doesn't yet exist.
  if (file.exists(here::here(submission_overview))) {
    overview <-
      # load xlsx
      # NOTE: only works when file is closed!
      openxlsx::read.xlsx(
        xlsxFile = here::here(submission_overview),
        colNames = TRUE, # import the header row as column names
        detectDates = TRUE, # import dates
        skipEmptyRows = TRUE, # do not import empty cells
        skipEmptyCols = TRUE
      ) %>%
      # Convert to tibble for further processing
      tibble::as_tibble(.)
  } else {
    # No overview exists, initialize empty tibble
    overview <-
      # add current info for correct column names and column types
      submission_info %>%
      # remove data (will be added later) but keep structure
      dplyr::slice(-1)
  } # END if overview exists.

  #### UPDATE OVERVIEW ####
  overview <-
    # Add most recent submission
    tibble::add_row(
      .data = overview,
      submission_info
    ) %>%
    # RESTRUCTURE
    dplyr::arrange(
      dplyr::desc(status), # statuses: "To Do", "Feedback", "Herkansing", "Done"
      ifelse(status != "To Do",
        dplyr::desc(grade_before_date), # closest to-do date at the top of the sheet.
        grade_before_date
      ) # most recent completed at the top
    )


  #### SAVE XLSX ####
  # Initialize xlsx with correct styling
  xlsx_overview <- initalize_submission_xlsx(overview)

  openxlsx::saveWorkbook(
    wb = xlsx_overview,
    file = here::here(submission_overview),
    overwrite = TRUE
  )
}
