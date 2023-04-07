#' Manually Add a Submission
#'
#' Quickly add a new submission (with the correct information) from a manual input.
#' This is relevant in the case of a re-take, or when the final assignments
#' ... are submitted directly via email.
#'
#' @param submission_overview file path - local to the `grading` project -
#' to the submission overview excel file.
#' @param student_id numeric - student_number for the manual entry. All corresponding information (name, email etc.) will be added from the previous submissions.
#' @param submission_date character with the submission date in format
#' "dd/mm/yyyy"
#' @param n_workdays defaults to 20; number of working days until the grading
#' is due.
#' @assingment character; name of the assignment (e.g., "A" or "Final")
#'
add_manual_submission <- function(student_id,
                                  submission_overview,
                                  submission_date = lubridate::today(),
                                  n_workdays = 20,
                                  assignment) {

  # LOAD current submission overview
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

  # FIND INFORMATION
  submission <-
    overview %>%
    dplyr::filter(student_number == student_id) %>%
    # Keep only one row in case of multiple retakes
    dplyr::slice_head(n = 1)

  # UPDATE INFORMATION
  # Reset status
  submission$status <- "To Do"
  # Update the submission date
  submission$submission_date <-
    lubridate::as_date(submission_date,
                       format = "%d/%m/%Y")
  # Recompute grade_before_date
  submission$grade_before_date <-
    bizdays::offset(submission$submission_date,
                    n = n_workdays,
                    cal = bizdays::create.calendar(name = "work_calender",
                                                   weekdays = c("saturday", "sunday")))
  # Update
  submission$assignment <- assignment
  # Reset to empty date
  submission$graded_on <- NA

  # Reset grade
  submission$grade <- NA

  # ADD TO OVERVIEW
  overview <-
    dplyr::bind_rows(overview,
                     submission) %>%
    # RESTRUCTURE
    dplyr::arrange(
      dplyr::desc(status), # statuses: "To Do", "Feedback", "Herkansing", "Done"
      ifelse(status != "To Do",
             dplyr::desc(grade_before_date), # closest to-do date at the top of the sheet.
             grade_before_date
      ) # most recent completed at the top
    )


  #### UPDATE XLSX ####
  # Initialize xlsx with correct styling
  xlsx_overview <- initalize_submission_xlsx(overview)

  openxlsx::saveWorkbook(
    wb = xlsx_overview,
    file = here::here(submission_overview),
    overwrite = TRUE
  )

  print(glue::glue("Uiterlijk Nakijken Voor: {submission$grade_before_date}"))

}
