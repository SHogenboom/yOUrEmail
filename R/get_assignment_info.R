#' Determine whether an assingment is a first submission, a retake (HERKANSING), or the same (copied) grade.
#'
#' @param deelopdracht numeric 1 - 3; the number of the deelopdracht to find the information of.

get_assignment_label <- function(deelopdracht, nth_tentamenkans, grade_form) {
  # Get the correct label (as is mentioned in the bulkformulieren)
  deelopdracht_label <-
    dplyr::case_when(
      deelopdracht == 1 ~ "Deelopdracht 1 (Thema 4)",
      deelopdracht == 2 ~ "Deelopdracht 2 (Thema 5)",
      deelopdracht == 3 ~ "Deelopdracht 3 (Thema 6)"
    )

  # Output the bulkformulier label
  assignment_label <-
    dplyr::case_when(
      # First grade for each Deelopdracht = submit all grades
      nth_tentamenkans == 1 ~ deelopdracht_label,
      # Grade of deelopdracht has not changed
      # = no need to resubmit (e.g., grade was already a 'pass')
      length(unique(yOUrEmail::get_info(
        grade_form,
        deelopdracht_label
      ))) <= 1 ~ NA,
      # All other cases (eg., new grade) = resubmit
      TRUE ~ paste0(deelopdracht_label, " HERKANSING")
    )

  return(assignment_label)
}


get_assignment_grade <- function(deelopdracht, nth_tentamenkans, grade_form) {
  # Get the correct label (as is mentioned in the bulkformulieren)
  deelopdracht_label <-
    dplyr::case_when(
      deelopdracht == 1 ~ "Deelopdracht 1 (Thema 4)",
      deelopdracht == 2 ~ "Deelopdracht 2 (Thema 5)",
      deelopdracht == 3 ~ "Deelopdracht 3 (Thema 6)"
    )

  # Get assignment label
  assignment_label <- get_assignment_label(deelopdracht, nth_tentamenkans, grade_form)

  # Only extract the assignment grade if the grade is new
  # (first submission or HERKANSING).
  assignment_grade <-
    dplyr::case_when(
      is.na(assignment_label) ~ NA,
      # use last to deal with herkansingen
      TRUE ~ last(yOUrEmail::get_info(grade_form, deelopdracht_label))
    )

  # Only deelopdracht 1 / 3 contain numeric grades
  if (deelopdracht %in% c(1, 3)) {
    assignment_grade <- round(as.numeric(assignment_grade), 1)
  }

  return(assignment_grade)
}
