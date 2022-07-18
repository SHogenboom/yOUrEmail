#' Get Information from Submission Email
#'
#' Parse the submission email for relevant data (e.g., student's name and id)
#'
#' @param email an email environment (ms_outlook_email) with the contents of a
#' single email sent by `submit@oupsy.nl`
#' @param n_workdays default: 20. The number of working days allowed to grade
#' the received submission.
#'
#' @return a tibble with one row containing all the relevant information from the email
#'
get_submission_info <- function(email,
                                n_workdays = 20) {

  # Initialize output
  dat <- tibble::tibble(.rows = 1)

  #### GRADE STATUS ####
  # Add the default status 'to do' for later use in the grading overview
  dat %<>%
    tibble::add_column("status" = "To Do")

  # Process Email-Subject
  # The subject contains information on:
  # The course / course run
  # The Students' Name as known at the OU
  # The Student Number

  subject <-
    # Get subject
    email$properties$subject %>%
    # Clean non-relevant information
    stringr::str_remove_all(pattern = "Het|verslag|van|id|\\(|\\)|:|,|A ") %>%
    # Remove redundant whitespaces
    stringr::str_squish()

  #### COURSE INFO ####
  # Load course data from package
  data("course_codes")

  # Identify the course:
  course <-
    stringr::str_extract(
      string = subject,
      # Two characters (e.g., PB) followed by 1 or more numbers
      pattern = "[A-Za-z]{2}[0-9]{1,}"
    ) %>%
    # Convert any lowercase letters to uppercase for consistency
    stringr::str_to_upper(.)

  # The submission system sometimes includes only the course_id and sometimes
  # the course_run.
  if (course %in% course_codes$course_id) {
    # The submission contains only the course_id
    # It is not possible to know the course run because students can submit even
    # after courses have finished.
    dat %<>%
      tibble::add_column(
        "course_id" = course,
        "course_name" = course_codes %>%
          dplyr::filter(course_id == course) %>%
          dplyr::select(course_name) %>%
          unique() %>%
          dplyr::pull(),
        "course_run" = "*",
      )
  } else if (stringr::str_c(course, "B") %in% course_codes$course_run) {
    # Add the 'B' to the course_runs as logged by us
    # The submission contains the course_run
    dat %<>%
      tibble::add_column(
        "course_id" = course_codes %>%
          dplyr::filter(course_run == stringr::str_c(course, "B")) %>%
          dplyr::select(course_id) %>%
          unique() %>%
          dplyr::pull(),
        "course_name" = course_codes %>%
          dplyr::filter(course_run == stringr::str_c(course, "B")) %>%
          dplyr::select(course_name) %>%
          unique() %>%
          dplyr::pull(),
        "course_run" = course,
      )
  } else {
    # If course info is not known add the available information
    # Course IDs always consist of 2 letters followed by 4 numbers (total 6 characters)
    dat %<>%
      tibble::add_column(
        "course_id" = stringr::str_sub(course, 1, 6),
        "course_name" = "*",
        "course_run" = course,
      )
  } # END IF

  #### STUDENT NUMBER ####
  # A sequence of 9 numbers
  student_number <- stringr::str_extract(
    string = subject,
    pattern = "[0-9]{1,9}$"
  )

  # Save to output
  dat %<>%
    tibble::add_column("student_number" = as.numeric(student_number))

  #### NAME FROM EMAIL ####
  # The students name as attached to the email address
  dat %<>%
    tibble::add_column(
      "student_name" =
        email$
          properties$
          ccRecipients[[1]]$
          emailAddress$
          name
    )

  #### EMAIL ADDRESS ####
  # The students email address
  dat %<>%
    tibble::add_column(
      "student_email" =
        email$
          properties$
          ccRecipients[[1]]$
          emailAddress$
          address
    )

  #### SUBMISSION DATE ####
  dat %<>%
    tibble::add_column(
      "submission_date" =
        lubridate::as_date(
          email$
            properties$
            createdDateTime
        )
    )

  #### SUBMISSION NOTE ####
  # Students have the option to add a note to the automatically generated email
  # If such a note is enclosed it is presented between two quotation marks "..."
  submission_note <-
    email$
    properties$
    body$
    content %>%
    qdapRegex::rm_between(
      left = '"',
      right = '"',
      extract = TRUE
    ) %>%
    unlist() %>%
    # remove linebreaks
    stringr::str_remove_all(pattern = "[\r\n]")


  dat %<>%
    tibble::add_column("student_submission_note" = ifelse(is.na(submission_note),
      "no student note",
      submission_note
    ))

  #### GRADE BEFORE ####
  # Manually compute the date before which the submission should be graded.
  # Defaults to 20 working days from the submission date.
  dat %<>%
    tibble::add_column("grade_before_date" = .$submission_date + lubridate::wday(n_workdays))

  #### MANUAL COLUMN ####
  # Initialize for manual completion
  dat %<>%
    tibble::add_column("assignment" = "#") %>%
    tibble::add_column("graded_on" = lubridate::today()) %>%
    tibble::add_column("grade" = numeric(1)) %>%
    tibble::add_column("grading_notes" = character(1))

  return(dat)
}
