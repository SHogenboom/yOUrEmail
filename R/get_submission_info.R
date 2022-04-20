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

  #### COURSE RUN ####
  # Identify the course run:
  course_run <-
    stringr::str_extract(
      string = subject,
      # Two characters (e.g., PB) followed by 1 or more numbers
      pattern = "[A-Za-z]{2}[0-9]{1,}"
    )

  # Save to output
  dat %<>%
    tibble::add_column("course_run" = course_run)

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
    tibble::add_column("submission_note" = submission_note)

  #### GRADE BEFORE ####
  # Manually compute the date before which the submission should be graded.
  # Defaults to 20 working days from the submission date.
  dat %<>%
    tibble::add_column("grade_before_date" = .$submission_date + lubridate::wday(n_workdays))

  #### GRADE STATUS ####
  # Add the default status 'to do' for later use in the grading overview
  dat %<>%
    tibble::add_column("grade_status" = "To Do")

  return(dat)
}
