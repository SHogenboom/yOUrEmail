#' RETRIEVE NEW SUBMISSIONS
#'
#' When students submit their assignment to the submission system.
#' Both teachers and students receive an email from `submit@oupsy.nl` with
#' student details and the submitted documents.
#'
#' The submission system is only used for first-time submissions.
#' That means retakes and feedback rounds (e.g., Literatuurstudie) are not
#' handled by the submission system.
#'
#' @param submissions_folder default: "Inbox". Specify the name of the folder (characters)
#' in which the submissions emails are stored. If you have an automatic rule setup
#' which moves the submissions to a different folder (e.g., "Nakijken") specify the name
#' of that folder.
#' @param submissions_email default: "submit@oupsy.nl". Specify the email address from
#' which the submissions are received.
#' @param n_latest_emails default: 10. Specify the maximum number of new submissions
#' expected in a single day (N = 10). Increase this number so that
#' all new submissions will be retrieved but only a minimal number of 'non-relevant' emails
#' will be included.
#'
#' @return a list of emails which all have the class "ms_outlook_email".
#'
get_submissions <- function(submissions_folder = "Inbox",
                                        submissions_email = "submit@oupsy.nl",
                                        n_latest_emails = 10) {

  # Authenticate & get Outlook Environment
  outlook <- Microsoft365R::get_business_outlook()

  # Get submissions
  submissions_from_system <-
    # Access the outlook environment
    outlook$
      # Get the submissions folder's content (defaults to inbox)
      get_folder(folder_name = submissions_folder)$
      # Retrieve only the emails from the submission system
      list_emails(
      search =
        glue::glue(
          "from:{submissions_email}"
        ),
      # Retieve only the n last email to reduce processing time
      n = n_latest_emails
    )

  return(submissions_from_system)
}
