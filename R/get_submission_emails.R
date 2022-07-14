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
#' @param outlook no default. Authenticate & retrieve your outlook environment
#' @param submissions_folder default: "Inbox". Specify the name of the folder (characters)
#' in which the submissions emails are stored. If you have an automatic rule setup
#' which moves the submissions to a different folder (e.g., "Nakijken") specify the name
#' of that folder.
#' @param submissions_email default: "submit@oupsy.nl". Specify the email address from
#' which the submissions are received.
#' @param email_status default: unread. Options include: "unread" (only new submissions),
#' "read" (only processed submissions), "all" (all submissions)
# .'
#' @return a list of emails which all have the class "ms_outlook_email".
#'
get_submissions <- function(outlook,
                            submissions_folder = "Inbox",
                            submissions_email = "submit@oupsy.nl",
                            email_status = "unread") {

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
          "from:{submissions_email}",
          dplyr::case_when(
            email_status == "all" ~ "",
            email_status == "read" ~ " AND isread:yes",
            email_status == "unread" ~ " AND isread:no"
          )
        )
    )

  return(submissions_from_system)
}
