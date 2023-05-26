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
#' @return a list of emails which all have the class "ms_outlook_email".
#'
get_submissions <- function(outlook,
                            submissions_folder = "Inbox") {

  # Get submissions
  submissions_from_system <-
    # Access the outlook environment
    outlook$
    # Get the submissions folder's content (defaults to inbox)
    get_folder(folder_name = submissions_folder)$
    # Retrieve all emails in the folder
    list_emails()

  # Extract only the emails which are unread
  submissions <-
    lapply(X = submissions_from_system,
           FUN = function(em) {
             # Keep only new emails
             if (!em$properties$isRead) {
               return(em)
             } else {
               # Return a null object which is removed later
               return()
             }
           })

  # remove empty elements
  new_submissions <- purrr::compact(submissions)

  # Return only the new emails
  return(new_submissions)
}
