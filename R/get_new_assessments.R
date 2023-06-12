#' RETRIEVE NEW ASSESMENTS
#'
#' When teachers submit their assessment to the examinator, the examinator receives an email which contains the grading form. The information from the grading form needs to be parsed and added to an overview.
#'
#' @param outlook no default. Authenticate & retrieve your outlook environment
#' @param new_assessments_folder default: "Inbox". Specify the name of the folder (characters) in which the unprocessed assessments emails are stored. If you have an automatic rule setup which moves the submissions to a different folder (e.g., "Nakijken") specify the name of that folder.

#' @return a list of emails which all have the class "ms_outlook_email".
#'
get_new_assessments <-
  function(outlook,
           new_assessments_folder = "Inbox") {
    # Get new assessment emails
    new_assessments <-
      # Access the outlook environment
      outlook$
        # Get the assessment folder's content (defaults to inbox)
        get_folder(folder_name = new_assessments_folder)$
        # Retrieve all emails in the folder
        list_emails()

    return(new_assessments)
  }
