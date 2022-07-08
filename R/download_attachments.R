#' Download, Store, and Rename Attachments
#'
#' Each submission includes the relevant files as attachments. Usually at least
#' one document is included which follows a standardized naming convenstion:
#' course_run // type of document (e.g. 'verslag') // student name // student number
#' // filename as provided by the student. // file extension.
#' We save the files in the designated student folder and rename the file so that
#' ... only the standardized part of the filename is retained.
#'
#' @param email an email environment (ms_outlook_email) with the contents of a
#' single email sent by `submit@oupsy.nl`
#' @param student_folder a file path to the designated student folder where the
#' attachments should be saved.
#'
#' @return nothing
#'
download_attachments <- function(email,
                                 student_folder) {

  # List the available attachments
  all_attachments <- email$list_attachments()

  # Apply to all attachments
  lapply(
    X = all_attachments,
    FUN = function(attachment) {

      # Get attachment name
      attachment_name <- attachment$properties$name

      # Download the attachment with the original name
      # NOTE: somehow this function does not allow storing in a destination
      # ... folder unless the specified name is the EXACT copy of the
      # ... file that is downloaded
      # NOTE: does not return an object
      attachment$download(
        dest = here::here(
          student_folder,
          attachment_name
        ),
        overwrite = TRUE
      )

      # Rename the current file
      # All files have the standardized structure followed by the students'
      # ... file name included in [...]. We remove the students' file name
      # ... as it will not follow a standardized format.
      file.rename(
        from = here::here(
          student_folder,
          attachment_name
        ),
        to = here::here(
          student_folder,
          stringr::str_remove(
            string = attachment_name,
            pattern = " \\[(.*?)\\]"
          )
        )
      )
    }
  )
}
