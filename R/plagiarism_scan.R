#' Perform Plagiarism Scan
#'
#' Forward the Students' report to the Plagiarism scan software.
#'
#' @param email ms_outlook_email; the students' email (used to generate the subject)
#' @param student_folder the file path to where the attachments (including the report) were downloaded
#' @param urkund_email character; your personal urkund email
#'
forward_plagiarism_scan <- function(email,
                                    urkund_email,
                                    student_folder) {

  # Email is created with method chaining to add all relevant parts of the email
  plagiarism_email <-
    # Initialize an empty email
    outlook$create_email()$
    # Send to the Urkund plagiarism email
    set_recipients(to = urkund_email)$
    # Add subject
    set_subject(glue::glue("FW: {email$properties$subject}")) $
    # Add 'verslag'(ie., report) as attachment
    add_attachment(object = here::here(student_folder,
                                       # Find which filename in the student folder includes 'verslag'
                                       list.files(path = student_folder,
                                                  pattern = "verslag")))


  # Send the actual email
  plagiarism_email$send()

}
