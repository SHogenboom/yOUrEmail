#' Send a Confirmation Email
#'
#' Automatically generate a confirmation email to signal that
#' * The assignment has been received
#' * The assignment has been added to the `grading_overview`
#' * The student can expect a grade before the `grade_before_date`
#'
#' STORED IN DRAFTS
#'
#' @param outlook no default. Authenticate & retrieve your outlook environment
#' @param submission_info information from the email of a single submission
#' @param email an email environment (ms_outlook_email) with the contents of a
#' single email sent by `submit@oupsy.nl`
#' @param teacher_name the name of the teacher doing the grading
#' @param n_working_days number of working days before grading deadline
#'

draft_confirmation_email <- function(outlook,
                                    submission_info,
                                    email,
                                    teacher_name,
                                    n_working_days) {

  # Email is created with method chaining to add all relevant parts of the email
  confirmation_email <-
    # Access the email
    email$
      # Create a reply (draft)
      create_reply(send_now = FALSE)$
      # Add only the student as the recipient
      set_recipients(to = submission_info$student_email)$
      # Add the confirmation contents
      set_body(
      blastula::compose_email(
        body = blastula::md(
          glue::glue(
            "Beste {stringr::word(submission_info$student_name)},

            Hierbij bevestig ik de ontvangst van je opdracht voor {submission_info$course_id}.

            Rekening houdend met de maximale nakijktermijn van {n_working_days} werkdagen
            ontvang je uiterlijk **{format.Date(submission_info$grade_before_date,
            format = '%d %B %Y')}** feedback en/of je cijfer.

            Met vriendelijke groet,

            {teacher_name}
            "
          )
        )
      )
    )

  message(
  glue::glue("The email has been created as a DRAFT.",
             "The student provided the following note: \n",
             "{submission_info$student_submission_note}",
             "\n\n",
             "Please go to your email, confirm all information is correct, and press send!"
             ))

}
