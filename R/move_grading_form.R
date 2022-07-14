#' Add the correct grading form to the students' folder
#'
#' Each course / course_run requires a unique grading form. Copy the correct grading
#' form to the students' submission folder. Rename the file to follow the same structure
#' as the received submission documents.
#'
#' @param grading_form_folder a file path - local to the `grading` project - where
#' all grading forms are stored. See the `workflow` document for the required file structure.
#' @param student_folder a file path - local to the `grading` project - where the students' files are stored.
#' @param submission_info information from the student's submission including `course`
#'  and `course_run` information
#'
#' @return nothing
#'
move_grading_form <- function(grading_form_folder,
                              student_folder,
                              submission_info) {

  # Not all courses contain course_run information in the submission details.
  if (is.na(submission_info$course_run)) {
    file.copy(from = here::here(grading_form_folder,
                                submission_info$course_id,
                                glue::glue("{submission_info$course_id}.xlsx")),
              to = here::here(student_folder,
                              glue::glue("{submission_info$course_id} beoordeling {submission_info$student_name} ({submission_info$student_number}).xlsx")
              ),
              copy.date = FALSE,
              overwrite = TRUE

    )
  } else {
    # Course contains a course_run specification which requires a distinct grading form
    file.copy(
      from = here::here(grading_form_folder,
                        submission_info$course_id,
                        glue::glue("{submission_info$course_run}.xlsx")),
      to = here::here(student_folder,
                      glue::glue("{submission_info$course_id} beoordeling {submission_info$student_name} ({submission_info$student_number}).xlsx")
      ),
      copy.date = FALSE,
      overwrite = TRUE
    )
  }
}
