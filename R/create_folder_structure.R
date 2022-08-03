#' Create Standardized Folder Structure for Submissions
#'
#' Submissions include the relevant files as attachments to the emails. We want
#' to store them according to a standardized file structure. Making it easy for
#' teachers to access the files when they are ready to grade a new submissions.
#'
#' @param submission_info information of the submission which should contain at
#' least the course_id, course_run, student_name, and student_number
#' @param grading_folder path to the top-level grading folder. The entire folder
#'  structure within this folder is created automatically.
#'
#' @return path to student folder

create_folder_structure <- function(submission_info,
                                    grading_folder) {

  #### COURSE ####
  course_folder <- here::here(
    grading_folder,
    submission_info$course_id
  )

  # Check existence of the course folder,
  # ... create it if it does not exist
  if (!dir.exists(course_folder)) {
    dir.create(course_folder)
  }

  #### COURSE RUN ####
  # Check if a course run is available, for some this is not the case then no
  # ... course run folder should be created.
  if (submission_info$course_run != "") {
    course_run_folder <- here::here(
      course_folder,
      submission_info$course_run
    )

    # Some courses do not have different course runs,
    # if the course_id is the same as the course_run we do not need the extra
    # ... nested folder structure

    if (submission_info$course_id != submission_info$course_run) {
      # Check existence of the course folder,
      # ... create it if it does not exist
      if (!dir.exists(course_run_folder)) {
        dir.create(course_run_folder)
      } # END if dir.exists
    } # END if course_id/course_run
  } else {
    course_run_folder <- course_folder
  } # END if course_run


  #### STUDENT ####
  student_folder <- here::here(
    course_run_folder,
    glue::glue_collapse(c(
      stringr::str_replace_all(submission_info$student_name,
        pattern = " ",
        replacement = "_"
      ),
      submission_info$student_number
    ),
    sep = "_"
    )
  )

  # Check existence of the course folder,
  # ... create it if it does not exist
  if (!dir.exists(student_folder)) {
    dir.create(student_folder)
  }

  return(student_folder)
}
