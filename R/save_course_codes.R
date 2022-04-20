#' Save Course Codes
#'
#' Create an overview of the available course codes which exist within the OU
#' administration.
#'
#' Parsed from an export received via InfoHub
#'
#' NOTE: the file itself is not available due to data protection guidelines
#' @param filePath a file path which points to the InfoHub export which lives locally
#' ... on ones computer (`FI_CursusDeelname`).
#' @return nothing
#'
save_course_codes <- function(filePath) {

  # read file contents
  course_codes <-
    readr::read_delim(
      file = filePath,
      delim = "#",
      escape_double = FALSE,
      trim_ws = TRUE,
      show_col_types = FALSE
    ) %>%
    # keep relevant columns
    dplyr::select(CursusCode, Cursus, Cursusrun) %>%
    # Keep only unique combinations
    unique() %>%
    # Keep only those relevant to the Psychology department
    # PB: Psychology Bachelor
    # PM: Psychology Master
    # S: unknown but subjects contains psychology in the course name
    dplyr::filter(stringr::str_detect(
      string = CursusCode,
      "^PB[0-9]|^PM[0-9]|^S[0-9]"
    )) %>%
    # rename columns for consistency
    dplyr::rename(
      "course" = "Cursus",
      "course_id" = "CursusCode",
      "course_run" = "Cursusrun"
    ) %>%
    # Extract course_names by removing the course codes
    dplyr::mutate("course_name" = stringr::str_remove_all(string = course,
                                                          pattern = "^(.*?) - "))


  # save contents for use in package
  usethis::use_data(course_codes,
    overwrite = TRUE
  )
}
