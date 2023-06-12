#' Process all new assessments
#'
#' @param assessment_emails all emails that require processing
#' @param stored_assessments_folder filepath to where the downloaded files should be stored
#' @param processed_assessments_folder name of outlook folder to where the processed emails should be moved
#'
#' @return the info for the bulkforms

process_new_assessments <- function(assessment_emails,
                                    stored_assessments_folder,
                                    processed_assessments_folder) {

  # Initialize output
  all_new_assessments <- list()

  # LOOP OVER EMAILS
  for (email in assessment_emails) {
    #### CREATE TMP FOLDER ####
    # Create a temporary folder to store the attachments. The final name needs to be changed depending on the contents of the grade form.
    student_folder_tmp <-
      here::here(
        stored_assessments_folder,
        glue::glue("new_student")
      )

    # do not alert if folder already exists.
    dir.create(student_folder_tmp,
               showWarnings = FALSE
    )

    #### DOWNLOAD ATTACHMENTS ####
    invisible(
      yOUrEmail::download_attachments(email,
                                      student_folder = student_folder_tmp,
                                      clean_file_names = FALSE
      )
    )

    #### EXTRACT INFORMATION ####
    results <-
      yOUrEmail::get_assessment_information(student_folder_tmp)
    info <- results$info
    nth_tentamenkans <- results$nth_tentamenkans

    #### CHECK VERKLARING EIGEN WERK ####
    # List all downloaded files
    downloaded_files <- list.files(
      path = student_folder_tmp,
      recursive = TRUE
    )

    # Unlist any zipped files
    if (any(grepl(downloaded_files,
                  pattern = ".zip"))) {
      zipped_files <- utils::unzip(
        list.files(
          path = student_folder_tmp,
          pattern = ".zip",
          full.names = TRUE
        ),
        list = TRUE
      )
    }

    # Combine
    all_files <-
      c(
        downloaded_files,
        unlist(zipped_files$Name)
      )

    # CHECK for patterns
    if (any(grepl(all_files,
                  pattern = "VEW|Verklaring|verklaring|eigen|ethiek|wetgeving|wetenschappelijk|integriteit"
    ))) {
      info$VEW <- "Aanwezig"
    } else {
      info$VEW <- "AFWEZIG"
    }

    # Save output
    all_new_assessments[[length(all_new_assessments) + 1]] <- info

    #### CREATE STUDENT FOLDER ####
    # Create a new name for the target folder
    final_student_folder <- glue::glue("{info$student_id}-{info$student}")
    final_student_folder <- gsub(final_student_folder,
                                 pattern = " ",
                                 replacement = "-"
    )

    # Check if student folder already exists
    if (!dir.exists(here::here(
      stored_assessments_folder,
      final_student_folder
    ))) {
      # Create the new folder
      dir.create(path = here::here(
        stored_assessments_folder,
        final_student_folder
      ))
    }

    # Create a subfolder for the tentamenkans
    dir.create(path = here::here(
      stored_assessments_folder,
      final_student_folder,
      glue::glue("Tentamenkans-{nth_tentamenkans}")
    ))

    # Copy the files
    invisible(
      lapply(
        X = list.files(student_folder_tmp),
        FUN = function(filename) {
          file.copy(
            from = here::here(
              student_folder_tmp,
              filename
            ),
            to = here::here(
              stored_assessments_folder,
              final_student_folder,
              glue::glue("Tentamenkans-{nth_tentamenkans}")
            ),
            overwrite = TRUE,
            copy.date = TRUE
          )
        }
      )
    )

    # Remove the temporary folder
    # See: https://stackoverflow.com/questions/28097035/how-to-remove-a-directory-in-r
    unlink(here::here(student_folder_tmp), recursive = TRUE)

    #### MOVE EMAIL ####
    # Get the outlook instance of the folder
    processed_assessments_folder_outlook <-
      outlook$get_folder(processed_assessments_folder)

    # Move the email, no message
    invisible(email$move(processed_assessments_folder_outlook))

  }

  return(all_new_assessments)
}
