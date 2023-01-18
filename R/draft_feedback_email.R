#' A standardized email for sending feedback to students

#' @param outlook no default. Authenticate & retrieve your outlook environment
#' @param submission_overview file path to the submission overview
#' @param grading_folder file path to where the student files/grades live.
#' @param student_id unique numeric indicator of a student
#' @param teacher_name the name of the teacher doing the grading


draft_feedback_email <- function(outlook,
                                 submission_overview,
                                 grading_folder,
                                 student_id,
                                 teacher_name) {

  #### STUDENT INFO ####
  student_info <-
    # Load submission overview
    # load xlsx
    # NOTE: only works when file is closed!
    openxlsx::read.xlsx(
      xlsxFile = here::here(submission_overview),
      colNames = TRUE, # import the header row as column names
      detectDates = TRUE, # import dates
      skipEmptyRows = TRUE, # do not import empty cells
      skipEmptyCols = TRUE
    ) %>%
    # Convert to tibble for further processing
    tibble::as_tibble(.) %>%
    # Find student's information
    dplyr::filter(student_number == student_id) %>%
    # Keep only completed
    dplyr::filter(status == "Done") %>%
    # Keep last entry
    dplyr::slice(which.max(submission_date))

  # STUDENT FOLDER
  # List all student folders for a course
  all_student_folders <-
    list.dirs(path = here::here(grading_folder,
                                student_info$course_id,
                                ifelse(student_info$course_run %in% c(NA, "*"),
                                       "",
                                       trimws(student_info$course_run))),
              # prevent finding nested folders (e.g., tentamenkans 2)
              recursive = FALSE)

  # Find the folder - independent of naming conventions - with the correct student_id
  student_folder <- all_student_folders[grepl(x = all_student_folders,
                                              pattern = student_id)]

  #### COMPILE EMAIL ####
  feedback_email <-
    # Initialize
    outlook$create_email()$
    # Add recipients
    set_recipients(to = student_info$student_email)$
    # Add subject
    set_subject(glue::glue("Feedback {student_info$course_id}: {student_info$student_name} ",
                           "({student_info$student_number})"))$
    set_body(
      blastula::compose_email(
        body = blastula::md(
          glue::glue(
            "Beste {stringr::word(student_info$student_name)},

              Hierbij ontvang je feedback op je opdracht. Specifieke voorbeelden en
              opmerkingen kun je vinden in het document (*zie bijlage*).
              Het is niet haalbaar om herhaalde patronen (e.g., APA fouten, taalfouten)
              keer op keer in de tekst aan te geven. Zie de specifieke feedback
              dus echt als voorbeelden van aspecten die je misschien op meerdere
              plekken kunt voorkomen, moet aanpassen, of juist kunt herhalen.

              Heb je vragen over je feedback schroom dan niet om even een kort mailtje te sturen voor opheldering!
              Beter even verduidelijken dan dat je feedback verkeerd toepast.
              Wanneer je de opdracht straks afhebt kun je die direct naar mij opsturen.

              Succes met het afronden van de opdrachten,
              {teacher_name}

              PS: *Ik ontvang ook graag feedback op mijn feedback!*
              Is mijn feedback duidelijk verwoord - kun je er (in de toekomst) mee aan de slag?
              Is de feedback nuttig voor andere vakken en/of toekomstige opdrachten?"
          )
        )
      )
    )

  # ADD ATTACHEMENTS
  # Find all student files with info (e.g., created date etc)
  student_files <-
    file.info(
      list.files(here::here(student_folder),
                 full.names = TRUE,
                 recursive = TRUE))

  # Select most recent file
  recent_student_file <-
    rownames(student_files)[which.max(student_files$mtime)]

  feedback_email <-
    feedback_email$
    # Add grade form separately
    add_attachment(recent_student_file)

  #### UPDATE SUBMISSION OVERVIEW ####
  overview <-
    # Load submission overview
    # load xlsx
    # NOTE: only works when file is closed!
    openxlsx::read.xlsx(
      xlsxFile = here::here(submission_overview),
      colNames = TRUE, # import the header row as column names
      detectDates = TRUE, # import dates
      skipEmptyRows = TRUE, # do not import empty cells
      skipEmptyCols = TRUE
    ) %>%
    # RESTRUCTURE
    dplyr::arrange(
      dplyr::desc(status), # statuses: "To Do", "Feedback", "Herkansing", "Done"
      ifelse(status != "To Do",
             dplyr::desc(grade_before_date), # closest to-do date at the top of the sheet.
             grade_before_date
      ) # most recent completed at the top
    )

  #### UPDATE SUBMISSION OVERVIEW XLSX ####
  # Initialize xlsx with correct styling
  xlsx_overview <- yOUrEmail::initalize_submission_xlsx(overview)

  openxlsx::saveWorkbook(
    wb = xlsx_overview,
    file = here::here(submission_overview),
    overwrite = TRUE
  )


  warning("The email has been created as a DRAFT.
          Please go to your email, confirm all information is correct, and press send!")
}

