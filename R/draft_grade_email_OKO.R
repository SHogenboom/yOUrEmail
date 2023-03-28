#' The appropriate email depends on:
#' ... the grade per deelopdracht

#' @param outlook no default. Authenticate & retrieve your outlook environment
#' @param submission_overview file path to the submission overview
#' @param grading_folder file path to where the student files/grades live.
#' @param student_id unique numeric indicator of a student
#' @param teacher_name the name of the teacher doing the grading

draft_grade_email_OKO <- function(outlook,
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

  # Find the folder - indendent of naming conventions - with the correct student_id
  student_folder <- all_student_folders[grepl(x = all_student_folders,
                                              pattern = student_id)]

  #### PARAMETERS ####
  course_id <- student_info$course_id

  if (course_id == "PB1612") {
    # Cross-Sectioneel
    examinator_name <- "Julia Fischmann & Gjalt-Jorn Peters"
    examinator_email <- "oko-vragen@ou.nl"
    pass <- 5.5
    compensation <- NA
  } else {
    stop("Course_id not coded")
  }

  #### ZIP ATTACHMENTS ####
  # NOTE: requires temporarily changing of working directory due to how .zip is
  # ... used in the Windows environment.
  old_wd <- getwd()
  setwd(here::here(student_folder))

  # Create a name for the zip (used later)
  zip_name <- glue::glue(
    "{course_id} beoordeling {student_info$student_name} ",
    "({student_info$student_number}).zip")

  # Zip the files
  utils::zip(
    # Name of the output zip
    zipfile = zip_name,
    # Files to include
    files = list.files(recursive = TRUE,
                       full.names = TRUE)
  )

  # Reset working directory
  setwd(old_wd)

  #### EMAIL CONTENT ####
  # Adress with first name - conform OU communication guidelines
  title <-
    glue::glue("Beste {stringr::word(student_info$student_name)},

          ")

  # GRADE
  # Bijv. "Je hebt het deeltentamen van PB012 met een 6.5 afgerond."

  # Multiple assignments, split the grades for each assignment
  grades <- strsplit(student_info$grade, split = "|", fixed = TRUE)[[1]]
  grade_consequences <-
    c(ifelse(as.numeric(grades[1]) >= pass, 'Gehaald', 'Herkansing'),
      ifelse(grades[2] %in% c('Voldoende', "voldoende"), 'Gehaald', 'Herkansing'),
      ifelse(as.numeric(grades[3]) >= pass, 'Gehaald', 'Herkansing'))

  grade_info <-
    glue::glue("
             Hierbij ontvang je de beoordeling voor de eindopdrachten van het {student_info$course_name} ",
               "({ifelse(student_info$course_run %in% c(NA, '*'),
             student_info$course_id,
             student_info$course_run)}): \n


             * *Deelopdracht 1 (Thema 4):* {grades[1]} ({grade_consequences[1]})
             * *Deelopdracht 2 (Thema 5):* {grades[2]} ({grade_consequences[2]})
             * *Deelopdracht 3 (Thema 6):* {grades[3]} ({grade_consequences[3]})
             \n
             ")

  # Pass/Fail/Compensation
  if (all(grade_consequences == "Gehaald")) {
    # Course was passed
    consequence <-
      glue::glue("**Gefeliciteerd - je bent voor alle deelopdrachten geslaagd!**")

  } else {
    # One of the deelopdrachten did not pass
    consequence <-
      glue::glue("**Helaas heb je èèn of meerdere deeltentamen(s) onvoldoende afgerond**.
               Dat betekent dat je de onvoldoende deelopdracht(en) moet herkansen.")
  }

  # EXPLANATIONS OF PROCEDURE
  procedure <-
    glue::glue("

             Een onderbouwing voor de beoordeling is te vinden in de beoordelingsrubric (zie bijlage).
             De opmerkingen die ik bij de beoordeling heb geplaatst zijn zowel kritisch als opbouwend van aard.
             Wanneer kritiekpunten buiten de rubric lijken te vallen kun je ze interpreteren als een extra service/uitleg.
             Er zullen dan ook geen punten voor afgetrokken zijn - maar ik hoop wel dat ze nuttig zullen zijn voor toekomstige verslagen / onderzoekspractica!

             *Dit is een voorlopige beoordeling*; ik cc daarom de examinator{ifelse(grepl(x = examinator_name, pattern = '&'), 'en', '')} ({examinator_name}) zodat de definitieve beoordeling opgesteld kan worden.
             Hoewel onze beoordelingen in de regel overeenkomen is er een kleine kans dat je definitieve beoordeling afwijkt.
             In dat geval geldt de beoordeling van de examinator. Als dit voor jou speelt, informeert de examinator je daar zo snel mogelijk over.
             Als je geen bericht ontvangt van de examinator kun je ervan uitgaan dat de definitieve beoordeling overeenkomt met mijn voorlopige beoordeling.
             De beoordeling is officieel als deze door het tentamenbureau is verwerkt en in Studiepad is bijgeschreven.
             *De verwerking door het tentamenbureau kan enkele weken duren*.

             ")

  # RETAKE OR EVALUATE
  if (all(grade_consequences == "Gehaald")) {
    # Course was passed
    follow_up <-
      glue::glue("In de cursusstructuur in yOUlearn staat helemaal aan het eind een cursusevaluatieformulier klaar.
              Het zou heel fijn zijn als je deze in vult - hiermee kunnen we onze cursussen blijven verbeteren.
              Alvast hartelijk dank voor de tijd en moeite!")
  } else {
    # Course not eligible for compensation OR grade too low
    follow_up <-
      glue::glue("De herkansing kun je straks direct aan mij mailen. Graag ontvang ik een versie waarin de wijzigingen zijn bijgehouden.
               Ik hoop dat je met mijn feedback in het beoordelingsformulier genoeg aanknopingspunten hebt voor verbetering.
               Als je tijdens de herkansing nog vragen hebt over de betekenis van de feedback - schroom dan niet om contact met mij op te nemen!
               Voor inhoudelijke vragen kun je ook altijd terecht op het discussieforum van de cursus of op [https://onderzoeksvragen.ou.nl/](https://onderzoeksvragen.ou.nl/).")
  }

  # Closure
  closing <-
    glue::glue("

             Met vriendelijke groet,

            {teacher_name}")

  # Request feedback from students
  ask_feedback <-
    glue::glue("

             PS: *Ik ontvang ook graag feedback op mijn feedback!*
             Is mijn feedback duidelijk verwoord - kun je er (in de toekomst) mee aan de slag?
             Is de feedback nuttig voor andere vakken en/of toekomstige opdrachten?")

  #### COMPILE EMAIL ####
  grade_email <-
    # Initialize
    outlook$create_email()$
    # Add recipients
    set_recipients(to = student_info$student_email,
                   cc = examinator_email)$
    # Add subject
    set_subject(glue::glue("Beoordeling {course_id}: {student_info$student_name} ",
                           "({student_info$student_number})"))$
    set_body(
      blastula::compose_email(
        body = blastula::md(
          c(
            title, # Beste {student_name}
            grade_info, # This is your grade
            consequence, # You have passed/failed/require compensation
            procedure, # Feedback in grading form; final grade by examinator
            follow_up, # evaluate cours or send retakes to teacher
            closing, # Met vriendelijke groet {teacher_name}
            ask_feedback # request feedback on feedback.
          )
        )
      )
    )

  # ADD ATTACHEMENTS
  # Everything combined in a zip & grade form separately for easy use by examinatoren
  grade_email <-
    grade_email$
    # Add grade form separately
    add_attachment(list.files(here::here(student_folder),
                              pattern = "[Bb]eoordeling.*\\.xlsx",
                              full.names = TRUE))$
    # Add zipped files
    add_attachment(here::here(student_folder,
                              zip_name))

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
  xlsx_overview <- initalize_submission_xlsx(overview)

  openxlsx::saveWorkbook(
    wb = xlsx_overview,
    file = here::here(submission_overview),
    overwrite = TRUE
  )


  warning("The email has been created as a DRAFT.
          Please go to your email, confirm all information is correct, and press send!")
}

