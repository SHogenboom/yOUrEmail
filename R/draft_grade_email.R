#' The appropriate email depends on:
#' ... the course (some course have grade compensations)
#' ... the grade

#' @param outlook no default. Authenticate & retrieve your outlook environment
#' @param submission_overview file path to the submission overview
#' @param grading_folder file path to where the student files/grades live.
#' @param student_id unique numeric indicator of a student
#' @param teacher_name the name of the teacher doing the grading


draft_grade_email <- function(outlook,
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
    # Keep last entry
    dplyr::slice(which.max(submission_date))

  # STUDENT FOLDER
  # List all student folders for a course
  all_student_folders <-
    list.dirs(path = here::here(grading_folder,
                                student_info$course_id,
                                ifelse(student_info$course_run %in% c(NA, "*"),
                                       "",
                                       course_run)),
              # prevent finding nested folders (e.g., tentamenkans 2)
              recursive = FALSE)

  # Find the folder - indendent of naming conventions - with the correct student_id
  student_folder <- all_student_folders[grepl(x = all_student_folders,
                                              pattern = student_id)]

  #### PARAMETERS ####
  course_id <- student_info$course_id

  if (course_id == "PB0812") {
    # Cross-Sectioneel
    examinator_name <- "Mira Duif & Gjalt-Jorn Peters"
    examinator_email <- "oco-vragen@ou.nl"
    pass <- 5.5
    compensation <- NA
  } else if (course_id == "PB0712") {
    # Literatuurstudie
    examinator_name <- "Jenny van Beek"
    examinator_email <- "jenny.vanbeek@ou.nl"
    pass <- 5.5
    compensation <- NA
  } else if (course_id == "PB0412") {
    # Experimenteel
    examinator_name <- "Ron Pat-El"
    examinator_email <- "ron.pat-el@ou.nl"
    pass <- 5.5
    compensation <- 5
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
  # Je hebt het deeltentamen van PB012 met een 6.5 afgerond.
  grade_info <-
    glue::glue("
             Je hebt een **{student_info$grade}** voor het deeltentamen van {student_info$course_name} ",
             "({ifelse(student_info$course_run %in% c(NA, '*'),
             student_info$course_id,
             student_info$course_run)})")

  # Pass/Fail/Compensation
  if (as.numeric(student_info$grade) > pass) {
    # Course was passed
    consequence <-
      glue::glue("**Gefeliciteerd - je bent daarom voor het deeltentamen geslaagd!**")

  } else if (!is.na(compensation) & as.numeric(student_info$grade) > compensation) {
    # Course is eligible for compensation
    consequence <-
      glue::glue("**Helaas heb je het deeltentamen daarmee met een onvoldoende afgerond.
               Mogelijk kun je deze onvoldoende nog wel compenseren met je eindtentamen (zie yOUlearn voor meer informatie).")

  } else {
    # Course not eligible for compensation OR grade too low
    consequence <-
      glue::glue("**Helaas heb je het deeltentamen daarmee met een onvoldoende afgerond.
               Dat betekent dat je deze opdracht moet herkansen.")
  }

  # EXPLANATIONS OF PROCEDURE
  procedure <-
    glue::glue("

             Een onderbouwing voor het behaalde cijfer is te vinden in de beoordelingsrubric (zie bijlage).
             De opmerkingen die ik bij de beoordeling heb geplaatst zijn zowel kritisch als opbouwend van aard.
             Wanneer kritiekpunten buiten de rubric lijken te vallen kun je ze interpreteren als een extra service/uitleg.
             Er zullen dan ook geen punten voor afgetrokken zijn - maar ik hoop wel dat ze nuttig zullen zijn voor toekomstige verslagen / onderzoekspractica!

             *Dit is een voorlopige beoordeling*; ik cc daarom de examinator ({examinator_name}) zodat de definitieve beoordeling opgesteld kan worden.
             Hoewel onze beoordelingen in de regel overeenkomen is er een kleine kans dat je definitieve beoordeling afwijkt.
             In dat geval geldt de beoordeling van de examinator. Als dit voor jou speelt, informeert de examinator je daar zo snel mogelijk over.
             Als je geen bericht ontvangt van de examinator kun je ervan uitgaan dat de definitieve beoordeling overeenkomt met mijn voorlopige beoordeling.
             De beoordeling is officieel als deze door het tentamenbureau is verwerkt en in Studiepad is bijgeschreven.
             *De verwerking door het tentamenbureau kan enkele weken duren*.

             ")

  # RETAKE OR EVALUATE
  if (as.numeric(student_info$grade) > pass) {
    # Course was passed
    follow_up <-
      glue::glue("In de cursusstructuur in yOUlearn staat helemaal aan het eind een cursusevaluatieformulier klaar.
              Het zou heel fijn zijn als je deze in vult - hiermee kunnen we onze cursussen blijven verbeteren.
              Alvast hartelijk dank voor de tijd en moeite!")

  } else if (!is.na(compensation) & as.numeric(student_info$grade) > compensation) {
    # Course is eligible for compensation
    follow_up <-
      glue::glue("Wanneer je de cursus voldoende hebt afgerond willen we je vragen om de cursusevaluatie in te vullen.
               Deze staat aan het eind van de cursusstructuur in yOUlearn.
               Het zou heel fijn zijn als je deze in vult - hiermee kunnen we onze cursussen blijven verbeteren.
               Alvast hartelijk dank voor de tijd en moeite!

               Moet je dit deeltentamen toch herkansen? Dan kun je de herkansing straks direct aan mij mailen.
               Ik hoop dat je met mijn feedback genoeg aanknopingspunten hebt voor verbetering.
               Als je tijdens de herkansing nog vragen hebt over de betekenis van de feedback - schroom dan niet om contact met mij op te nemen!
               Voor inhoudelijke vragen kun je ook altijd terecht op het discussieforum van de cursus en op [https://onderzoeksvragen.ou.nl/](https://onderzoeksvragen.ou.nl/).

               ")

  } else {
    # Course not eligible for compensation OR grade too low
    follow_up <-
      glue::glue("De herkansing kun je straks direct aan mij mailen.
               Ik hoop dat je met mijn feedback genoeg aanknopingspunten hebt voor verbetering.
               Als je tijdens de herkansing nog vragen hebt over de betekenis van de feedback - schroom dan niet om contact met mij op te nemen!
               Voor inhoudelijke vragen kun je ook altijd terecht op het discussieforum van de cursus en op [https://onderzoeksvragen.ou.nl/](https://onderzoeksvragen.ou.nl/).")
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
    add_attachment(here::here(student_folder,
                              zip_name))$
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

  warning("The email has been created as a DRAFT.
          Please go to your email, confirm all information is correct, and press send!")
}
