#' Get submission information according to the 'Bulkschema'
#'
#' @param student_folder_tmp filepath to the student's temporary folder.

get_assessment_information <- function(student_folder_tmp) {
  #### GRADE FORM ###
  # ASUSMPTION: gradeform is the only file that ends in .xlsx or .xls
  grade_form_path <-
    list.files(student_folder_tmp,
               pattern = ".xlsx|.xls",
               recursive = TRUE
    )

  # LOAD gradeform into R memory
  grade_form <-
    # load xlsx
    # NOTE: only works when file is closed!
    openxlsx::read.xlsx(
      xlsxFile = here::here(
        student_folder_tmp,
        grade_form_path
      ),
      sheet = "Titelpagina",
      colNames = TRUE, # import the header row as column names
      detectDates = TRUE, # import dates
      skipEmptyRows = TRUE, # do not import empty cells
      skipEmptyCols = TRUE
    )

  #### DETERMINE N TENTAMENKANSEN ####
  # Determine from which 'Tentamenkans' to extract the information.
  # ASSUMPTION; the 'Deelopdracht' with the most filled grades = the number of 'Tentamenkansen'.
  nth_tentamenkans <-
    max(
      length(unique(yOUrEmail::get_info(grade_form, "Deelopdracht 1 (Thema 4)"))),
      length(unique(yOUrEmail::get_info(grade_form, "Deelopdracht 2 (Thema 5)"))),
      length(unique(yOUrEmail::get_info(grade_form, "Deelopdracht 3 (Thema 6)")))
    )

  #### GET INFO ####
  # Initialize output.
  # NOTE these column names correspond to the column names used in the 'Bulkformulier'
  all_info <-
    tibble(
      # Voorletters en achternaam student
      "student" = yOUrEmail::get_info(grade_form, "Naam student"),
      # Studentnummer
      "student_id" = yOUrEmail::get_info(grade_form, "Studentnummer"),
      # For some reason the form contains an empty column.
      "empty" = NA,
      # ontvangst- / inleverdatum (per tentamenkans)
      "inleverdatum" =
        yOUrEmail::get_info(grade_form, "Inleverdatum")[nth_tentamenkans],
      # Beoordelingsdatum (laatste)
      "beoordelingsdatum" =
        yOUrEmail::get_info(grade_form, "Nakijkdatum")[nth_tentamenkans],
      # Beoordelaar
      "Beoordelaar" = yOUrEmail::get_info(grade_form, "Beoordelaar"),
      # Deelopdracht 1: opdracht, casus, tentamen, notitie, mondeling etc.
      # One grade = first try,
      # multiple SAME grades = Do not submit again,
      # multiple DIFFERENT grades = 'Herkansing'
      "D1" = yOUrEmail::get_assignment_label(deelopdracht = 1, nth_tentamenkans, grade_form),
      # resultaat (cijfer)
      "D1_resultaat" = yOUrEmail::get_assignment_grade(deelopdracht = 1, nth_tentamenkans, grade_form),
      "D2" = yOUrEmail::get_assignment_label(deelopdracht = 2, nth_tentamenkans, grade_form),
      # resultaat (voldoende / onvoldoende)
      "D2_resultaat" = yOUrEmail::get_assignment_grade(deelopdracht = 2, nth_tentamenkans, grade_form),
      "D3" = yOUrEmail::get_assignment_label(deelopdracht = 3, nth_tentamenkans, grade_form),
      # resultaat (cijfer)
      "D3_resultaat" = yOUrEmail::get_assignment_grade(deelopdracht = 3, nth_tentamenkans, grade_form)
    )

  return(list(
    "info" = all_info,
    "nth_tentamenkans" = nth_tentamenkans
  ))
}
