#' Update the Submission Overview
#'
#' Add the most recent submission
#' Restructure such that the overview shows all 'TO DO''s at the top,
#' with the shortest 'grade_before_date' taking priority
#'
#' @param submission_info information from the most recent submission
#' @param submission_overview file path - local to the `grading` project -
#' to the submission overview excel file.
#'
update_submission_overview <- function(submission_info,
                                       submission_overview) {

  #### LOAD CURRENT OVERVIEW ####
  # Load the existing submission_overview or create one if it doesn't yet exist.
  if (file.exists(here::here(submission_overview))) {
    overview <-
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
      tibble::as_tibble(.)
  } else {
    # No overview exists, initialize empty tibble
    overview <-
      # add current info for correct column names and column types
      submission_info %>%
      # remove data (will be added later) but keep structure
      dplyr::slice(-1)
  } # END if overview exists.

  #### UPDATE OVERVIEW ####
  overview <-
    # Add most recent submission
    tibble::add_row(
      .data = overview,
      submission_info
    ) %>%
    # RESTRUCTURE
    dplyr::arrange(
      dplyr::desc(status), # statuses: "To Do", "Feedback", "Herkansing", "Done"
      ifelse(status != "To Do",
        dplyr::desc(grade_before_date), # closest to-do date at the top of the sheet.
        grade_before_date
      ) # most recent completed at the top
    )

  #### INITIALIZE EMPTY EXCEL ####
  # Create empty xlsx
  xlsx_overview <- openxlsx::createWorkbook()

  # Change font
  openxlsx::modifyBaseFont(
    wb = xlsx_overview,
    fontName = "Avenir Next LT Pro"
  )
  # Add sheet
  openxlsx::addWorksheet(
    wb = xlsx_overview,
    sheetName = "Overview",
    gridLines = FALSE
  )
  # Freeze header
  openxlsx::freezePane(
    wb = xlsx_overview,
    sheet = "Overview",
    firstRow = TRUE,
    firstCol = FALSE
  )

  # Set date format to 01 Feb 2022
  options(openxlsx.dateFormat = "dd mmm yyyy")

  # Set default column sizes
  # Note: manual changes will not take effect the next time a new entry is made
  # ... you will need to update these values.
  # Column widths are in excel dimensions (visible when changing the width of the column)
  openxlsx::setColWidths(
    wb = xlsx_overview,
    sheet = "Overview",
    cols = 1:ncol(overview),
    c(
      7.3, # status
      10.5, # course_id
      18.1, # course_name (with overflow)
      13.7, # course_run
      16.8, # student_number
      14.8, # student_name (with overflow)
      15.2, # student_email (with overflow)
      16.5, # submission_date
      25, # student_submission_note (with overflow)
      29.3 # grade_before_date
    )
  )
  #### STYLING ####
  # Specify the styling of the xlsx file.

  # Custom header row
  header_styling <-
    openxlsx::createStyle(
      fontColour = "white",
      numFmt = "TEXT", # cell class
      border = "bottom",
      borderColour = "#002060", # dark blue,
      borderStyle = "thin",
      fgFill = "#002060", # dark blue background
      halign = "left", # align left to reduce space when adding the filter option later on
      valign = "center",
      textDecoration = "bold"
    )

  # center all other text
  text_styling <-
    openxlsx::createStyle(
      halign = "center",
      valign = "center",
      wrapText = TRUE,
      border = "TopBottom",
      borderColour = "#5F9F9F" # cadet blue
    )

  openxlsx::addStyle(
    wb = xlsx_overview,
    sheet = "Overview",
    style = text_styling,
    gridExpand = TRUE,
    cols = 1:ncol(overview),
    rows = 2:(nrow(overview) + 1) # +1 for start at row 2
  )

  #### Conditional formatting ####
  # NOTE: currently throws an error where the file cannot be automatically saved
  # ... and the conditional formatting is removed.

  # # TO DO
  # to_do_styling <-
  #   openxlsx::createStyle(
  #     fontColour = "#9C0006", # dark red
  #     bgFill = "#FFC7CE" # light red
  #   )
  #
  # openxlsx::conditionalFormatting(wb = xlsx_overview,
  #                       sheet = "Overview",
  #                       cols = "status",
  #                       rows = 1:nrow(overview)+1,
  #                       rule = "=='To Do'",
  #                       style = to_do_styling
  # )
  #
  # # DONE
  # done_styling <-
  #   openxlsx::createStyle(fontColour = "#006100", # dark green
  #               bgFill = "#C6EFCE" # light green
  #               )
  #
  # # FEEDBACK & HERKANSING
  # feedback_herkansing <-
  #   openxlsx::createStyle(fontColour = "#BF8F00", # yellow/brownish
  #               bgFill = "#FFE699" # light yellow
  #   )

  #### UPDATE XLSX ####
  openxlsx::writeData(
    wb = xlsx_overview,
    sheet = "Overview",
    headerStyle = header_styling,
    x = overview,
    startCol = 1,
    startRow = 1,
    colNames = TRUE,
    rowNames = FALSE,
    withFilter = TRUE,
    keepNA = FALSE # show NA as empty cells
  )

  #### ADD FORMULA ####
  # # Add a formula that computes the number of working days before a grade is due
  # NOTE: currently throws an error where the file needs to be restored.
  # ... may be due to the use of a dutch formula.

  # grade_before_days <- "NETTO.WERKDAGEN(VANDAAG();J2)"
  # openxlsx::writeFormula(wb = xlsx_overview,
  #                        sheet = "Overview",
  #                        x = grade_before_days,
  #                        startCol = ncol(overview) + 1, # last column
  #                        startRow = 2) # skip header row
  #

  #### SAVE XLSX ####
  openxlsx::saveWorkbook(
    wb = xlsx_overview,
    file = here::here(submission_overview),
    overwrite = TRUE
  )
}
