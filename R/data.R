#' Course Codes
#'
#' An overview of the course codes, course runs, and course names
#' PB: Psychology Bachelor
#' PM: Psychology Master
#' S: unknown but subjects contains psychology in the course name
#'
#' @author Open Universiteit - InfoHub
#'
#' @format a tibble
#' \describe{
#'   \item{course}{A combination of the CursusCode and a human readable name for each course}
#'   \item{course_name}{Only the human readable name of the course}
#'   \item{course_id}{2 characters + integers: a unique identifier per course.
#'   Letter prefixes indicate the program to which the course belongs.}
#'   \item{course_run}{2 characters + integers: a unique identifier per course run.
#'   A course may run multiple times per year with small changes to the contents.}
#' }
#'
"course_codes"
