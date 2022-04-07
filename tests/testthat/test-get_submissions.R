#' Authentication

testthat::test_that("Outlook: Authentication", {
  # Authenticate & get Outlook Environment
  outlook <- Microsoft365R::get_business_outlook()

  # TO DO:
  # Expect ms_outlook environment class
})

#' Get Submissions
#'
#' Retrieve the submissions that were send by the submissions system
#' Function should return a list of relevant emails
testthat::test_that("Submissions: Retrieve from Outlook", {

  # TEST: Default values
  # submissions_folder = "Inbox" - folder always exists so no errors expected
  # submissions_email = submit@oupsy.nl" - should return 0 to inf results depending
  # ... on the setup of the users' Inbox
  # n_latest_emails = 10
  # Expectation: the function should succeed and return a list of 0 or more emails
  testthat::expect_length(get_submissions(),
                          10)

  # TEST: implausible submissions folder
  # Expectation: error due to inability to locate the folder
  testthat::expect_error(get_submissions(submissions_folder = "qwertyuiop"))

  # TEST: implausible submissions email
  # Expectation: success but an empty list of emails
  submission_emails <- get_submissions(submissions_email = "qwertyuiop@ou.nl")
  testthat::expect_true(length(submission_emails) == 0)

  # TEST: a different number of latest_emails
  # Expectation: list of length as specified
  testthat::expect_length(get_submissions(n_latest_emails = 20),
                          20)

  # TEST: a negative number of latest_emails
  # Expectation: it is not possible to have a list of negative length
  # ... the microsoft365r function returns a 'NULL'
  testthat::expect_null(get_submissions(n_latest_emails = -10))

})
