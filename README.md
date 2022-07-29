# yOUrEmail

The Teachers of the Psychology Department of the Open Universiteit (NLD) receive the to-be-graded papers via their work e-mail.

The purpose of this package is to deal with all manual steps outlined below.

## Current -Manual- Workflow

1. [`get_submission_emails`] Receive an E-mail from `submit@oupsy.nl` which contains:
    * The submission date; determines the deadline for grading (+20 working days)
    * The course / course version; determines which grading form should be used
    * Student Name; determines the 'aanhef' of the emails
    * Student Number; for easy registration
    * Student Email; determines to which e-mail adres the replies and grading are sent
    * Submission Notes; relevant notes as added by the student upon submission
    * The relevant files as attachments
1. [`email_confirmation`] Send the student a confirmation email that the submission has been received.
1. [`create_folder_structure`] Create a Course > Course Version > Student_ID sub-folder
1. [`download_attachments`] Download attachments
    * Save in the created folder
    * Clean-up file names to keep only the standardized structure
1. [`move_grading_firm`] Add the correct grading form to the folder
1. [`update_submission_overview`] Add a registration for the submission in the 'submission overview'
    * Allows a teacher to keep track of all open submissions & grading deadlines.
1. [`no function`] Mark email as completed
1. Grade the paper (NOTE: cannot be done automatically!)
1. [TO DO] Update the submission overview with the grade information, update status to 'DONE'
1. [TO DO] Compress all relevant files into a .zip folder
1. [TO DO] Send 'Beoordelings' Email
    * CC the examinator of the course
    * Add personalized grade information and pass/fail judgement
    * Include information on feedback, herkansingen, course evaluations
    * Add the .zip folder as attachment
    * Request feedback on feedback & filling in of the course evaluation.
