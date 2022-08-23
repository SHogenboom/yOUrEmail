save_examinatoren_overview <- function() {

  examinatoren <-
    dplyr::tibble() |>
    # Onderzoekspracticum Experimenteel Onderzoek
    dplyr::bind_rows(dplyr::tibble("course_id" = "PB0412",
                                   "examinator_name" = "Ron Pat-el",
                                   "examinator_email" = "ron.pat-el@ou.nl")) |>
    # Onderzoekspracticum Cross-Sectioneel Onderzoek
    dplyr::bind_rows(dplyr::tibble("course_id" = "PB0812",
                                   "examinator_name" = "Mira Duif & Gjalt-Jorn Peters",
                                   "examinator_email" = "oco-vragen@ou.nl")) |>
    # Onderzoekspracticum Literatuurstudie
    dplyr::bind_rows(dplyr::tibble("course_id" = "PB0712",
                                   "examinator_name" = "Jenny van Beek",
                                   "examinator_email" = "jenny.vanbeek@ou.nl"))

    saveRDS(examinatoren,
            here::here("data",
                       "examinatoren_overview.Rds"))

}

