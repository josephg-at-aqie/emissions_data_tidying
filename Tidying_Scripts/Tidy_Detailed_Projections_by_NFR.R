# Load packages ------------

packages <- c("tidyverse", "readxl", "janitor", "purrr")

# Check if packages needed are installed, if they are then load, if not install
for (package in packages) {
    if (!require(package, character.only = TRUE)) {
        install.packages(package)
        library(package, character.only = TRUE)
    }
}

# Define variables and function ------------

file_paths <- list.files(path = "Prepared_Data/Detailed_Projections", pattern = "*.xlsx", full.names = TRUE)
sheet_name <- "ProjectionsByNFR"

# Function to read and combine all the pivot tables from one xlsx file
read_projections_table <- function(file_path) {
    readxl::read_excel(file_path, sheet = "ProjectionsByNFR", skip = 5)
}

# Load and format data ------------

# This code reads the sheets from all .xlsx files in Prepared_Data/Pivot_Tables and combines them.
# The data is then converted to a longer formatted, and the values and column names are tidied.
projections_data <- map_dfr(file_paths, read_projections_table, .id = "Year_Index") |>
    # Convert the index variable for the different files into a variable storing the reporting year in the filename.
    mutate(Reporting_Year = as.numeric(gsub("\\D", "", basename(file_paths[as.numeric(Year_Index)])))) |>
    # Remove the index variables.
    select(Reporting_Year, everything(), -Year_Index) |>
    # No subtotal rows to remove
    # filter(!grepl("Total", Pollutant, ignore.case = TRUE)) |>
    # Pivot longer (combine the emissions from all years into one column)
    pivot_longer(
        cols = matches("^\\d{4}$"), # Matches any column who's title is a string of 4 numbers
        names_to = "year",
        values_to = "emission",
        values_drop_na = TRUE
    ) |>
    mutate(year = as.numeric(year)) |>
    # # Tidy column names
    janitor::clean_names() |>
    # Standardise pollutant names
    mutate(pollutant = case_when(
        pollutant == "NOx" ~ "Nitrogen Oxides as NO2",
        pollutant == "VOC" ~ "Non Methane VOC",
        pollutant == "NH3" ~ "Ammonia",
        pollutant == "SO2" ~ "Sulphur Dioxide",
        TRUE ~ pollutant # Keep everything else unchanged
    )) |>
    # Standardise units
    mutate(emission_unit = case_when(
        emission_unit == "kilotonnes" ~ "kilotonne",
        emission_unit == "Ktonne" ~ "kilotonne",
        emission_unit == "ktonnes" ~ "kilotonne",
        TRUE ~ emission_unit # Keep everything else unchanged
    )) |>
    # Add spaces to source_sheet which will be converted to pollutant group
    mutate(pollutant_group = case_when(
        pollutant %in% c("PM2.5", "PM10", "Black Carbon") ~ "Particulate Matter",
        pollutant %in% c("Ammonia", "Nitrogen Oxides as NO2", "Sulphur Dioxide", "Non Methane VOC") ~ "Air Pollutants"
    )) |>
    # Rename columns for clarity
    rename(
        nfr = nfr_code,
        unit = emission_unit,
    ) |>
    select(reporting_year, pollutant_group, everything())

# Add compliance totals ------------

# NO2 - have to calculate no2_compliance to exclude agriculture
no2 <- subset(projections_data, projections_data$pollutant == "Nitrogen Oxides as NO2")

no2_comp <- no2 %>%
    filter(!str_detect(nfr, "^3B|^3D")) %>%
    mutate(pollutant = "Nitrogen Oxides as NO2 NECR")

# VOC's - have to calculate voc_compliance to exclude agriculture
voc <- subset(projections_data, projections_data$pollutant == "Non Methane VOC")

voc_comp <- voc |>
    filter(!str_detect(nfr, "^3B|^3D")) |>
    mutate(pollutant = "Non Methane VOC NECR")

# NH3
nh3 <- subset(projections_data, projections_data$pollutant == "Ammonia")

# NH3 NECR Compliance Total (adjusted total)
# We only have the projections by NFR so we can't calculate the adjustment
# (i.e. by excluding TAN - Digestates)
# Instead we read in the adjustments which are copied into a CSV from the projections annex IV template
nh3_adjustments <- read.csv("Prepared_Data/Other/projections_ammonia_adjustments.csv") |> mutate(nfr = "3Da2c")
nh3_comp <- nh3 |>
    left_join(nh3_adjustments, by = c("reporting_year", "nfr", "year")) |>
    mutate(
        adjustment = replace_na(adjustment, 0), # Handle missing adjustments
        emission = emission + adjustment
    ) |>
    select(-adjustment) |> # Remove adjustment column
    mutate(pollutant = "Ammonia NECR")

projections_data <- rbind(projections_data, no2_comp, voc_comp, nh3_comp)


# QA Checks ------------

# This section provides a number of QA checks to be performed on the data.

# Check data structure and column types
str(projections_data)

# Check for missing data
colSums(is.na(projections_data))
sum(complete.cases(projections_data)) == nrow(projections_data)
projections_data |> filter(is.na(emission))

# Check for duplicates
projections_data |> filter(duplicated(projections_data))

# Check text columns for errors such as typos
unique(projections_data$pollutant_group)
unique(projections_data$pollutant)
unique(projections_data$unit)

# Check for negative values in emissions
projections_data |> filter(emission < 0)

# Check years are consistent
range(projections_data$year)
projections_data |> filter(year < 1900 | year > 2050)

# Check units are consistent
table(projections_data$unit)

# Check missing reporting_year-year pairs
all_combinations <- expand.grid(
    reporting_year = unique(projections_data$reporting_year),
    year = unique(projections_data$year)
)
missing_combinations <- anti_join(all_combinations, projections_data, by = c("reporting_year", "year"))
view(missing_combinations)
# Check missing pollutant-reporting_year pairs
all_combinations <- expand.grid(
    reporting_year = unique(projections_data$reporting_year),
    pollutant = unique(projections_data$pollutant)
)
missing_combinations <- anti_join(all_combinations, projections_data, by = c("reporting_year", "pollutant"))
view(missing_combinations)
# As expected, we didn't always report for all projection years or for all pollutants

# Check for extreme values in emissions
yearly_totals <- projections_data |>
    group_by(year, pollutant) |>
    summarise(total_emission = sum(emission, na.rm = TRUE), .groups = "drop")
extreme_values <- yearly_totals |>
    group_by(pollutant) |>
    summarise(
        min_total = min(total_emission, na.rm = TRUE),
        q1 = quantile(total_emission, 0.25, na.rm = TRUE),
        median = median(total_emission, na.rm = TRUE),
        q3 = quantile(total_emission, 0.75, na.rm = TRUE),
        max_total = max(total_emission, na.rm = TRUE)
    ) |>
    mutate(
        iqr = q3 - q1,
        lower_bound = q1 - 1.5 * iqr,
        upper_bound = q3 + 1.5 * iqr
    )
outliers <- yearly_totals |>
    left_join(extreme_values, by = "pollutant") |>
    filter(total_emission < lower_bound | total_emission > upper_bound) |>
    select(year, pollutant, total_emission)
print(outliers)
# It looks like there are quite are no outliers based on just the projections

# Plot time series for each pollutant (sum of all reporting years)
# Check if ggforce is installed
if (!require("ggforce", character.only = TRUE)) {
    install.packages("ggforce")
}
yearly_totals <- projections_data |>
    group_by(year, pollutant, reporting_year) |>
    summarise(total_emission = sum(emission, na.rm = TRUE), .groups = "drop")
num_pollutants <- length(unique(yearly_totals$pollutant))
num_pages <- ceiling(num_pollutants / 4) # 4 per page
for (page in 1:num_pages) {
    p <- ggplot(
        yearly_totals,
        aes(
            x = year,
            y = total_emission,
            group = reporting_year,
            color = as.factor(reporting_year)
        )
    ) +
        geom_line() +
        geom_point() +
        ggforce::facet_wrap_paginate(~pollutant, scales = "free_y", ncol = 2, nrow = 2, page = page) +
        labs(
            title = paste("Total Emissions Over Time by Pollutant (Page", page, "of", num_pages, ")"),
            x = "Year", y = "Total Emissions",
            color = "Reporting Year"
        ) +
        theme_minimal()

    print(p) # Display the plot
    Sys.sleep(0.5) # Small delay to help with rendering in some cases
}


# Save data ------------

# Split and save each reporting_year separately
split(projections_data, projections_data$reporting_year) |>
    purrr::iwalk(~ write.csv(
        .x,
        file = paste0("Tidied_Data/Projected_Emissions_by_NFR/projected_emissions_by_NFR_reporting_", .y, ".csv"),
        row.names = FALSE
    ))
