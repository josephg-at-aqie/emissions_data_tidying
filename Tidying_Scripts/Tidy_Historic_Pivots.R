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

file_paths <- list.files(path = "Prepared_Data/Pivot_Tables", pattern = "*.xlsx", full.names = TRUE)
sheet_names <- c("AirPollutants", "HeavyMetals", "POPs&PAHs", "ParticulateMatter")

# Function to read and combine all the pivot tables from one xlsx file
read_pivot_tables <- function(file_path) {
    purrr::map_dfr(
        sheet_names,
        ~ read_excel(file_path, sheet = .x, skip = 13) |> mutate(Source_Sheet = .x),
        .id = "Sheet_Index"
    )
}

# Load and format data ------------

# This code reads the sheets from all .xlsx files in Prepared_Data/Pivot_Tables and combines them.
# The data is then converted to a longer formatted, and the values and column names are tidied.
historic_data <- map_dfr(file_paths, read_pivot_tables, .id = "Year_Index") |>
    # Convert the index variable for the different files into a variable storing the reporting year in the filename.
    mutate(Reporting_Year = as.numeric(gsub("\\D", "", basename(file_paths[as.numeric(Year_Index)])))) |>
    # Remove the index variables.
    select(Reporting_Year, Source_Sheet, everything(), -Year_Index, -Sheet_Index) |>
    # Remove subtotal rows
    filter(!grepl("Total", Pollutant, ignore.case = TRUE)) |>
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
        pollutant == "PM10 (Particulate Matter < 10um)" ~ "PM10", # Simplify format
        pollutant == "PM2.5 (Particulate Matter < 2.5um)" ~ "PM2.5", # Simplify format
        pollutant == "Dioxins (PCDD/F)" ~ "Dioxins", # Simplify format
        pollutant == "Napthalene" ~ "Naphthalene", # Fix typo
        pollutant == "Decabromo diphenylether" ~ "Decabromodiphenyl ether", # Standard name
        pollutant == "Octabromo diphenylether" ~ "Octabromodiphenyl ether", # Standard name
        pollutant == "Pentabromo diphenylether" ~ "Pentabromodiphenyl ether", # Standard name
        pollutant == "Pentabromodiphenyl Ether" ~ "Pentabromodiphenyl ether", # Fix capitalization
        pollutant == "Acenapthene" ~ "Acenaphthene", # Fix typo
        pollutant == "Acenapthylene" ~ "Acenaphthylene", # Fix typo
        TRUE ~ pollutant # Keep everything else unchanged
    )) |>
    # Standardise units
    mutate(emission_unit = case_when(
        emission_unit == "Tonnes" ~ "tonne",
        emission_unit == "kg" ~ "kilogram",
        emission_unit == "t" ~ "tonne",
        TRUE ~ emission_unit # Keep everything else unchanged
    )) |>
    # Replace long hyphens with short hyphens in the source column
    mutate(source_name = str_replace_all(source_name, "–", "-")) |>
    # Add spaces to source_sheet which will be converted to pollutant group
    mutate(
        source_sheet = str_replace_all(source_sheet, "([a-z])([A-Z])", "\\1 \\2"),
        source_sheet = str_replace_all(source_sheet, "([a-zA-Z])&([a-zA-Z])", "\\1 & \\2") # Needed for POPs&PAHs
    ) |>
    # Rename columns for clarity
    rename(
        nfr = nfr_code,
        activity = activity_name,
        source = source_name,
        unit = emission_unit,
        pollutant_group = source_sheet
    )

# Add compliance totals ------------

# NO2 - have to calculate no2_compliance to exclude agriculture
no2 <- subset(historic_data, historic_data$pollutant == "Nitrogen Oxides as NO2")

no2_comp <- no2 %>%
    filter(!str_detect(nfr, "^3B|^3D")) %>%
    mutate(pollutant = "Nitrogen Oxides as NO2 NECR")

# VOC's - have to calculate voc_compliance to exclude agriculture
voc <- subset(historic_data, historic_data$pollutant == "Non Methane VOC")

voc_comp <- voc %>%
    filter(!str_detect(nfr, "^3B|^3D")) %>%
    mutate(pollutant = "Non Methane VOC NECR")

# NH3
nh3 <- subset(historic_data, historic_data$pollutant == "Ammonia")

# NH3 NECR Compliance Total (adjusted total)
nh3_comp <- nh3 %>%
    filter(!str_detect(source, "Digestates - TAN")) %>% # Should be using 3Da2c
    mutate(pollutant = "Ammonia NECR")

historic_data <- rbind(historic_data, no2_comp, voc_comp, nh3_comp)

# QA Checks ------------

# This section provides a number of QA checks to be performed on the data.

# Check data structure and column types
str(historic_data)

# Check for missing data
colSums(is.na(historic_data))
sum(complete.cases(historic_data)) == nrow(historic_data)
historic_data |> filter(is.na(emission))

# Check for duplicates
historic_data |> filter(duplicated(historic_data))

# Check text columns for errors such as typos
unique(historic_data$pollutant_group)
unique(historic_data$pollutant)
unique(historic_data$unit)

# Check for negative values in emissions
historic_data |> filter(emission < 0)

# Check years are consistent
range(historic_data$year)
historic_data |> filter(year < 1900 | year > 2030)

# Check units are consistent
table(historic_data$unit)

# Check missing pollutant-year pairs
all_combinations <- expand.grid(
    year = unique(historic_data$year),
    pollutant = unique(historic_data$pollutant)
)
missing_combinations <- anti_join(all_combinations, historic_data, by = c("year", "pollutant"))
view(missing_combinations |> filter(year >= 1990)) # Shows missing year-pollutant pairs from 1990 onwards
# Looks like we stopped reporting the following from 2015 onwards:
# Pentabromodiphenyl ether, Octabromodiphenyl ether, Short Chain Chlorinated Paraffins (C10-13)
# Also looks like we reported Pyrene from 1982 onwards.

# Check missing pollutant-reporting_year pairs
all_combinations <- expand.grid(
    reporting_year = unique(historic_data$reporting_year),
    pollutant = unique(historic_data$pollutant)
)
missing_combinations <- anti_join(all_combinations, historic_data, by = c("reporting_year", "pollutant"))
view(missing_combinations)
# Notably we didn't report black carbon and methane till 2017

# Check for extreme values in emissions
yearly_totals <- historic_data |>
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
# It looks like there are quite a few outliers

# Plot time series for each pollutant
# Check if ggforce is installed
if (!require("ggforce", character.only = TRUE)) {
    install.packages("ggforce")
}
yearly_totals <- historic_data |>
    group_by(year, pollutant, reporting_year) |>
    summarise(total_emission = sum(emission, na.rm = TRUE), .groups = "drop")
num_pollutants <- length(unique(yearly_totals$pollutant))
num_pages <- ceiling(num_pollutants / 4) # 4 per page
for (page in 1:num_pages) {
    p <- ggplot(yearly_totals, aes(x = year, y = total_emission, group = reporting_year, color = as.factor(reporting_year))) +
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
split(historic_data, historic_data$reporting_year) |>
    purrr::iwalk(~ write.csv(
        .x,
        file = paste0("Tidied_Data/Historic_Emissions/historic_emissions_reporting_", .y, ".csv"),
        row.names = FALSE
    ))

# Aggregate by NFR code (removing source and activity)
historic_data_by_NFR <- historic_data |>
    group_by(reporting_year, pollutant_group, pollutant, nfr, unit, year) |>
    summarise(emission = sum(emission, na.rm = TRUE), .groups = "drop")

# Save aggregated version by reporting year
split(historic_data_by_NFR, historic_data_by_NFR$reporting_year) |>
    purrr::iwalk(~ write.csv(
        .x,
        file = paste0("Tidied_Data/Historic_Emissions_by_NFR/historic_emissions_by_NFR_reporting_", .y, ".csv"),
        row.names = FALSE
    ))
