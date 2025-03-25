# General

library(tidyverse)

apply_general_filters <- function(df) {
  df |> 
    filter(worker_type == "Employee") |> 
    filter(job_family_group != "Human Resources") |> 
    filter(segment_function != "Industrial Systems")
}

add_buckets <- function(df) {
  df |> 
    mutate(career_level_bucket = factor(case_when(
    career_level %in% c("E2", "E3", "E4", "E5") ~ "VP",
    career_level %in% c("M5", "E1") ~ "Director",
    career_level %in% c("M2", "M3", "M4") ~ "Manager",
    career_level == "M1" ~ "Supervisor",
    career_level %in% c("P1", "P2", "P3", "P4", "P5", "P6") ~ "Professional",
    career_level %in% c("AT1", "AT2", "AT3", "AT4") ~ "Administrative_Technical",
    is.na(career_level) ~ "DL_IDL",
    TRUE ~ "ERROR_unmapped_level"),
    levels = c("VP", "Director", "Manager", "Supervisor", "Professional", "Administrative_Technical", "DL_IDL", "ERROR_unmapped_level"),
    ordered = TRUE))
}


## CVD Metrics

# set the grouping variables of interest
# if changing grouping variables, also add output_sheet_name
grouping_vars <- list(NULL, 
                      "segment_function", 
                      c("segment_function", "division_function"), 
                      c("segment_function", "division_function", "business_unit_sub_function"),
                      c("segment_function", "division_function", "location"),
                      c("segment_function", "career_level_bucket"),
                      "career_level_bucket")

# set name for output sheet name for each grouping variable
output_sheet_names <- c("Enterprise",
                        "Segment",
                        "Segment by Division",
                        "Division by Subfunction",
                        "Division by Location",
                        "Segment by Level",
                        "Enterprise by Level")

### Voluntary Turnover


#| label: calculate_turnover_at_file

# type is professional or manufacturing
calculate_voluntary_turnover_monthly <- function(df = pa_at, grouping_var = grouping_vars, type = "prof") {
  
  if (type == "prof") {
    pa_at <- pa_at |> 
      filter(career_level_bucket != "DL_IDL")
  } else if (type == "manuf") {
    pa_at <- pa_at |> 
      filter(career_level_bucket == "DL_IDL")
  } else {
    stop("Invalid type. Enter 'prof' or 'manuf'.")
  }

  # allows join for NULL aka enterprise
  if (is.null(grouping_var)) {
    grouping_var <- "Enterprise"
    pa_at <- pa_at |> mutate(Enterprise = "Enterprise")
  }

  # for each date in last_day_of_month do this function
  monthly_results <- map_df(last_day_of_month, function(date) {
  
  # get monthly headcount inclusive
  headcount <- pa_at |> 
    filter(hire_date <= date & (is.na(termination_date) | termination_date >= date)) |> 
    group_by(across(all_of(grouping_var))) |> 
    summarize(headcount = n(), .groups = 'drop')
  
  # get monthly voluntary terminations
  voluntary_terminations <- pa_at |> 
    filter(termination_category == "Terminate Employee > Voluntary" &
          termination_date >= floor_date(date, "month") &
          termination_date <= date) |> 
    group_by(across(all_of(grouping_var))) |> 
    summarize(voluntary_terminations = n(), .groups = 'drop')

  # combine headcount and VTs, calculate voluntary turnover
  result <- left_join(headcount, voluntary_terminations, by = grouping_var) |> 
    mutate(
      date = date,
      month = month(date, label = TRUE),
      voluntary_terminations = coalesce(voluntary_terminations, 0), # returns 0 if none
      voluntary_turnover = round(voluntary_terminations / headcount, 4)
      )
  
  return(result)
})

# calculate quarterly turnover
  quarterly_results <- monthly_results |> 
    mutate(quarter_num = quarter(date, "quarter")) |> 
    group_by(across(all_of(grouping_var)), quarter_num) |> 
    summarize(
      voluntary_terminations = sum(voluntary_terminations),
      headcount = mean(headcount, na.rm = TRUE),
      voluntary_turnover = round(voluntary_terminations / headcount, 4),
      .groups = 'drop'
    ) |> 
    mutate(month = paste0("Q", quarter_num))
  
 # calculate YTD turnover
  ytd_results <- monthly_results |> 
    group_by(across(all_of(grouping_var))) |> 
    summarize(
      voluntary_terminations = sum(voluntary_terminations),
      headcount = mean(headcount, na.rm = TRUE),
      voluntary_turnover = round(voluntary_terminations / headcount, 4),
      .groups = 'drop'
    ) |> 
    mutate(
      # date = "YTD",
      month = "YTD",
      ytd_annualized = round(voluntary_turnover * (12 / length(last_day_of_month)), 4)
    )
  
  # combine monthly and YTD results
  final_results <- bind_rows(
    monthly_results |> mutate(date = as.character(date)),
    quarterly_results,
    ytd_results
  ) |> 
    select(-date, -quarter_num)
  
  return(final_results)
}


# if you need to read and combine monthly files (e.g., multiple term or headcount files)
read_and_combine <- function(directory_path, row_skip) {
  files <- list.files(directory_path, full.names = TRUE, pattern = "*.xlsx")
  data_list <- lapply(files, function(file) {
    data <- readxl::read_xlsx(file, sheet = "Raw Data", skip = row_skip) |> 
      janitor::clean_names() |> 
      filter(worker_type == "Employee") 
    data <- data |> 
      mutate(source_file = basename(file)) |> 
      mutate(effective_date = str_extract(source_file, 
                                          "(?<=2024 ).*(?=\\.xlsx)") |> ymd())
    return(data)
  })
  combined_data <- bind_rows(data_list)
  return(combined_data)
}

# calculate internal fill
# create function for internal fill calculation using df and running for each grouping_var by each month
calculate_internal_fill_monthly <- function(df = full_df, grouping_var) {

  # calculate internal fill by month
  df_month <- df |> 
    mutate(month = as.Date(floor_date(time_in_job_profile_start_date, "month"))) |> 
    group_by(across(all_of(grouping_var)), month, internal_external) |> 
    summarize(count = n(), .groups = "drop_last") |> 
    group_by(across(all_of(grouping_var)), month) |> 
    complete(internal_external = c("Internal", "External"), 
             fill = list(count = 0)) |>  
    mutate(total = sum(count),
           fill_rate = round(count / total, 4)) |> 
    ungroup() |> 
    filter(internal_external == "Internal") |> 
    arrange(month) |> 
    select(-internal_external) 

  # do  the same calculation by quarter
  quarterly <- df |> 
    mutate(quarter_num = quarter(time_in_job_profile_start_date, "quarter")) |> 
    group_by(across(all_of(grouping_var)), quarter_num, internal_external) |> 
    summarize(count = n(), .groups = "drop_last") |> 
    group_by(across(all_of(grouping_var)), quarter_num) |> 
    complete(internal_external = c("Internal", "External"), 
             fill = list(count = 0)) |>  
    mutate(total = sum(count),
           fill_rate = round(count / total, 4)) |> 
    ungroup() |> 
    filter(internal_external == "Internal") |> 
    arrange(quarter_num) |> 
    select(-internal_external) |> 
    mutate(month = paste0("Q", quarter_num))

  # group together the months (aka group by grouping_var) to get YTD
  ytd <- df_month |> 
    group_by(across(all_of(grouping_var))) |> 
    summarize(count = sum(count),
              total = sum(total),
              fill_rate = count / total, .groups = "drop_last") |> 
    mutate(month = "ytd")
  
  # bind quarterly and ytd to month data
  binded_df <- df_month |> 
    mutate(month = as.character(month)) |> 
    bind_rows(quarterly) |> 
    bind_rows(ytd) |> 
    rename('Roles Filled Internally' = count,
           'Total Filled Roles' = total,
           'Internal Fill Rate' = fill_rate)
  
  return(binded_df)
}


