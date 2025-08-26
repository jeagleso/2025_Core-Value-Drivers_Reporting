# General

library(tidyverse)
library(workdayr)

# theme colors
r_darkblue <- "#002F6C"
r_green <- "#71B840"
r_gray <- "#6D6F71"
r_lightblue <- "#648EB6"
r_brightblue <- "#00B0F0"
r_orange <- "#E97132"

# defaults to start_date first day of the previous year, end date last day of the current year, and effective_date first day of current month
get_workday_at <- function(start_date = as.Date(floor_date(Sys.Date(), "year") - years(1)), 
                            end_date = as.Date(floor_date(Sys.Date(), "year") + years(1) - days(1)), 
                            effective_date = as.Date(floor_date(Sys.Date(), "month"))) {
  
  # use workdayr package to pull active and terminated file
  report_path <- workdayr::get_workday_report(
    report_name = '610171734/People_Analytics_-_AT', 
    username = Sys.getenv("WD_AT_UN"), 
    password = Sys.getenv("WD_AT_PW"), 
    params = list(report_start_date = start_date,
                  report_end_date = end_date,
                  Effective_as_of_Date = effective_date,
                  format ='csv'), 
    organization = 'regalrexnord',
    filepath = tempfile(),
    overwrite = TRUE,
    endpoint = 'https://wd2-impl-services1.workday.com/ccx/service/customreport2/'
  )

  # read in report for assignment
  read_csv(report_path)
}

apply_general_filters <- function(df) {
  df |> 
    filter(worker_type == "Employee") |> 
    filter(job_family_group != "Human Resources" | is.na(job_family_group)) |>
    filter(segment_function != "Industrial Systems") |>
    filter(division_function != "Corp Human Resources")
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


#' Calculate monthly, quarterly, and YTD voluntary turnover with stable types
#'
#' @param df A data frame with at least the columns:
#'   hire_date, termination_date, termination_category, and grouping columns.
#' @param grouping_var A character vector of column names to group by.
#'   If NULL, an "Enterprise" grouping level will be created.
#'
#' @return A tibble with columns:
#'   <grouping_var>, month (chr), voluntary_terminations (dbl),
#'   headcount (dbl), voluntary_turnover (dbl), ytd_annualized (dbl)
#'
#' Notes:
#' - `termination_category` is filtered to "Terminate Employee > Voluntary".
#' - `last_day_of_month` is expected to be available in the environment.
#'
#| label: calculate_turnover_at_file

# type is professional or manufacturing
calculate_voluntary_turnover_monthly <- function(df = pa_at, grouping_var = grouping_vars) {

  # allows join for NULL aka enterprise
  if (is.null(grouping_var)) {
    grouping_var <- "Enterprise"
    df <- df |> mutate(Enterprise = "Enterprise")
  }

  # for each date in last_day_of_month do this function
  monthly_results <- map_df(last_day_of_month, function(date) {
  
  # get monthly headcount inclusive
  headcount <- df |> 
    filter(hire_date <= date & (is.na(termination_date) | termination_date >= date)) |> 
    group_by(across(all_of(grouping_var))) |> 
    summarize(headcount = n(), .groups = 'drop')
  
  # get monthly voluntary terminations
  voluntary_terminations <- df |> 
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
      headcount = round(mean(headcount, na.rm = TRUE), 1),
      voluntary_turnover = round(voluntary_terminations / headcount, 4),
      .groups = 'drop'
    ) |> 
    mutate(month = paste0("Q", quarter_num))
  
 # calculate YTD turnover
  ytd_results <- monthly_results |> 
    group_by(across(all_of(grouping_var))) |> 
    summarize(
      voluntary_terminations = sum(voluntary_terminations),
      headcount = round(mean(headcount, na.rm = TRUE), 1),
      voluntary_turnover = round(voluntary_terminations / headcount, 4),
      .groups = 'drop'
    ) |> 
    mutate(
      month = "ytd",
      ytd_annualized = as.numeric(round(voluntary_turnover * (12 / length(last_day_of_month)), 4))
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

# get all sheets from an excel file
read_excel_allsheets <- function(filename) {
  sheets <- readxl::excel_sheets(filename)
  x <- lapply(sheets, function(X) {
    data <- readxl::read_xlsx(filename, sheet = X, guess_max = 5000)

    # If the sheet has employee_id and effective_date columns, handle duplicates
    if(all(c("employee_id", "effective_date") %in% names(data))) {
      # Define priority order for business_process_type
      priority_order <- c(
        "Promote Employee Inbound",
        "Hire",
        "Transfer Employee"
      )
      
      # Create a factor for business_process_type with custom order
      data <- data |>
        mutate(
          process_priority = case_when(
            business_process_type %in% priority_order ~ 
              factor(business_process_type, levels = priority_order),
            TRUE ~ factor(business_process_type)
          )
        ) |>
        # Group by employee_id and effective_date
        group_by(employee_id, effective_date) |>
        # Arrange by priority and take first row
        arrange(process_priority) |>
        slice(1) |>
        # Remove the temporary priority column
        select(-process_priority) |>
        ungroup()
    }
    
    return(data)
  })
  names(x) <- sheets
  x
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
    summarize(count = n_distinct(employee_id, effective_date), .groups = "drop_last") |> 
    ungroup() |> 
    complete(nesting(!!!syms(grouping_var)), month, 
              internal_external = c("Internal", "External"), 
                fill = list(count = 0)) |>  
    group_by(across(all_of(grouping_var)), month) |>  
    mutate(total = sum(count),
           fill_rate = round(count / total, 4)) |> 
    ungroup() |> 
    filter(internal_external == "Internal") |> 
    arrange(month) |> 
    select(-internal_external) |> 
    mutate(month = month(month, label = TRUE))

  # do  the same calculation by quarter
  quarterly <- df |> 
    mutate(quarter_num = quarter(time_in_job_profile_start_date, "quarter")) |> 
    group_by(across(all_of(grouping_var)), quarter_num, internal_external) |> 
    summarize(count = n_distinct(employee_id, effective_date), .groups = "drop_last") |> 
    ungroup() |> 
    complete(nesting(!!!syms(grouping_var)), quarter_num, 
              internal_external = c("Internal", "External"), 
              fill = list(count = 0)) |>  
    group_by(across(all_of(grouping_var)), quarter_num) |>  
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
           'Internal Fill Rate' = fill_rate) |> 
    select(-quarter_num)
  
  return(binded_df)
}


