library(tidyverse)

# Assuming these are already defined in your environment:
# vt_prev, vt_current, column_order, column_order_dynamic, remaining_months_tibble

attribute <- 'Segment'
previous_output <- vt_prev
current_output <- vt_current

# Step 1: Previous year actuals
prev <- previous_output[[attribute]] |> 
  filter(segment_function == "Industrial Powertrain Solutions (IPS)") |> 
  mutate(data_source = "PY Actual") |> 
  select(-headcount, -voluntary_terminations, -ytd_annualized) |> 
  pivot_wider(names_from = month, values_from = voluntary_turnover) |> 
  select(where(is.character), "ytd", all_of(column_order))

# Step 2: Goal (AOP)
goal <- previous_output[[attribute]] |> 
  select(where(is.character), voluntary_turnover) |> 
  mutate(data_source = "AOP") |> 
  mutate(Goal = voluntary_turnover * 0.90) |> 
  select(-voluntary_turnover) |> 
  pivot_wider(names_from = month, values_from = Goal) |> 
  select(where(is.character), "ytd", all_of(column_order))

# Step 3: Current year actuals
current <- current_output[[attribute]] |> 
  mutate(data_source = "CY Actual") |> 
  select(-headcount, -voluntary_terminations, -ytd_annualized) |> 
  pivot_wider(names_from = month, values_from = voluntary_turnover) |> 
  select(where(is.character), "ytd", all_of(column_order_dynamic))

# Step 4: YTD Annualized
ytd_annualized <- current_output[[attribute]] |> 
  filter(month == "ytd", !is.na(ytd_annualized)) |> 
  select(-headcount, -voluntary_terminations, -month, -voluntary_turnover)

# Step 5: Forecast
forecast <- ytd_annualized |> 
  mutate(forecast = ytd_annualized / 12) |> 
  bind_cols(remaining_months_tibble) |> 
  mutate(across(everything(), ~ ifelse(is.na(.), 
                                       case_when(str_detect(cur_column(), "Q") ~ forecast * 3, 
                                                 TRUE ~ forecast), .))) |> 
  select(-ytd_annualized, -forecast)

# Step 6: Commit/Forecast
commit_forecast <- current |> 
  left_join(forecast) |> 
  mutate(data_source = "Commit/Forecast")

# Step 7: Combine all
combined <- prev |> 
  bind_rows(goal) |> 
  bind_rows(commit_forecast) |> 
  left_join(ytd_annualized) |> 
  mutate(segment_function = case_when(
    segment_function == "Corporate" ~ "Corporate",
    TRUE ~ str_extract(segment_function, pattern = "\\(([^)]+)\\)")
  )) |> 
  mutate(segment_function = str_remove_all(segment_function, "[()]")) |> 
  mutate(cvd = case_when(
    identical(current_output, vt_current) ~ "Professional Voluntary Turnover",
    identical(current_output, mf_vt_current) ~ "Manufacturing Voluntary Turnover",
    TRUE ~ "Voluntary Turnover"
  )) |> 
  relocate(cvd, .before = data_source)

# Step 8: Final formatting
reformat <- combined |> 
  rename(FY = ytd,
         ytd_value = ytd_annualized) |> 
  relocate(ytd_value, .before = data_source) |> 
  relocate(FY, .after = last_col()) |> 
  mutate(FY = case_when(data_source == "Commit/Forecast" ~ ytd_value,
                        TRUE ~ FY))

# View the final result
reformat
