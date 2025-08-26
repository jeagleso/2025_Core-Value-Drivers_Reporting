monthly_turnover_data <- map(grouping_vars, ~ calculate_voluntary_turnover_monthly(pa_at, .x)) |> 
  set_names(output_sheet_names)

combined_turnover_data <- 
  bind_rows(monthly_turnover_data, .id = "grouping_level") |> 
  mutate(year = "2025")

# previous year
# output file name for PY 
vt_output_name_py <- "outputs/Professional Voluntary Turnover Rate Monthly 2024-01-01 to 2024-12-31.xlsx"

# get sheet names
vt_sheet_names_py <- excel_sheets(vt_output_name_py)

# only get sheet names with metrics, not headcount or term sheets
vt_sheet_names_py <- vt_sheet_names_py[1:6]

# pull in each of the sheets
vt_data_py <- vt_sheet_names_py |> 
    map(~ read_excel(vt_output_name_py, sheet = .x) |> 
    mutate(sheet = .x))

# bind together
vt_df_py <- bind_rows(vt_data_py) |> 
    mutate(year = "2024") |> 
    rename(grouping_level = sheet)

all_turnover_data <- combined_turnover_data |> 
  bind_rows(vt_df_py) |> 
  select(-Enterprise) |> 
  mutate(seg_name = case_when(
    segment_function == "Automation and Motion Control (AMC)" ~ "AMC",
    segment_function == "Industrial Powertrain Solutions (IPS)" ~ "IPS",
    segment_function == "Power Efficiency Solutions (PES)" ~ "PES",
    segment_function == "Corporate" ~ "Corporate",
    TRUE ~ NA
)) |> 
  mutate(year = as.factor(year))

# visuals
viz_vt_historical <- all_turnover_data |> 
  filter(month != "ytd" & !str_detect(month, "Q")) |> 
  filter(grouping_level == "Enterprise") |> 
  mutate(month = factor(month, levels = month.abb, ordered = TRUE))

viz_vt_historical |> 
  ggplot(aes(x = month, y = voluntary_turnover, group = year, color = year)) + 
  geom_line(size = 1.2) +
  geom_point(size = 1.2) +
  scale_y_continuous(labels = scales::percent_format(accuracy = 0.1), limits = c(0, 0.015)) +
  scale_color_manual(values = c("2024" = r_green, "2025" = r_darkblue)) +
  labs(
    title = "Monthly Voluntary Turnover vs. 2024",
    x = NULL,
    y = NULL
  ) +
  theme_minimal(base_size = 14) +
  theme(
    panel.grid.minor = element_blank()
  )
