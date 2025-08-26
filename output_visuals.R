library(ggtext)
library(hrbrthemes)

monthly_turnover_data <- map(grouping_vars, ~ calculate_voluntary_turnover_monthly(pa_at, .x)) |> 
  set_names(output_sheet_names)

combined_turnover_data <- 
  bind_rows(monthly_turnover_data, .id = "grouping_level") |> 
  mutate(year = "2025")

# previous year
# output file name for PY 
vt_output_name_py <- "outputs/Professional Voluntary Turnover Rate Monthly 2024-01-01 to 2024-12-31.xlsx"

# get sheet names
vt_sheet_names_py <- readxl::excel_sheets(vt_output_name_py)

# only get sheet names with metrics, not headcount or term sheets
vt_sheet_names_py <- vt_sheet_names_py[1:6]

# pull in each of the sheets
vt_data_py <- vt_sheet_names_py |> 
    map(~ readxl::read_excel(vt_output_name_py, sheet = .x) |> 
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

py_turnover <- all_turnover_data |> 
  filter(month == "ytd", year == "2024", grouping_level == "Enterprise") |> 
  pull(voluntary_turnover)

turnover_goal <- (py_turnover * 0.90) / 12

# visuals

# vt historical
viz_vt_historical <- all_turnover_data |> 
  filter(month != "ytd" & !str_detect(month, "Q")) |> 
  filter(grouping_level == "Enterprise") |> 
  mutate(month = factor(month, levels = month.abb, ordered = TRUE))

label_most_recent <- viz_vt_historical |> 
  filter(year == 2025) |> 
  slice_max(order_by = month, n = 1)

viz_vt_historical |> 
  ggplot(aes(x = month, y = voluntary_turnover, group = year, color = year)) + 
  geom_hline(yintercept = turnover_goal, color = r_gray, size = 1.5, alpha = 0.5) +
  geom_line(size = 1.5) +
  geom_point(size = 3) +
  geom_text(data = label_most_recent, 
    aes(label = scales::percent(voluntary_turnover, accuracy = 0.1)),
  hjust = -0.2, vjust = 1.2) +
  scale_y_continuous(labels = scales::percent_format(accuracy = 0.1), limits = c(0, 0.015)) +
  scale_color_manual(values = c("2024" = r_darkblue, "2025" = r_green)) +
  labs(
    title = "<span style='color:#00A651;'>Monthly Voluntary Turnover</span> vs. <span style='color:#003865;'>2024</span>",
    x = NULL,
    y = NULL
  ) +
  theme_ipsum() +
  theme(
    plot.title = element_markdown(),
    plot.margin = margin(2, 2, 2, 2),
    legend.position = "none",
    panel.grid.minor = element_blank(),
    panel.grid.major.x = element_blank()
  )

ggsave("outputs/visuals/enterprise_vt_monthly_historical.png", width = 6.73, height = 2.1, units = "in")
