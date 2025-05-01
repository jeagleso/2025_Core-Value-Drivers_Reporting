tmp <- vt_current$Segment |>
        mutate(data_source = "CY Actual") |> 
        select(-headcount, -voluntary_terminations, -ytd_annualized) |> 
        pivot_wider(names_from = month, values_from = voluntary_turnover) |> 
        select(where(is.character), "ytd", all_of(column_order_dynamic))

tmp_annualized <- vt_current$Segment |> 
        select(-headcount, -voluntary_terminations, -month, -voluntary_turnover) |> 
        filter(!is.na(ytd_annualized))

remaining_months_tibble <- tibble(remaining_months,
                                  value = NA) |> 
  pivot_wider(names_from = remaining_months, values_from = value)

tmp_forecast <- tmp_annualized |> 
  mutate(forecast = ytd_annualized / 12) |> 
  bind_cols(remaining_months_tibble) |> 
  mutate(across(everything(), ~ ifelse(is.na(.), case_when(str_detect(cur_column(), "Q") ~ forecast * 3, TRUE ~ forecast), .))) |> 
  select(-ytd_annualized, -forecast)

tmp_fin <- tmp |> left_join(tmp_forecast)


tmp <- ifr_current$Segment  |>
        mutate(data_source = "CY Actual") |> 
        select(-`Roles Filled Internally`, -`Total Filled Roles`) |> 
        pivot_wider(names_from = month, values_from = `Internal Fill Rate`) |> 
        select(where(is.character), "ytd", all_of(column_order_dynamic))

tmp_forecast <- tmp |> 
  select(1:ytd) |> 
  bind_cols(remaining_months_tibble) |> 
  mutate(across(everything(), ~ ifelse(is.na(.), ytd, .)))
