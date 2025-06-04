# CL follow up requests

# get career level by division for director deep dive
tmp <- calculate_voluntary_turnover_monthly(df = pa_at, grouping_var = c("division_function", "career_level_bucket"))

write.csv(tmp, "tmp.csv")

tmp_ifr_gender <- calculate_internal_fill_monthly(df = full_df, grouping_var = c("career_level_bucket", "gender")) |> 
    filter(month == "ytd")

tmp_ifr_re <- calculate_internal_fill_monthly(df = full_df, grouping_var = c("career_level_bucket", "race_ethnicity"))


# get latest headcount data for baseline comparison
may_hc <- monthly_turnover_data$May |> 
    left_join(df_pa_at |> select(employee_id, career_level_bucket, gender, race_ethnicity))

may_hc |> 
    group_by(career_level_bucket, gender) |>
    summarize(n = n()) |>
    mutate(percent = round(n / sum(n), 2))

race ethnicity includes US only
may_hc_re <- may_hc |> 
    filter(location_address_country == "United States of America") |> 
    mutate(race_ethnicity_label = str_extract(race_ethnicity, "^[^\\(]+"),
            minority_status = case_when(
                race_ethnicity_label == "White " ~ "Non-Minority",
                race_ethnicity_label %in% c("Not Specified", "I do not wish to answer.", NA) ~ "Unknown",
                TRUE ~ "Minority"
            ))
    
may_hc_re |> 
    group_by(race_ethnicity_label) |>
    summarize(n = n()) |>
    mutate(percent = round(n / sum(n), 2))

may_hc_re |> 
#    filter(minority_status != "Unknown") |> 
    group_by(career_level_bucket, minority_status) |>
    summarize(n = n()) |>
    mutate(percent = round(n / sum(n), 2))
