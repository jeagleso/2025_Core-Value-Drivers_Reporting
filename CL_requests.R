# CL follow up requests

# get career level by division for director deep dive
tmp <- calculate_voluntary_turnover_monthly(df = pa_at, grouping_var = c("division_function", "career_level_bucket"))

write.csv(tmp, "tmp.csv")

tmp_ifr_gender <- calculate_internal_fill_monthly(df = full_df, grouping_var = c("career_level_bucket", "gender"))

tmp_ifr_re <- calculate_internal_fill_monthly(df = full_df, grouping_var = c("career_level_bucket", "race_ethnicity"))
