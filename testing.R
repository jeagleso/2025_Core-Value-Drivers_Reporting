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

tmp_vt_current <- vt_current <- read_excel_allsheets(filename = "tmp.xlsx")

tmp_div_loc <- readxl::read_xlsx("tmp.xlsx", sheet = "Division by Location", col_types = c("guess", "guess", "guess", "numeric", "numeric", "guess", "numeric", "numeric"))



tmp_vt_current$`Division by Location`<- tmp_vt_current$`Division by Location` |> 
        mutate(ytd_annualized = as.numeric(ytd_annualized))

tmp_mfr_fy <- vt_current$`Division by Location` |> 
        filter(month == "ytd", !is.na(ytd_annualized)) |> 
        select(-headcount, -voluntary_terminations, -month, -voluntary_turnover)


vt_current$Segment |> 
        filter(month == "ytd") |> 
        select(-headcount, -voluntary_terminations, -month, -ytd_annualized) |> 
        rename(ytd_performance = voluntary_turnover)

vt_current$Segment |> 
        filter(month == "ytd", !is.na(ytd_annualized)) |> 
        select(-headcount, -voluntary_terminations, -month, -voluntary_turnover)

library(workdayr)

# Get report raw data
report_path <- get_workday_report(
  report_name = '610171734/People_Analytics_-_AT', 
  username = Sys.getenv("WD_AT_UN"), 
  password = Sys.getenv("WD_AT_PW"), 
  params = list(report_start_date ='2025-01-01',
                report_end_date = '2025-04-01',
                Effective_as_of_Date = '2025-04-01',
                format ='csv'), 
  organization = 'regalrexnord',
  filepath = tempfile(),
  overwrite = TRUE,
  endpoint = 'https://wd2-impl-services1.workday.com/ccx/service/customreport2/'
)

# read it in as a tibble
tmp <- readr::read_csv(report_path)
