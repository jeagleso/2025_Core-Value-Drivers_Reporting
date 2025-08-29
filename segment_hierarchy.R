# get segment hierarchy

library(tidyverse)

# read in active and terminated
active_and_terminated <- readxl::read_xlsx("inputs/People Analytics - AT.xlsx", skip = 4) |> 
  janitor::clean_names()

active <- active_and_terminated |> 
  filter(currently_active == "Yes")

segment_hierarchy <- active |> 
  select(segment_assignment, business_unit_sub_function, division_function, segment_function) |> 
  distinct()

segment_hierarchy |> 
  group_by(segment_assignment) |> 
  summarize(n_segs = n_distinct(segment_function)) |> 
  arrange(desc(n_segs))
