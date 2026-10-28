# Packages ----------------------------------------------------------------
# Basic data processing
library(dplyr)
library(stringr)

# Data reading options in R
library(readr)
library(openxlsx)
library(readxl)

# SharePoint library built for R
library(Microsoft365R)

# Basic Inputs need for Site, Library and documents -----------------------
the_site = "AgencyInformation"
the_library <- "Data Testing"

# Test documents
excel_file <- "transit_data_test.xlsx"
csv_file <- "transit_data_test.csv"
updated_excel_file <- "updated_transit_data_test.xlsx"

# Connect to SharePoint ---------------------------------------------------
# Piset has authorized the token need to authenticate PSRC staff
# so this should work without needing any parameters

print(str_glue("Connecting to the SharePoint library {the_library} in the {the_site} SharePoint site"))

sp_site <- get_sharepoint_site(
  site_url = paste0("https://psrcwa.sharepoint.com/sites/EDRMS-", the_site),
  tenant = "psrcwa.onmicrosoft.com"
)

sp_library <- sp_site$get_drive(the_library)

# Read a file on SharePoint directly into R ---------------------------
# load_dataframe opens an excel or csv file
# you can use the same arguments in load_dataframe that you can use
# in readxl or read_csv

print(str_glue("Reading {csv_file} into memory using load_dataframe")) 
transit_csv <- sp_library$load_dataframe(csv_file,
                                         show_col_types = FALSE)

print(str_glue("Reading {excel_file} into memory using load_dataframe")) 
transit_xlsx <- sp_library$load_dataframe(excel_file,
                                          sheet = "Data",
                                          skip = 2)

# Download the file to a temporary local path --------------------------
# You can also download a file and then process it from SharePoint.
# This would be a way to work with a file type that doesn't cleanly
# work with load_dataframe. I used an Excel file but a non-tidyverse library
# to highlight this.

print(str_glue("Reading {excel_file} into memory using a local file")) 
local_file <- tempfile(fileext = ".xlsx")

sp_library$download_file(
  src = excel_file,
  dest = local_file,
  overwrite = TRUE
)

transit_open_xlsx_df <- read.xlsx(
  local_file,
  sheet = 2,
  startRow = 3
)

# Write the modified data back to the local Excel file -----------------
# Say we want to add a few new columns and then upload this file to SharePoint
# First lets add the columns to any of our three dataframes, they are all the
# same, just in-memory versions of the data.

updated_transit_data <- transit_csv |>
  mutate(`Boardings per Hour` = round(Boardings / Hours, 2)) |>
  mutate(`Annual Change in Boardings` = Boardings - lag(Boardings)) |>
  mutate(`Annual Change in Hours` = Hours - lag(Hours))

# To write to SharePoint, we first need this in-memory file written to disc
# Using a temp_file is a great way to avoid cluttering you computer
# local_file created this temp file that will be removed with the unlink
# command at the end.

write.xlsx(
  updated_transit_data,
  file = local_file,
  overwrite = TRUE
)

# Upload the modified file to SharePoint -------------------------------
# Let's give it a new name to see if we can add a new file

print(str_glue("Uploading {updated_excel_file} to SharePoint")) 
sp_library$upload_file(
  src = local_file,
  dest = updated_excel_file
)

# Clean up the temporary local file ------------------------------------
unlink(local_file)

# Updating the Metadata ---------------------------------------------------
# This seems a little cumbersome but is programmatic.
# To change to metadata, you first need to get the ID of the object
# on SharePoint for the file you are updating.
# I used the value in FileLeafRef to find the ID here
sp_test <- sp_site$get_list(the_library)
items <- sp_test$list_items(as_data_frame = TRUE)
matched_row <- which(items$FileLeafRef == updated_excel_file)

# Once we have the id that matches our filename, we can get that item
# to update the metadata
sp_item <- sp_test$get_item(items$id[matched_row])

# Now we can update the SharePoint objects metadata in R. Change the year
# or topic to see if it worked

print(str_glue("Uploading {updated_excel_file} metadata on SharePoint")) 
sp_item$update(
  fields = list(
    Year = "Fall 2026",
    Topic = "R code testing of metadata"
  )
)
