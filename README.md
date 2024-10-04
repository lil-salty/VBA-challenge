# VBA-challenge
ASU Data Analytics Bootcamp module 2 VBA challenge 
Lily Saltonstall Oct2024

Made a VBA script to run on stock data set and complete the following requirements:

Requirements
Retrieval of Data (20 points)
The script loops through one quarter of stock data and reads/ stores all of the following values from each row:
ticker symbol (5 points)
volume of stock (5 points)
open price (5 points)
close price (5 points)

Column Creation (10 points)
On the same worksheet as the raw data, or on a new worksheet all columns were correctly created for:
ticker symbol (2.5 points)
total stock volume (2.5 points)
quarterly change ($) (2.5 points)
percent change (2.5 points)

Conditional Formatting (20 points)
Conditional formatting is applied correctly and appropriately to the quarterly change column (10 points)
Conditional formatting is applied correctly and appropriately to the percent change column (10 points)
Calculated Values (15 points)
All three of the following values are calculated correctly and displayed in the output:
Greatest % Increase (5 points)
Greatest % Decrease (5 points)
Greatest Total Volume (5 points)

Looping Across Worksheet (20 points)
The VBA script can run on all sheets successfully.
GitHub/GitLab Submission (15 points)
All three of the following are uploaded to GitHub/GitLab:
Screenshots of the results (5 points)
Separate VBA script files (5 points)
README file (5 points)


Sources for script include:
ASU BootCamp Module activities
  Locations: declaring variables, initializing variables, storing values, making if then loops, making for next loops, searching cell values, comparing cell values, percent change calculations, colorindex on cells/values
ASU BootCamp Homework Study Group (Participants: Andrew Lane, Angelina Wright, Amy Johnson, Andrew Jaynes, Aubriana Osborn, Aditi Nankar, Victoria Mendez)
  Locations: presub section for all worksheet loop, with/end with loop, using lastrow, "sub stock_analysis(ws As Worksheet):", formatting in VSC to better see loop start and ends
GitHub (user: laurajordan845)
  Locations: End(x1Up), numberformat on ranges, using "& 2 + n" (w/ n as integer, n = 0) inside of ranges, loops inside of loops ordering
StackOverflow questions
  Locations: using "& 2 + n" (w/ n as integer, n = 0) inside of ranges, troubleshooting
