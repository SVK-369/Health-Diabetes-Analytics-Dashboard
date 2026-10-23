# Health-Diabetes-Analytics-Dashboard

​1. Data Cleaning
​Question 1: Check for missing values marked with '?'
​Explanation: In Excel, search for ? using Find & Replace (Ctrl + F) or filter columns to see entries containing ?.  
​Excel Formula (to count ? in a column):
=COUNTIF(A:A, "?")
​Question 2: Fill missing values of 'month' with 'Sep' and 'year' with its average rounded to the nearest integer
​Explanation:
​Replace any ? in the month column with 'Sep'.  
​Calculate the average of valid years: =ROUND(AVERAGE(Year_Column), 0). Then replace ? in the year column with that calculated integer.  
​Question 3: Fill missing values for 'smoker', 'Hospital tier', and 'City tier' using mode
​Explanation: Find the most frequently occurring value (mode) for each column and replace ? or blank cells with it.  
​Method: Filter each column to identify the top/mode category, or use =MODE.SNGL(...) for numerical tiers.  
​Question 4: Fill missing 'State ID' values
​Explanation: Fill missing or ? entries in State ID with "Unknown".  
​Excel Formula / Method: Use Find & Replace (? \rightarrow Unknown) or conditional logic:
=IF(OR(A2="?", ISBLANK(A2)), "Unknown", A2)  
​2. Data Transformation
​Question 1: Split 'names' column into 'Title', 'First Name', and 'Last Name'
​Excel Feature: Use Text to Columns under the Data tab (Delimiter = Space).  
​Excel Formulas:
​Title: =LEFT(A2, FIND(" ", A2) - 1)  
​First Name: =MID(A2, FIND(" ", A2) + 1, FIND(" ", A2, FIND(" ", A2) + 1) - FIND(" ", A2) - 1)  
​Last Name: =RIGHT(A2, LEN(A2) - FIND(" ", A2, FIND(" ", A2) + 1))  
​Question 2: Convert 'NumberOfMajorSurgeries' to numerical data
​Explanation: If entries contain text like "No surgery" or "3 surgeries", replace them with clean numbers (0, 3, etc.) using Find & Replace or =VALUE(SUBSTITUTE(...)).  
​Question 3: Check for inconsistencies in 'Heart Issues' and 'smoker' columns
​Explanation: Check for typos, capitalization differences (e.g., yes, Yes, YES), or trailing spaces. Standardize them to consistent values like Yes / No using =PROPER(TRIM(Cell)) or Find & Replace.  
​Question 4: Create 'Weight Status' based on BMI
BMI RangeWeight Status
Below 18.5Underweight
18.5 – 24.9Normal Weight
25.0 – 29.9Overweight
30.0 and AboveObesity
​Excel Formula:
=IFS(BMI_cell < 18.5, "Underweight", BMI_cell <= 24.9, "Normal Weight", BMI_cell <= 29.9, "Overweight", BMI_cell >= 30, "Obesity")  
​Question 5: Create 'Diabetes Status' based on HbA1C

Question 5: Create 'Diabetes Status' based on HbA1CHbA1C RangeDiabetes Status
Below 5.7Normal
5.7 – 6.4Prediabetes
6.5 and AboveDiabetes
Excel Formula:
=IFS(HbA1C_cell < 5.7, "Normal", HbA1C_cell <= 6.4, "Prediabetes", HbA1C_cell >= 6.5, "Diabetes")

Question 6: Merge 'year', 'month', and 'date' into 'Date of Birth' (Format: DD-MMM-YYYY)
​Excel Formula:
=TEXT(DATEVALUE(date_cell & "-" & month_cell & "-" & year_cell), "dd-mmm-yyyy")  
​Question 7: Calculate 'Age' based on Date of Birth and collection date (8th June 2023)
​Excel Formula:
=DATEDIF(DOB_cell, DATE(2023, 6, 8), "Y")  
​Question 8: Format 'charges' column as currency ($)
​Explanation: Select the charges column \rightarrow Press Ctrl + 1 \rightarrow Select Currency \rightarrow Set symbol to $.  
​3. Data Exploration, Analysis & Visualization
​Consolidation (Healthcare Sheet)
​Create a new sheet named Healthcare.  
​Merge all 3 tables using VLOOKUP or XLOOKUP with Customer ID as the primary key:
=XLOOKUP(Customer_ID, Table2[Customer ID], Table2[Target_Column])  
​Keep only the specified columns: Customer ID, First Name, BMI, HbA1C, Heart Issues, Any Transplants, Cancer history, NumberOfMajorSurgeries, smoker, Weight Status, Diabetes Status, Date of Birth, charges, Hospital tier, City tier, State ID, Age.  
​Chart Specifications
​1. Pie/Donut Charts
​Cancer history distribution among smokers vs non-smokers:
​Pivot Table: Rows = smoker, Columns = Cancer history, Values = Count of Customer ID.  
​Chart: Pie/Donut Chart showing percentage share.  
​Major surgeries & average HbA1C by transplant history:
​Pivot Table: Rows = Any Transplants, Values = Sum of NumberOfMajorSurgeries, Average of HbA1C.  
​Chart: Donut Chart or Combo Chart.  
​2. Column/Bar Charts
​Healthcare charges by Weight Status and Diabetes Status:
​Pivot Table: Rows = Weight Status, Columns = Diabetes Status, Values = Average of charges.  
​Chart: Clustered Column Chart.  
​Average charges for each Hospital tier within different states:
​Pivot Table: Rows = State ID, Columns = Hospital tier, Values = Average of charges.  
​Chart: Clustered Column/Bar Chart.  
​3. Line/Scatter Plots
​Correlation between Age vs BMI and HbA1C:
​Chart: Scatter Plot with Age on X-axis and BMI / HbA1C on Y-axis (add trendlines to check correlation).  
​Relationship between Age and Healthcare charges:
​Chart: Scatter Plot with Age on X-axis and charges on Y-axis (add trendline).


​4. Dashboard Creation
​Layout: Place all generated charts neatly onto a clean dashboard sheet.  
​Add Slicers:
​Select any Pivot Chart \rightarrow Go to PivotChart Analyze \rightarrow Insert Slicer.  
​Select Weight Status and Diabetes Status.  
​Connect Slicers:
​Right-click each Slicer \rightarrow Report Connections \rightarrow Check all Pivot Tables to enable interactive cross-filtering across the full dashboard.

