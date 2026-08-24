# Myexcel_repo
# Retail & Bike Sales Analysis – Excel Project

Excel project completed as part of the **Level 3 Data Technician Bootcamp** (Just IT Training Ltd). The project works with retail transaction data, student grades, and bike sales data to demonstrate core spreadsheet skills used in a data technician role: table structuring, formulas, filtering and sorting, PivotTables, the SWITCH function, and chart building.

## Project overview

The work is split across several linked exercises:

- **Retail sales dataset** – transaction-level data (Transaction ID, Date, Customer ID, Gender, Age, Product Category, Quantity, Price per Unit) converted into a formal Excel Table and analysed using formulas and filters.
- **Student grades dataset** – exam scores across English, Mathematics and Science, used to practise averages, conditional formatting and flagging students who need support.
- **County/product sales dataset** – regional sales volumes by product, summarised with a PivotTable and categorised using the SWITCH function.
- **Bike sales dataset** – multi-year revenue and profit data, summarised with PivotTables and visualised with PivotCharts.

## Skills demonstrated

### Formulas
- `SUM`, `AVERAGE` – totalling and averaging commission and sales figures
- `SUMIF`, `AVERAGEIF` – conditional totals and averages based on category/criteria
- `DATE`, `MONTH`, `YEAR` – working with and extracting date components from transaction dates
- `UNIQUE` – pulling distinct lists of values (e.g. products, counties) from raw data
- `VLOOKUP` – looking up and returning matching values across tables
- `SWITCH` – categorising sales volume into High/Medium/Low bands, e.g.:

  ```excel
  =SWITCH(TRUE, C2 > 600, "High", C2 >= 300, "Medium", "Low")
  ```

### Data structuring, filtering and sorting
- Converting raw ranges into formal Excel Tables (`Ctrl+T`) for structured referencing
- Filtering (e.g. sorting Age from largest to smallest)
- Conditional formatting (colour scales) to visually flag performance, e.g. highlighting students below the average threshold as needing support

### PivotTables
- Summarising sales by County and Product Category
- Cross-tabulating order quantity by Age Group and Country
- Building multi-year revenue vs profit summaries
- Cleaning source data (removing hidden spaces, confirming numerical data types) so PivotTables calculate correctly

### PivotCharts and visualisations
- Clustered/stacked column charts (sales by gender/age, revenue by country and product category)
- Line charts comparing annual profit vs revenue over time
- Pie charts showing revenue share by age group

## Example outputs

- PivotTable summarising **Sales Volume by County and Product**, with a SWITCH-based category column
- PivotChart comparing **Annual Profit vs Annual Revenue (2017–2021)**
- Stacked bar chart of **Product Revenue by Country** (Accessories, Bikes, Clothing)
- Pie chart of **Revenue Comparison by Age Group**
- Conditional-formatted **student grades table** flagging students needing support

## Tools used

- Microsoft Excel (Tables, PivotTables, PivotCharts, conditional formatting, formulas)

## About this project

This project forms part of the Data Technician pathway, building foundational Excel skills ahead of moving on to SQL, Python, Power BI and Tableau within the same bootcamp.
