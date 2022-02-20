# Tracking and Trending
A VBA application designed to optimize input and analysis of facilities or otherwise categorical data. Optimal for maintenance for daily improvement (MDI) practices.

## Purpose
This file is designed to help input, maintain, and interpret MDI data. Data validation macros and user-forms help maintain data integrity and consistency. Data analysis macros can create calendars, reports, tables, and graphs. 

## Set up

### Add-Ins
This file requires several tools to be installed in the VBA project. To turn on the tools, go to file >> options >> customize ribbon >> and select the developer tab. After that, open the VBA window, go to tools >> references and make sure the following are selected:
- Visual Basic for Applications
- Microsoft Excel 16.0 Object Library
- OLE Automation
- Microsoft Office 16.0 Object Library
- Microsoft Forms 2.0 Object Library
- Microsoft Scripting Rutime
- Microsoft Powerpoint 16.0 Object Library

### Naming
As with all VBA applications, the consistent naming of modules, subprocedures, sheets, excel list objects, and excel table headers are important for function. If you change the naming or column header of one table, it can be difficult to know where else that name needs to be changed for the macros to function properly. It is advised that you keep as much close to the original as possible to maintain 100% functionality. That being said, there are procedures in place to allow renaming of a few columns within the main data table. More on that below.

## Contents

### Control Center

This is the sheet that houses the buttons that either activate or prompt userforms to enter data, create reports, or alter formatting. The macros consist of:
- Create Calendar
  - A report that creates an excel calendar worksheet of one year that identifies days where an incident or issue was recorded for a specified category
- Create KPI Charts
  - A report that creates a sheet and subsequent graph, which depicts what KPIs were hit over a interval of the users choice.
- Create Tag and Descriptors
  - A report that compiles all the tags, categories, KPIs, and other identifiers that were recorded in the MFDI data table over an interval of the users choice.
- Create Pivot Chart
  - A report that creates a pivot chart based on the users choice of a column of the main data table. Features include filtering by category, by total occurances of the event, and displaying the graph as a bar chart or running total.
- Generate Report
  - A report that transfers data from a user specified interval from the main data table to a powerpoint presentation. A static ppt template file must be present in an adjacent folder for success. Details on implementation below.
- Add New Entries
  - Generates a user-input form to streamline data entry to the main data table. Features include data validation and error handling.
- Update Countermeasures
  - A macro that reverts the main data table to the original format

### Countermeasures

The most important sheet is the "Countermeasures" sheet. On this sheet is a table named "Tbl_Counter." This table will hold all of the facilities/MDI data that needs to be stored, tracked, and trended. Each column maintains different information, and some columns allow multiple pieces of information to be stored. Some columns also have dynamic data validaton. The columns, a short description, their data type, their format, their validation properties, their naming convention, and an example are located in the table below.

| Column Title    | Meaning    |     Example/Format | Can Column be Renamed? |     Multiple entries/cell?    |     Data Validation?    |
|---|---|---|---|---|---|
| Issue ID | Unique ID number. They should not repeat, even for extensions.    | “221,   “290” | No |     No    |     No    |
| Issue Tier 1 Tag | A categorization word or phrase describing the general nature of the “Issue” cell.  (Tag Column) | “Alarm,”  “Substance Issue” | Yes |     No    |     Yes    |
| Issue Tier 2 Tag | A categorization word or phrase describing a more specific nature of the “Issue” cell.  (Tag Column) | “CIP,”  “Dry Link,”  “Material Exposure” | Yes |     No    |     Yes    |
| Cause Category | A categorization word or phrase describing the general nature of the “Cause” cell. (Tag Column) | “Equipment Failure,”  “Insufficient Instruction”    | Yes |     No    |     Yes    |
| Cause Detail | A categorization word or phrase describing a more specific nature of the “Cause” cell. (Tag Column) | “Decision Error,”  “Lacking Instruction”    | Yes |     No    |     Yes    |
| Issue Date | The date of the issue recorded.    | Any date data type (“21-Sep-21”)    | No |     No    |     No    |
| Category | The category, domain, or department that oversees or is impacted by the issue.    | “Quality,”  “Safety,”  “Delivery”    | No |     No    |     No    |
| KPI | A specific descriptor of the problem that is unique to the associated category.    | “Safety Incident,”  “Osha Recordable,”  “Open Deviation”    | No |     No    |     Yes    |
| Entry Descriptor | A unique descriptor of the entry, often found in other documentation/data systems (Descriptor Column) | “QE-XXXX,”   “Corp. Safety 45”    | Yes |     No    |     Yes    |
| Primary Equipment | The primary equipment(s) involved in the issue. (Multiple pieces of equipment should be delimited with a (“;_”).  (Descriptor Column) | “X-9274,”   “1963,”  “X-10938,”  “X-10093; X-10”    | Yes |     Yes    |     Yes    |
| Manufacturing Stage | The stage(s) involved in issue. (Multiple pieces of equipment should be delimited with a (“;_ ”).  (Descriptor Column) | “Synthesis,”   “Trimming,"  “Stamping,”  “Quality Check; Shipping”    | Yes |     Yes    |     Yes    |
| Batch | The batch(es) involved in the issue. (Multiple pieces of equipment should be delimited with a (“;_”) (Descriptor Column) | “10001,”   “10002,”  “10003; 10004” | Yes |     Yes    |     Yes    |
| Quality Classification | The quality designation of the issue. (Descriptor Column) | “Minor,”   “Major”    | Yes |     Yes    |     Yes    |
| Safety Tier | The safety designation of the issue. (Descriptor Column) | “Safety Tier 1,"   “Safety Tier 2”    | Yes |     No    |     Yes    |
| Issue | A description of the issue. Filled out in complete sentences.    | “A pool of a liquid was found on the floor in room number XYZ.”    | No |     No    |     No    |
| Cause | A description of the cause of the issue. Filled out in complete sentences.    | “A pipe in the room on equipment X-0000 was misaligned.”    | No |     No    |     No    |
| Countermeasure | After the issue is resolved, a description of the countermeasure/resolution that went into place for the issue.    | "Scheduled preventative maintenance for X-0000 under protocol 10.”    | No | No | No |
| Owner | The owner of the issue/event.    | “First Name Last Name”    | No | No | No |
| Date Due | The date the resolution of the issue/event is due.    | Any date data type (“21-Sep-21”)    | No | No | No |
| Date Closed | The date the issue/event was closed, and the resolution/report submitted.    | Any date data type (“21-Sep-21”)    | No | No | No |
| Status | If not yet closed, “Open.” All else “Closed.” | “Open,” or “Closed” | No | No |     No    |
| Issue Year | Auto populated from issue date. Issue year.    | “2020,"   “2021”    | No | No |     No    |
| Issue Month | Auto populated from issue date. Issue month.    | “1,”  “3,”    | No | No    |     No    |
| Month Name | Auto populated from issue date. Issue month name.    | “January”    | No | No    |     No    |
| Day of Month | Auto populated from issue date. Issue day of month.    | “12”    | No | No    |     No    |
| Yr-Month | Auto populated from issue date. Concatenation of issue year and month.    | “2020-7”    | No | No    |     No    |
| Days until Due | Days from issue date to assigned date due.    | “16”    | No | No    |     No    |
| Day Completed | Days from issue date to date closed.    | “16” | No | No    |     No    |
| On Time? | If days completed < days until due -> “Yes” | “Yes,” “No” | No | No | No |
| Early and Overdue Differential    | (-/+). (-) means entry closed before due date. (+) means entry closed after due date.    | “-4,” “2” | No | No    | No    |

### Data Validation
Some columns (as explained above) have data validation. But the data validation in this program is unique as it is dynamic. It is not designed to ensure that users are inputting one possible answer of a list of valid inputs. The list can be changed at will. 

In order to activate the data validation for a cell in a column, simply double click the cell. After double clicking, the cell will be anchored with a data validated drop down list consisting of all the other entries within that column. Having the data validation list active will not allow the user to enter something not within that list. To de-activate the dat validation, simply right click on the cell. Right clicking will not get rid of the contents of the cell, allowing the user immense flexibility in organizing entries for a particular column.

If you wish to add something to the drop down list, simply add it to your intended cell (with data validation disabled) and hit enter. When you activate the data validation afterwards your new entry will be embedded in the list.

### Tag and Descriptor Tables

This sheet contains a generated report that consists of data from a user-specified time interval. The sheet lays out in four sections horizontal of one another. All sections are color coordinated with a key at the top of each section. The sections are:

- All Tags
- Tags by Category
- Category and KPIs
- Other Identifiers

**All Tags**

This section of the sheet contains all the different tags used in tag columns as well as the frequency of their use, throughout the user-specified time interval.

**Tags by Category**

This section depicts charts the same as in the "All Tags" section, but it seperates contents based on the category where that tag was used, throughout the user-specified time interval.

**Category and KPIs***

A section with depicting the categorys and KPIs used and their frequency throughout the user-specified time interval.

**Other Identifiers**

A section containing all the other identifiers that were used and their frequency throughout the user-specified time interval.

_Tag Search Feature:_

On the left of the sheet is a button labeled "Test for Tag." In this cell, you can type words and if you press the button it will show you exact matches and matches partially containging your query below it, and where on the sheet it is located.

### Calendar Sheets

After running the Calendar macro, you will generate a very typical looking yearly calendar that has red and green cells. Red cells indicate that there was a data entry for the date of the cell. Green cells indicate there was no entry. 

At the top of the calendar (which is on a frozen pane) you can see what KPI's were potentially hit to change the cell's color. To the left of that is an "Update" button. In rder to not have to remake the Calendar every day, click the update button and it will refresh the calendars cells with the correct dates from the main data table.

### KPI Charts

Running the KPI Carts macro will generate a KPI Charts sheet with three main sections.

1. At the top will be a summarized chart of the KPIs that were hit for a choisen category over a chosen interval of time.
2. Directly below that are monthly charts indicating the exact days where a KPI was hit, if there was one, up to 10 total KPIs in one day.
3. To the right of section 1 will be a chart depicting the running total of KPIs hit over the specific time interval.

### Running Total Table and Trend Table
These sheets are generated from calling the "Create Pivot Table" macro. Depending on which the user selects, a sheet will be generated depicting a pivot table and a pivot chart of the desired information.

### Data Validation
It is imperative for this sheet to remain generally untouched. The data validation functionality of the Countermeasures sheet relies on this sheet to be named "Data Validation." No other contents should be manually added to this sheet as the sheet will autopopulate based on the data validation use on the main data table.

### Other Important Things to Know
When activating the "Input New Entries" the form requires the following inputs/formats to be entered for successful entry:

| Entry | DataType |
|---|---|
| Issue Date - Day | DD |
| Issue Date - Month | String |
| Issue Date - Year | YYYY |
| Category | String |
| KPI | String |
| Issue | String |
| Cause | String |
| Due Date - Day | MM |
| Due Date - Month | String |
| Due Date - Year | YYYY |
| Owner - First Name | String |
| Owner - Last Name | String |

Additionally, the "Category" drop down is autopopulated from main data table, but different entries can be manually input here as well. The "KPI", and all the tag and descriptor drop downs are populated from previous entries of that selected category. Of course, new entries can be manually input as well.









