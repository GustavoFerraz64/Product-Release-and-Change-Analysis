## Analysis of Product Release and Modification (LAP)

## Project Description

This script aims to automate the extraction of the database (`RelatorioLAP.csv`) of LAPs through AIT, analyze these LAPs, and generate an Excel file (`RelatorioLAP_Analisado.xlsx`) that will serve as the database for the BI dashboard `LAPs_BI6`. Additionally, it creates an Excel file with overdue LAPs since November 2024, including a field for a DPCP (Production Planning and Control Department) analyst to record the reason for the delay. Finally, the script automates sending emails to departments with overdue LAPs or those expected to be delayed within a week.

This project was developed to meet audit requirement 11.2.

## First Stage of the Script: Extract LAP Report from AIT

Initially, the script checks and deletes the file `RelatorioLAP.csv` from the user's downloads folder, ensuring the script runs smoothly. Then, Google Chrome is opened, and the TKE AIT website is accessed. On the portal, the path followed is: `Manufacturing Management > Engineering Change Management > LAP > Excel Report`, where the Excel report is downloaded.

## Second Stage of the Script: Analyze LAP Data

After downloading the report from AIT, Pandas reads the CSV file in the user's Downloads folder and the TKE Calendar to define working days. Pandas performs the following steps on the `RelatorioLAP.csv` file:

- Processes the dates in the file  
- Calculates the Lead Time for each LAP  
  The calculation is done by subtracting the release date by DPCP (the last department) from the release date by engineering for each LAP.  
- Determines whether there was a delay in the engineering release  
- Defines the number of departments for each LAP  
- Creates the column indicating the LAP release deadline, according to the criteria established by DPCP.  

|Category | Deadline (Days) |
|--|--|
|Production Adjustment | 60 |
|Critical Correction | 35 |
|New Product Release | 60 |
|Application Table | 45 |
|Approval | 60 |

- Adjusts the deadline column. If the due date falls on a non-working day, the day is adjusted to the nearest previous working day  
- Identifies which LAPs are overdue based on the established deadlines  
- Calculates Lead Time per department  
  For each LAP, the Lead Time of the departments is calculated. For independent departments (departments without a "Parent"), the release date of the LAP by the department (`Data Lib EA` column) is subtracted from the engineering release date (`Data de liberação` column).  
  For dependent departments (departments with a "Parent"), the release date of the LAP by the department is subtracted from the release date of the LAP by the parent department.  
- Defines the deadline for each department  
  The DPCP deadline is always the LAP deadline (DPCP is always the last department). The critical path of the LAP is defined, meaning the sequence of departments that depend on each other is determined. Based on this path, the deadline for each department is calculated by subtracting the LAP deadline, according to its reason, by the number of departments in this path. In the end, the due date for each department is set.  
  Departments that do not impact other departments (departments not listed in the "Parent" column for this LAP) have their deadline equal to the longest deadline of any department in this LAP, except for DPCP.  
  If a LAP has an empty "Parent" column, meaning it does not have departments that impact others, the department deadlines are evenly divided in days between DPCP and the remaining departments, so all departments have the same number of days to release the LAP.  
- Adjusts the department deadlines column. If the due date falls on a non-working day, the day is adjusted to the nearest previous working day.  
- Updates the delay justification spreadsheet.  
  At this stage, all overdue LAPs or those delivered late are stored in a spreadsheet that serves as a database for delay justifications. In this spreadsheet, there is a column where a DPCP analyst must justify the delay.  
- Saves the final report  
  The final report with the analyses performed is saved as `RelatorioLAP_Analisado.xlsx` in the path `\\srvfile01.tsur.local\DADOS_PBI\Compartilhado_BI\DPCP\25. LAPs`.  

## Third Stage of the Script: Sending Emails to Departments

In this final stage, emails are sent to the departments.  
First, an email is sent to the analyst responsible for LAPs in DPCP if there is any LAP with an unrecorded delay justification.  
Then, emails are sent to the remaining departments that have overdue LAPs or LAPs with a deadline within a week.  
