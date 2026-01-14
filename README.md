# Admission-Eligibility-Report Generator

A Python-based desktop application that reads PDF mark sheets, extracts student credit data, and generates admission eligibility reports in Excel format. It helps academic departments automate eligibility checks for FY → SY and SY → TY admissions.
<br>
The main objective of the project is to reduce manual effort and eliminate calculation errors in eligibility verification by automating PDF data extraction and report generation.

<h1>Key Features</h1>
<h3>PDF Data Extraction</h3>

The system reads student information directly from PDF mark sheets using pdfplumber.

Extracted details include:

Student Name

Credit Total (FY or SY)

Exam Period

<h3> Eligibility Checking</h3>

Admin defines credit limits (e.g., 22 to 44 credits).

The system checks:

FY Credits for FY → SY eligibility

FY + SY Credits for SY → TY eligibility

Automatically marks students as:

Eligible, or

Not Eligible

<h3> Excel Report Generation</h3>

Generates a formatted Excel report containing:

Student Name

Total Credits (FY/SY)

Eligibility Status

Exam Period

Proper headings, borders, and alignment are applied for professional presentation.

<h1> User-Friendly GUI</h1>.
<img width="566" height="340" alt="image" src="https://github.com/user-attachments/assets/7ec92648-8d9a-48f7-837e-e79b9b9a8109" />
<br>
<br>

<img width="401" height="152" alt="image" src="https://github.com/user-attachments/assets/c3fa4366-8f82-4cf2-b1dc-81e9c736126a" />
<br>
<br>
ouput in excel file
<br>
<img width="558" height="424" alt="image" src="https://github.com/user-attachments/assets/83106248-8d63-4889-99af-8d7c4d63bebc" />



