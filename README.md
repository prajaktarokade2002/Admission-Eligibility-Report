# Admission-Eligibility-Report
The Admission Eligibility Report is a desktop-based application designed to automatically generate eligibility reports for students moving from one academic year/class to the next. The system reads academic credit information from university-issued PDF mark sheets and determines whether each student is eligible for admission to the next class 
<br>
The main objective of the project is to reduce manual effort and eliminate calculation errors in eligibility verification by automating PDF data extraction and report generation.

<h1>Key Features</h1>
✔ PDF Data Extraction

The system reads student information directly from PDF mark sheets using pdfplumber.

Extracted details include:

Student Name

Credit Total (FY or SY)

Exam Period

✔ Eligibility Checking

Admin defines credit limits (e.g., 22 to 44 credits).

The system checks:

FY Credits for FY → SY eligibility

FY + SY Credits for SY → TY eligibility

Automatically marks students as:

Eligible, or

Not Eligible

✔ Excel Report Generation

Generates a formatted Excel report containing:

Student Name

Total Credits (FY/SY)

Eligibility Status

Exam Period

Proper headings, borders, and alignment are applied for professional presentation.

✔ User-Friendly GUI

Built with Tkinter

GUI options include:

Select PDF file

Choose Academic Year and Class

Choose Exam Period (Mar-April / Oct-Nov)

Select Eligibility Type (FY→SY or SY→TY)

Set credit range values

Export to chosen output foldert and eliminate calculation errors in eligibility verification by automating PDF data extraction and report generation.
