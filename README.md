# Admission-Eligibility-Report
The Admission Eligibility Report is a desktop-based application designed to automatically generate eligibility reports for students moving from one academic year/class to the next. The system reads academic credit information from university-issued PDF mark sheets and determines whether each student is eligible for admission to the next class 
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


