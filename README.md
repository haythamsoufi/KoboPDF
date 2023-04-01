This script exports submissions of a Kobo form to PDF. The exported PDFs are named according to the namevar field in the form, the submission date, and the validation status of the submission. The script uses Selenium and the Chrome driver to retrieve the submissions from Kobo, and to export them to PDF.

#Dependencies
Python 3.x
selenium
requests
Chrome driver (download from https://chromedriver.chromium.org/ and provide the path to the driver in the driver_path variable)

#Function
The script exports submissions of a Kobo form to PDF.

Parameters
username : str - the Kobo username
password : str - the Kobo password
token : str - the Kobo API token
start_date : str - the start date in the format "YYYY-MM-DD"
end_date : str - the end date in the format "YYYY-MM-DD"

Usage
username = "kobo_username"
password = "kobo_password"
token = "kobo_api_token"
start_date = "2022-01-01"
end_date = "2022-12-31"
form_id = "kobo_form_id"
namevar = "namevar_field_name"
status = "Not Approved"
export_submissions_to_pdf(username, password, token, start_date, end_date)
Return value
The function does not return anything. The exported PDFs are saved to the current working directory.

Exceptions
If any exception occurs during the execution of the script, the exception is printed to the console.
Algorithm
Set up Chrome options and driver
Log in to Kobo
Retrieve the list of form submissions
Parse the JSON response
Export submissions to PDF
Save the current page as PDF
Print the name of the exported PDF and the elapsed time
Repeat steps 5-7 for all submissions
Print the total elapsed time

Example
username = "asdasd"
password = "123123"
token = "5e26fb155e26fb155e26fb1515e26fb1545325"
start_date = "2022-04-17"
end_date = "2022-09-18"
form_id = "anEnNVRefwerv329qJdAPg"
namevar = "group_Introduction/Name"
status = "Approved"

Notes
The script assumes that the form has a namevar field that contains a unique name for each submission.
The script uses the Page.printToPDF command from the Chrome DevTools Protocol to save the current page as PDF.
