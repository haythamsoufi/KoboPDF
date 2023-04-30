The KoboPDF Tool is a desktop application that allows users to export Kobo form submissions to PDF format. The tool provides a user-friendly interface that allows users to enter their credentials, choose a Kobo form, and select a date range to export submissions to PDF format.

# Dependencies
1. Python 3.x
2. selenium
3. requests
4. Chrome driver (https://chromedriver.chromium.org/)

# Function
The script exports submissions of a Kobo form to PDF.

# Usage
To use the KoboPDF Tool, follow these steps:

1. Open the application.
2. Enter your Kobo credentials (username and password) or a Kobo API token.
3. Choose a Kobo form by entering the form ID.
4. Optionally, enter a name variable that corresponds to the field in your Kobo form that contains the name of Country/National Society/etc.. This will be used to name the output PDF file.
5. Select a date range to export submissions by entering the start and end dates.
6. Choose an export folder where the PDF files will be saved.
7. Select a status to export ("All", "Approved", "Not Approved", or "On Hold").
8. Click the "Export" button to start the export process.

While the export process is running, a progress window will be displayed, showing the progress of the export process. The progress window also provides an option to cancel the export process.

# Configuration
The KoboPDF Tool allows you to save your configuration settings for future use. To save your configuration settings, click the "Save configuration" button. This will save your settings in the config.ini file.

# Troubleshooting
If you encounter any issues while using the KoboPDF Tool, please create an issue on the GitHub repository.
