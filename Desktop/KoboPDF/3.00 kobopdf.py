import os
import sys
import timeit
import traceback
import configparser
import webbrowser
from base64 import b64decode
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PyQt5.QtWidgets import (QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QHBoxLayout, QFileDialog, 
                             QProgressBar, QTextEdit, QMessageBox, QComboBox, QGridLayout, QStyle, QDesktopWidget)
from PyQt5.QtGui import QIcon, QPixmap  # For application and window icon handling
from PyQt5.QtCore import QThread, pyqtSignal, Qt  # Core functionality of PyQt (signals, slots, and threading)
from datetime import datetime, timedelta
from webdriver_manager.chrome import ChromeDriverManager

CURRENT_VERSION = "v3.1"

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# Read default values from the config file
def read_default_values():
    config_file_path = os.path.join(os.getcwd(), "config.ini")
    config = configparser.ConfigParser()
    if os.path.exists(config_file_path):
        config.read(config_file_path)
    else:
        default_values = {
            "username": "",
            "password": "",
            "token": "",
            "start_date": "2023-01-01",
            "end_date": "2023-12-31",
            "form_id": "",
            "namevar": "",
            "status": "All",
            "export_folder": os.path.join(os.getcwd(), "Export folder")
        }
        config["DEFAULT"] = default_values
        with open(config_file_path, "w") as configfile:
            config.write(configfile)
    return config["DEFAULT"]

# Check for updates from GitHub
def check_for_updates():
    repo_url = "https://api.github.com/repos/haythamsoufi/KoboPDF/releases/latest"
    try:
        response = requests.get(repo_url)
        response.raise_for_status()
        latest_release = response.json()
        latest_version = latest_release["tag_name"]
        release_notes = latest_release["body"]
        download_url = latest_release["html_url"]
        
        return latest_version, release_notes, download_url
    except Exception as e:
        print(f"Error checking for updates: {e}")
        return None, None, None

# Worker thread class for the export process
class ExportWorker(QThread):
    update_progress = pyqtSignal(str, int, bool)
    finished = pyqtSignal(bool, str)

    def __init__(self, username, password, token, start_date, end_date, form_id, namevar, status, export_path):
        super().__init__()
        self.username = username
        self.password = password
        self.token = token
        self.start_date = start_date
        self.end_date = end_date
        self.form_id = form_id
        self.namevar = namevar
        self.status = status
        self.export_path = export_path
        self.cancel_flag = False

    def run(self):
        driver = None
        total_elapsed_time = 0  # Track the total elapsed time from start to finish

        try:
            chrome_options = Options()
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--window-size=1920x1080")
            chrome_options.add_argument("--log-level=3")
            chrome_options.add_argument("--headless")
            
            # Use WebDriverManager to get the appropriate ChromeDriver
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

            start_time = timeit.default_timer()  # Start overall timing

            # Log in to Kobo
            driver.get("https://kobonew.ifrc.org/accounts/login/")
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.NAME, "login")))
            driver.find_element(By.NAME, "login").send_keys(self.username)
            driver.find_element(By.NAME, "password").send_keys(self.password)
            driver.find_element(By.NAME, "password").submit()

            login_elapsed_time = timeit.default_timer() - start_time

            # Retrieve form submissions
            submission_start_time = timeit.default_timer()
            end_date_obj = datetime.strptime(self.end_date, '%Y-%m-%d') + timedelta(days=1)
            end_date = end_date_obj.strftime('%Y-%m-%d')
            if self.status == "All":
                url = f"https://kobonew.ifrc.org/api/v2/assets/{self.form_id}/data.json?query={{\"_submission_time\":{{\"$gte\":\"{self.start_date}\",\"$lt\":\"{end_date}\"}}}}"
            else:
                url = f"https://kobonew.ifrc.org/api/v2/assets/{self.form_id}/data.json?query={{\"_submission_time\":{{\"$gte\":\"{self.start_date}\",\"$lt\":\"{end_date}\"}}, \"_validation_status.label\":\"{self.status}\"}}"
            headers = {"Authorization": f"Token {self.token}"}
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            results = response.json()["results"]
            submissions = [s["_id"] for s in results]
            num_files = len(submissions)

            submission_elapsed_time = timeit.default_timer() - submission_start_time

            self.update_progress.emit(f"Time taken to log in: {login_elapsed_time:.2f} seconds\n", 0, False)
            self.update_progress.emit(f"Time taken to retrieve submissions: {submission_elapsed_time:.2f} seconds\n\n", 0, False)
            self.update_progress.emit(f"Total submissions to be exported: {num_files}\n" + "-" * 40 + "\n", 0, False)

            total_time = 0  # Total time taken for exporting individual files

            for i, submission in enumerate(submissions):
                if self.cancel_flag:
                    self.finished.emit(True, "Export process canceled by the user.")
                    return

                file_start_time = timeit.default_timer()
                submission_url = f"https://kobonew.ifrc.org/api/v2/assets/{self.form_id}/data/{submission}/enketo/view/"
                response = requests.get(submission_url, headers=headers)
                response.raise_for_status()
                url = response.json()["url"]
                driver.get(url)
                WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "form-title")))
                driver.switch_to.window(driver.window_handles[-1])

                # Validate the necessary fields exist for naming the PDF
                name = results[i].get(self.namevar, "unknown")
                submission_time = results[i].get('_submission_time', 'unknown').split('T')[0].replace('-', '')
                validation_status = results[i].get('_validation_status', {})
                label = validation_status.get('label', '')

                # Construct the file name with conditional underscore for the label
                if label:
                    pdf_name = f"{name}_{submission_time}_{label}.pdf"
                else:
                    pdf_name = f"{name}_{submission_time}.pdf"

                if os.path.exists(os.path.join(self.export_path, pdf_name)):
                    file_number = 1
                    while True:
                        if label:
                            numbered_pdf_name = f"{name}_{submission_time}_{label} ({file_number}).pdf"
                        else:
                            numbered_pdf_name = f"{name}_{submission_time} ({file_number}).pdf"

                        if not os.path.exists(os.path.join(self.export_path, numbered_pdf_name)):
                            pdf_name = numbered_pdf_name
                            break
                        file_number += 1

                pdf_path = os.path.join(self.export_path, pdf_name)

                result = driver.execute_cdp_cmd("Page.printToPDF", {
                    "printBackground": True,
                    "marginTop": 0,
                    "marginRight": 0,
                    "marginBottom": 0,
                    "marginLeft": 0,
                    "preferCSSPageSize": True
                })

                with open(pdf_path, "wb") as f:
                    f.write(b64decode(result["data"]))

                driver.switch_to.window(driver.window_handles[0])

                file_elapsed_time = timeit.default_timer() - file_start_time
                total_time += file_elapsed_time
                progress = (i + 1) / num_files * 100
                self.update_progress.emit(f"{i + 1}. {pdf_name} [in {file_elapsed_time:.2f} seconds]\n", progress, False)

            total_elapsed_time = timeit.default_timer() - start_time  # Calculate the total time from start to finish
            self.finished.emit(False, f"Total elapsed time: {total_elapsed_time:.2f} seconds\nAll submissions were exported to PDF successfully.\n")

        except Exception as e:
            traceback.print_exc()
            self.finished.emit(True, f"Error exporting submissions to PDF: {str(e)}")

        finally:
            if driver:
                driver.quit()

# Main GUI class
class ExportSubmissionsToPdf(QWidget):

    def __init__(self):
        super().__init__()
        self.cancel_flag = False
        self.default_values = read_default_values()
        self.initUI()
        self.check_for_updates()

    def initUI(self):
        self.setWindowTitle(f"KoboPDF Tool {CURRENT_VERSION}")

        # Load the icon and set it for the app and window
        icon_path = resource_path("Kobopdf.png")
        self.setWindowIcon(QIcon(icon_path))
        
        # Set window size and center it
        self.setGeometry(100, 100, 700, 500)
        self.center_window()

        main_layout = QVBoxLayout()

        # Title and GitHub link in a horizontal layout, centered
        title_github_layout = QHBoxLayout()
        title_github_layout.setAlignment(Qt.AlignCenter)

        title_label = QLabel(f"KoboPDF Tool {CURRENT_VERSION}")
        title_label.setStyleSheet("font-size: 20px; font-weight: bold;")
        title_github_layout.addWidget(title_label)

        github_link = QLabel('<a href="https://github.com/haythamsoufi/KoboPDF">KoboPDF on GitHub</a>')
        github_link.setOpenExternalLinks(True)
        title_github_layout.addWidget(github_link)

        main_layout.addLayout(title_github_layout)

        # Create grid layout for form fields with two columns
        grid_layout = QGridLayout()

        # Server configuration section
        grid_layout.addWidget(QLabel("Server Configuration"), 0, 0, 1, 2, Qt.AlignLeft)
        grid_layout.addWidget(QLabel("Username:"), 1, 0)
        self.username = QLineEdit()
        self.username.setText(self.default_values.get("username", ""))
        grid_layout.addWidget(self.username, 1, 1)

        grid_layout.addWidget(QLabel("Password:"), 2, 0)
        self.password = QLineEdit()
        self.password.setEchoMode(QLineEdit.Password)
        self.password.setText(self.default_values.get("password", ""))
        grid_layout.addWidget(self.password, 2, 1)

        grid_layout.addWidget(QLabel("Token:"), 3, 0)
        self.token = QLineEdit()
        self.token.setText(self.default_values.get("token", ""))
        grid_layout.addWidget(self.token, 3, 1)

        # Export settings section
        grid_layout.addWidget(QLabel("Export Settings"), 0, 2, 1, 2, Qt.AlignLeft)
        grid_layout.addWidget(QLabel("Start Date:"), 1, 2)
        self.start_date = QLineEdit()
        self.start_date.setText(self.default_values.get("start_date", "2023-01-01"))
        grid_layout.addWidget(self.start_date, 1, 3)

        grid_layout.addWidget(QLabel("End Date:"), 2, 2)
        self.end_date = QLineEdit()
        self.end_date.setText(self.default_values.get("end_date", "2023-12-31"))
        grid_layout.addWidget(self.end_date, 2, 3)

        grid_layout.addWidget(QLabel("Form ID:"), 3, 2)
        self.form_id = QLineEdit()
        self.form_id.setText(self.default_values.get("form_id", ""))
        grid_layout.addWidget(self.form_id, 3, 3)

        grid_layout.addWidget(QLabel("Namevar:"), 4, 2)
        namevar_layout = QHBoxLayout()
        self.namevar = QLineEdit()
        self.namevar.setText(self.default_values.get("namevar", ""))
        info_btn = QPushButton()
        info_btn.setIcon(self.style().standardIcon(QStyle.SP_MessageBoxInformation))
        info_btn.clicked.connect(self.show_info)
        namevar_layout.addWidget(self.namevar)
        namevar_layout.addWidget(info_btn)
        grid_layout.addLayout(namevar_layout, 4, 3)

        grid_layout.addWidget(QLabel("Status:"), 5, 2)
        self.status = QComboBox()
        self.status.addItems(['All', 'Approved', 'Not Approved', 'On Hold'])
        self.status.setCurrentText(self.default_values.get("status", "Approved"))
        grid_layout.addWidget(self.status, 5, 3)

        grid_layout.addWidget(QLabel("Export Folder:"), 6, 2)
        export_folder_layout = QHBoxLayout()
        self.export_folder = QLineEdit()
        self.export_folder.setText(self.default_values.get("export_folder", os.path.join(os.getcwd(), "Export folder")))
        export_folder_btn = QPushButton()
        export_folder_btn.setIcon(self.style().standardIcon(QStyle.SP_DirIcon))  # Use standard directory icon
        export_folder_btn.clicked.connect(self.select_export_folder)
        export_folder_layout.addWidget(self.export_folder)
        export_folder_layout.addWidget(export_folder_btn)
        grid_layout.addLayout(export_folder_layout, 6, 3)

        main_layout.addLayout(grid_layout)

        # Add buttons
        btn_layout = QHBoxLayout()
        self.export_btn = QPushButton("Export")
        self.export_btn.setIcon(self.style().standardIcon(QStyle.SP_DialogOkButton))  # Use standard export icon
        self.export_btn.clicked.connect(self.on_export)
        save_btn = QPushButton("Save Configuration")
        save_btn.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))  # Use standard save icon
        save_btn.clicked.connect(self.on_save_config)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.setIcon(self.style().standardIcon(QStyle.SP_DialogCancelButton))  # Use standard cancel icon
        cancel_btn.clicked.connect(self.close)

        btn_layout.addWidget(self.export_btn)
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)

        main_layout.addLayout(btn_layout)

        # Progress and output
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setMaximum(100)
        self.output_text = QTextEdit()
        self.output_text.setReadOnly(True)

        main_layout.addWidget(self.progress_bar)
        main_layout.addWidget(self.output_text)

        self.setLayout(main_layout)

    def center_window(self):
        # Center the window on the screen
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def select_export_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Export Folder")
        if folder:
            self.export_folder.setText(folder)

    def on_save_config(self):
        config = configparser.ConfigParser()
        config.read('config.ini')

        config.set('DEFAULT', 'username', self.username.text())
        config.set('DEFAULT', 'password', self.password.text())
        config.set('DEFAULT', 'token', self.token.text())
        config.set('DEFAULT', 'start_date', self.start_date.text())
        config.set('DEFAULT', 'end_date', self.end_date.text())
        config.set('DEFAULT', 'form_id', self.form_id.text())
        config.set('DEFAULT', 'namevar', self.namevar.text())
        config.set('DEFAULT', 'status', self.status.currentText())
        config.set('DEFAULT', 'export_folder', self.export_folder.text())  # Save the export folder path

        with open('config.ini', 'w') as configfile:
            config.write(configfile)

        QMessageBox.information(self, "Info", "Configuration saved successfully")

    def on_export(self):
        self.export_btn.setEnabled(False)  # Disable the export button

        username = self.username.text()
        password = self.password.text()
        token = self.token.text()
        start_date = self.start_date.text()
        end_date = self.end_date.text()
        form_id = self.form_id.text()
        namevar = self.namevar.text()
        status = self.status.currentText()
        export_path = self.export_folder.text()

        # Check if the export folder exists, if not, create it
        if not os.path.exists(export_path):
            os.makedirs(export_path)

        self.output_text.clear()
        self.progress_bar.setValue(0)

        self.export_worker = ExportWorker(username, password, token, start_date, end_date, form_id, namevar, status, export_path)
        self.export_worker.update_progress.connect(self.update_status)
        self.export_worker.finished.connect(self.on_export_finished)
        self.export_worker.start()

    def update_status(self, text, progress, finished):
        self.output_text.append(text)
        self.progress_bar.setValue(progress)
        if finished:
            # Remove the info message box here
            self.export_btn.setEnabled(True)  # Re-enable the export button

    def on_export_finished(self, error_occurred, message):
        self.update_status(message, 100, True)
        self.export_btn.setEnabled(True)  # Re-enable the export button
        if not error_occurred:
            QMessageBox.information(self, "Success", "Export completed successfully")
        else:
            QMessageBox.critical(self, "Error", message)

    def show_info(self):
        QMessageBox.information(self, "Namevar Information", 
            "This is the name of the field in your Kobo form that contains the name of Country/National Society/etc.\n\n"
            "This will be used to name the output PDF file.\n\n"
            "You can find the names by accessing this link:\n"
            "https://kobonew.ifrc.org/api/v2/assets/[YOUR FORM ID]/data.json and copy the name of the field you want to use\n\n"
            "You can use '_id' if you don't have a specific name field.")

    def check_for_updates(self):
        latest_version, release_notes, download_url = check_for_updates()
        if latest_version and latest_version != CURRENT_VERSION:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle("Update Available")
            msg.setText(f"A new version ({latest_version}) is available.\n\nRelease notes:\n{release_notes}")
            msg.setInformativeText("Do you want to visit the download page?")
            msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            msg.setDefaultButton(QMessageBox.Yes)
            retval = msg.exec_()
            if retval == QMessageBox.Yes:
                webbrowser.open(download_url)

# Main function to run the application
if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = ExportSubmissionsToPdf()
    ex.show()
    sys.exit(app.exec_())