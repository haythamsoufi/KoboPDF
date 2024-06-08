import os
import timeit
import requests
import traceback
import configparser
from base64 import b64decode
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import wx
import wx.lib.delayedresult as delayedresult
from threading import Thread

def export_submissions_to_pdf(username, password, token, start_date, end_date, form_id, namevar, status, export_path, cancel_flag, status_window):
    
    # Read configuration file
    config = configparser.ConfigParser()
    config.read('C:/Users/Haytham.Alsoufi/Desktop/PDFF/config.ini')
    driver_path = config.get('DEFAULT', 'driver_path')
    export_folder = os.getcwd()

    window = None

    # Initialize output_text
    output_text = ""

      
    # Set up Chrome options and driver
    chrome_options = Options()
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920x1080")
    chrome_options.add_argument("--log-level=3")
    chrome_options.add_argument("--headless")
    service = Service(executable_path=driver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)

    start_time = timeit.default_timer()

    try:
        # Log in to Kobo
        login_start_time = timeit.default_timer()
        driver.get("https://kobonew.ifrc.org/accounts/login/")
        wait = WebDriverWait(driver, 30)
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.NAME, "login")))
        driver.find_element(By.NAME, "login").send_keys(username)
        driver.find_element(By.NAME, "password").send_keys(password)
        driver.find_element(By.NAME, "password").submit()

        login_elapsed_time = timeit.default_timer() - login_start_time

        # Retrieve the list of form submissions
        submission_start_time = timeit.default_timer()
        url = f"https://kobonew.ifrc.org/api/v2/assets/{form_id}/data.json?query={{\"_submission_time\":{{\"$gte\":\"{start_date}\",\"$lte\":\"{end_date}\"}}, \"_validation_status.label\":\"{status}\"}}"
        headers = {"Authorization": f"Token {token}"}
        response = requests.get(url, headers=headers)
        response.raise_for_status()

        submission_elapsed_time = timeit.default_timer() - submission_start_time

        # Parse the JSON response
        results = response.json()["results"]
        submissions = [s["_id"] for s in results]
        total_time = 0
        print(f"Time taken to log in: {login_elapsed_time:.2f} seconds")
        print(f"Time taken to retrieve submissions: {submission_elapsed_time:.2f} seconds\n")

        # Export submissions to PDF
        export_start_time = timeit.default_timer()

        for i, submission in enumerate(submissions):
            
            start_time = timeit.default_timer()
            submission_url = f"https://kobonew.ifrc.org/api/v2/assets/{form_id}/data/{submission}/enketo/view/"
            response = requests.get(submission_url, headers=headers)
            response.raise_for_status()

            # Parse the JSON response
            url = response.json()["url"]

            driver.get(url)
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "form-title")))

            driver.switch_to.window(driver.window_handles[-1])

            # Set the desired PDF name
            pdf_name = f"{results[i][namevar]}_{results[i]['_submission_time'][:10].replace('-', '')}_{results[i]['_validation_status']['label']}.pdf"
            pdf_path = os.path.join(export_folder, pdf_name)

            # Save the current page as PDF
            
            result = driver.execute_cdp_cmd("Page.printToPDF", {
            "printBackground": True,
            "marginTop": 0,
            "marginRight": 0,
            "marginBottom": 0,
            "marginLeft": 0,
            "preferCSSPageSize": True
            })     

            # Calculate the elapsed time
            elapsed_time = timeit.default_timer() - start_time
            total_time += elapsed_time
            wx.CallAfter(status_window.update_status, f"{i+1}. {pdf_name} [in {elapsed_time:.2f} seconds]\n", cancel_flag)  # pass cancel_flag to update_status


            with open(pdf_path, "wb") as f:
                f.write(b64decode(result["data"]))

            driver.switch_to.window(driver.window_handles[0])

            # Add the printed output to the output window
            output_text += f"{i+1}. {pdf_name} [in {elapsed_time:.2f} seconds]\n"

            if cancel_flag[0]:
                output_text += f"Export process canceled by the user.\n"
                status_window.update_status("Export process canceled by the user.", cancel_flag)
                break

        else:
            output_text += f"All submissions exported successfully!\n"
            status_window.update_status("All submissions exported successfully!", cancel_flag)

        # Calculate the elapsed time
        total_elapsed_time = timeit.default_timer() - start_time

        # Close the Chrome driver
        driver.quit()

        export_time = timeit.default_timer()
        total_time = export_time - start_time
        export_elapsed_time = timeit.default_timer() - export_start_time
        total_elapsed_time = login_elapsed_time + submission_elapsed_time + export_elapsed_time

        num_files = len(submissions)
        print(f"\nAll {num_files} submissions exported to PDF successfully")
        print(f"Total elapsed time: {total_elapsed_time:.2f} seconds")


    except Exception as e:
        traceback.print_exc()
        print(f"Error exporting submissions to PDF: {str(e)}")

    finally:
        driver.quit()

class ExportSubmissionsToPdf(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, title="Export Submissions to PDF", size=(500, 500))
        self.cancel_flag = [False]

        self.InitUI()
        self.Centre()
        self.Show()
        
    def InitUI(self, username='', password='', token='', start_date='', end_date='', form_id='', namevar='', status='', export_path='', cancel_flag=None):
        panel = wx.Panel(self)

        cancel_flag = [False] if cancel_flag is None else cancel_flag
        
        vbox = wx.BoxSizer(wx.VERTICAL)
        self.window = self.GetParent()

        title = wx.StaticText(panel, label="Kobo to PDF Tool")
        title_font = title.GetFont()
        title_font.SetPointSize(20)
        title_font.SetWeight(wx.BOLD)
        title.SetFont(title_font)
        vbox.Add(title, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, border=20)
        vbox.Add((-1, 10))

        desc = wx.StaticText(panel, label="Enter your credentials and export settings:")
        desc_font = desc.GetFont()
        desc_font.SetPointSize(12)
        desc.SetFont(desc_font)
        vbox.Add(desc, flag=wx.EXPAND | wx.LEFT | wx.RIGHT, border=20)
        vbox.Add((-1, 20))

        # Create a grid sizer for the form
        grid = wx.FlexGridSizer(9, 2, 10, 10)
        grid.AddGrowableCol(1, 1)

        # Add the form fields
        labels = ['Username:', 'Password:', 'Token:', 'Start Date:', 'End Date:', 'Form ID:', 'Namevar:', 'Status:', 'Export Folder:']
        self.username = wx.TextCtrl(panel, value=username)
        self.password = wx.TextCtrl(panel, value=password, style=wx.TE_PASSWORD)
        self.token = wx.TextCtrl(panel, value=token)
        self.start_date = wx.TextCtrl(panel, value=start_date)
        self.end_date = wx.TextCtrl(panel, value=end_date)
        self.form_id = wx.TextCtrl(panel, value=form_id)
        self.namevar = wx.TextCtrl(panel, value=namevar)
        self.status = wx.ComboBox(panel, choices=['Approved', 'Not Approved', 'On Hold'], style=wx.CB_READONLY, value=status)
        self.export_folder = wx.DirPickerCtrl(panel, path=export_path)

        # Set default values from config.ini
        config = configparser.ConfigParser()
        config.read('C:/Users/Haytham.Alsoufi/Desktop/PDFF/config.ini')
        self.username.SetValue(config.get('DEFAULT', 'username'))
        self.password.SetValue(config.get('DEFAULT', 'password'))
        self.token.SetValue(config.get('DEFAULT', 'token'))
        self.start_date.SetValue(config.get('DEFAULT', 'start_date'))
        self.end_date.SetValue(config.get('DEFAULT', 'end_date'))
        self.form_id.SetValue(config.get('DEFAULT', 'form_id'))
        self.namevar.SetValue(config.get('DEFAULT', 'namevar'))
        self.status.SetValue(config.get('DEFAULT', 'status'))

        ctrls = [self.username, self.password, self.token, self.start_date, self.end_date, self.form_id, self.namevar, self.status, self.export_folder]
        
        for label, ctrl in zip(labels, ctrls):
            grid.Add(wx.StaticText(panel, label=label), flag=wx.ALIGN_CENTER_VERTICAL)
            grid.Add(ctrl, flag=wx.EXPAND)

        vbox.Add(grid, proportion=1, flag=wx.EXPAND | wx.LEFT | wx.RIGHT, border=20)
        vbox.Add((-1, 20))

        hbox = wx.BoxSizer(wx.HORIZONTAL)
        export_btn = wx.Button(panel, label="Export")
        hbox.Add(export_btn, flag=wx.RIGHT, border=10)

        self.cancel_button = wx.Button(panel, label="Cancel")
        hbox.Add(self.cancel_button)
        vbox.Add(hbox, flag=wx.ALIGN_CENTER | wx.BOTTOM, border=20)

        panel.SetSizer(vbox)

        # Bind events
        export_btn.Bind(wx.EVT_BUTTON, self.on_export)
        self.cancel_button.Bind(wx.EVT_BUTTON, self.on_cancel_button)

    def on_export(self, event):
        # Read input values from the form
        username = self.username.GetValue()
        password = self.password.GetValue()
        token = self.token.GetValue()
        start_date = self.start_date.GetValue()
        end_date = self.end_date.GetValue()
        form_id = self.form_id.GetValue()
        namevar = self.namevar.GetValue()
        status = self.status.GetValue()
        export_path = self.export_folder.GetPath()

        # Get configuration values
        config = configparser.ConfigParser()
        config.read('C:/Users/Haytham.Alsoufi/Desktop/PDFF/config.ini')
        driver_path = config.get('DEFAULT', 'driver_path')

        # Hide the main window
        self.Hide()

        # Show the export status window
        self.export_status_window = ExportStatusWindow(parent=self)
        self.export_status_window.Show()

        # Run the export_submissions_to_pdf function in a separate thread
        worker = WorkerThread(self.export_submissions_to_pdf_thread, username, password, token, start_date, end_date, form_id, namevar, status, export_path, cancel_flag=self.cancel_flag)
        worker.start()

    def export_submissions_to_pdf_thread(self, username, password, token, start_date, end_date, form_id, namevar, status, export_path, cancel_flag):
        # Run the export_submissions_to_pdf function in a separate thread
        export_submissions_to_pdf(username, password, token, start_date, end_date, form_id, namevar, status, export_path, cancel_flag, self.export_status_window)

        # Close the status window
        self.export_status_window.Destroy()

        # Close the main window
        self.Close()

    def on_cancel_button(self, event):
        self.cancel_flag[0] = True
        self.Close()

class WorkerThread(Thread):
   
    def __init__(self, func, *args):
        Thread.__init__(self)
        self.func = func
        self.args = args

    def run(self):
        self.func(*self.args)

class ExportStatusWindow(wx.Frame):
    def __init__(self, parent=None):
        wx.Frame.__init__(self, parent, title="Export Status", size=(400, 300))
        panel = wx.Panel(self)

        vbox = wx.BoxSizer(wx.VERTICAL)
        self.status_label = wx.StaticText(panel, label="Exporting submissions:")
        vbox.Add(self.status_label, flag=wx.EXPAND | wx.ALL, border=20)

        self.output_text = wx.TextCtrl(panel, style=wx.TE_READONLY | wx.TE_MULTILINE)
        vbox.Add(self.output_text, proportion=1, flag=wx.EXPAND | wx.LEFT | wx.RIGHT, border=20)

        hbox = wx.BoxSizer(wx.HORIZONTAL)
        self.cancel_button = wx.Button(panel, label="Cancel")
        hbox.Add(self.cancel_button, flag=wx.ALIGN_CENTER | wx.RIGHT, border=10)
        vbox.Add(hbox, flag=wx.ALIGN_CENTER | wx.BOTTOM, border=20)

        panel.SetSizer(vbox)

        # Bind events
        self.cancel_button.Bind(wx.EVT_BUTTON, self.on_cancel_button)

    def update_status(self, text, cancel_flag=None):
        if self.output_text:
            self.output_text.AppendText(text)
            if cancel_flag and cancel_flag[0]:
                self.update_status("Export process canceled by the user.\n")
                return

    def on_cancel_button(self, event):
        self.GetParent().cancel_flag[0] = True
        self.Close()

app = wx.App()
ExportSubmissionsToPdf()
app.MainLoop()