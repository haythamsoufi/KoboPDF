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
from threading import Thread
import wx.adv

def export_submissions_to_pdf(username, password, token, start_date, end_date, form_id, namevar, status, export_path, cancel_flag, status_window):
    
    config_file = 'config.ini'
    if not os.path.exists(config_file):
        # Create a new config.ini file with default values if it doesn't exist
        with open(config_file, 'w') as f:
            f.write("[DEFAULT]\n")
        # Read configuration file
    config = configparser.ConfigParser()
    config.read(config_file)
    driver_path = config.get('DEFAULT', 'driver_path')
    export_folder = os.path.join(os.getcwd(), "Export folder")

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
        num_files = len(submissions)
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
            pdf_path = os.path.join(export_path, pdf_name)

            # Create the Export folder if it does not exist
            if not os.path.exists(export_folder):
                os.makedirs(export_folder)

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
            progress = (i + 1) / len(submissions) * 100
            wx.CallAfter(status_window.update_status, f"{i+1}. {pdf_name} [in {elapsed_time:.2f} seconds]\n", progress, cancel_flag, num_files)


            with open(pdf_path, "wb") as f:
                f.write(b64decode(result["data"]))

            driver.switch_to.window(driver.window_handles[0])

            # Add the printed output to the output window
            output_text += f"{i+1}. {pdf_name} [in {elapsed_time:.2f} seconds]\n"

            if cancel_flag[0]:
                break

        # Calculate the elapsed time
        total_elapsed_time = timeit.default_timer() - start_time

        # Close the Chrome driver
        driver.quit()

        export_time = timeit.default_timer()
        total_time = export_time - start_time
        export_elapsed_time = timeit.default_timer() - export_start_time
        total_elapsed_time = login_elapsed_time + submission_elapsed_time + export_elapsed_time

        num_files = len(submissions)
        output_text += f"All {num_files} submissions exported to PDF successfully\n"
        output_text += f"Total elapsed time: {total_elapsed_time:.2f} seconds\n"


    except Exception as e:
        traceback.print_exc()
        print(f"Error exporting submissions to PDF: {str(e)}")

    finally:
        driver.quit()

class ExportSubmissionsToPdf(wx.Frame):

    def create_hyperlink(self, panel, url):
        link = wx.adv.HyperlinkCtrl(panel, label=url, url=url)
        return link

    def __init__(self):
        wx.Frame.__init__(self, None, title="Export Submissions to PDF", size=(500, 540))
        icon = wx.Icon("C:/Users/Haytham.Alsoufi/Desktop/PDFF/Kobopdf.ico", wx.BITMAP_TYPE_ICO)
        self.SetIcon(icon)
        self.cancel_flag = [False]

        self.InitUI()
        self.Centre()
        self.Show()
        
    def InitUI(self, username='', password='', token='', start_date='', end_date='', form_id='', namevar='', status='', export_path='', cancel_flag=None):
        panel = wx.Panel(self)

        cancel_flag = [False] if cancel_flag is None else cancel_flag
        
        vbox = wx.BoxSizer(wx.VERTICAL)
        self.window = self.GetParent()
        
        title = wx.StaticText(panel, label="KoboPDF Tool")
        title_font = title.GetFont()
        title_font.SetPointSize(20)
        title_font.SetWeight(wx.BOLD)
        title.SetFont(title_font)
        vbox.Add(title, flag=wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT | wx.TOP, border=20)
        vbox.Add((-1, 0))

        # Add a hyperlink to the Kobo website
        link = self.create_hyperlink(panel, "https://github.com/haythamsoufi/KoboPDF")
        vbox.Add(link, flag=wx.ALIGN_CENTER  | wx.LEFT, border=10)
        vbox.Add((-1, 10))

        desc = wx.StaticText(panel, label="Enter your credentials and export settings:")
        desc_font = desc.GetFont()
        desc_font.SetPointSize(12)
        desc.SetFont(desc_font)
        vbox.Add(desc, flag=wx.ALIGN_CENTER | wx.LEFT | wx.RIGHT, border=20)
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
        export_folder = os.path.join(os.getcwd(), "Export folder")
        self.export_folder = wx.DirPickerCtrl(panel, path=export_folder)

        # Set default values from config.ini
        config = configparser.ConfigParser()
        config.read('config.ini')
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

        vbox.Add(grid, proportion=1, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, border=20)

        hbox = wx.BoxSizer(wx.HORIZONTAL)
        export_btn = wx.Button(panel, label="Export")
        hbox.Add(export_btn, flag=wx.ALIGN_CENTER, border=10)

        self.save_button = wx.Button(panel, label="Save configuration")
        hbox.Add(self.save_button, flag=wx.LEFT, border=10)

        self.cancel_button = wx.Button(panel, label="Cancel")
        hbox.Add(self.cancel_button, flag=wx.LEFT, border=10)

        vbox.Add(hbox, flag=wx.ALIGN_CENTER | wx.BOTTOM, border=20)

        panel.SetSizer(vbox)

        # Bind events
        export_btn.Bind(wx.EVT_BUTTON, self.on_export)
        self.cancel_button.Bind(wx.EVT_BUTTON, self.on_cancel_button)
        self.save_button.Bind(wx.EVT_BUTTON, self.on_save_button)

    def on_save_button(self, event):
        config = configparser.ConfigParser()

        # Read the current configuration file
        config.read('config.ini')

        # Update the config with the entered values
        config.set('DEFAULT', 'username', self.username.GetValue())
        config.set('DEFAULT', 'password', self.password.GetValue())
        config.set('DEFAULT', 'token', self.token.GetValue())
        config.set('DEFAULT', 'start_date', self.start_date.GetValue())
        config.set('DEFAULT', 'end_date', self.end_date.GetValue())
        config.set('DEFAULT', 'form_id', self.form_id.GetValue())
        config.set('DEFAULT', 'namevar', self.namevar.GetValue())
        config.set('DEFAULT', 'status', self.status.GetValue())

        # Write the updated configuration back to the file
        with open('config.ini', 'w') as configfile:
            config.write(configfile)

        # Show a message box to inform the user that the data has been saved
        wx.MessageBox("Configuration saved successfully", "Info", wx.OK | wx.ICON_INFORMATION)


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
        config.read('config.ini')
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

        # Check if the ExportStatusWindow is still open before destroying it
        if self.export_status_window and not self.export_status_window.IsBeingDeleted():
            self.export_status_window.Destroy()

    def on_cancel_button(self, event):
        self.cancel_flag[0] = True
        self.Close()

class WorkerThread(Thread):
   
    def __init__(self, func, *args, cancel_flag=None):
        Thread.__init__(self)
        self.func = func
        self.args = args
        self.cancel_flag = cancel_flag

    def run(self):
        self.func(*self.args, cancel_flag=self.cancel_flag)

class ExportStatusWindow(wx.Frame):
    def __init__(self, parent=None):
        wx.Frame.__init__(self, parent, title="Exporting submissions", size=(500, 500))
        icon = wx.Icon("C:/Users/Haytham.Alsoufi/Desktop/PDFF/Kobopdf.ico", wx.BITMAP_TYPE_ICO)
        self.SetIcon(icon)
        self.Centre()
        panel = wx.Panel(self)

        vbox = wx.BoxSizer(wx.VERTICAL)
        vbox.AddSpacer(20)

        self.gauge = wx.Gauge(panel, range=100, size=(-1, 25))
        vbox.Add(self.gauge, flag=wx.EXPAND | wx.LEFT | wx.RIGHT, border=20)
        vbox.AddSpacer(5)

        self.total_files_text = wx.StaticText(panel, label=f"Total submissions to be exported: {0}")
        vbox.Add(self.total_files_text, flag=wx.EXPAND | wx.LEFT | wx.RIGHT, border=20)
        vbox.AddSpacer(5)

        self.output_text = wx.TextCtrl(panel, style=wx.TE_READONLY | wx.TE_MULTILINE)
        vbox.Add(self.output_text, proportion=1, flag=wx.EXPAND | wx.LEFT | wx.RIGHT, border=20)
        vbox.AddSpacer(20)

        hbox = wx.BoxSizer(wx.HORIZONTAL)
        self.cancel_button = wx.Button(panel, label="Stop")
        hbox.Add(self.cancel_button, flag=wx.ALIGN_CENTER | wx.RIGHT, border=10)
        self.exit_button = wx.Button(panel, label="Close")
        self.exit_button.Enable(False)
        hbox.Add(self.exit_button, flag=wx.ALIGN_CENTER | wx.LEFT, border=10)
        vbox.Add(hbox, flag=wx.ALIGN_CENTER | wx.BOTTOM, border=20)

        panel.SetSizer(vbox)

        # Bind events
        self.cancel_button.Bind(wx.EVT_BUTTON, self.on_cancel_button)
        self.exit_button.Bind(wx.EVT_BUTTON, self.on_exit_button)

    def update_status(self, text, progress=None, cancel_flag=None, num_files=None):
        if self.output_text:
            self.output_text.AppendText(text)

            self.total_files_text.SetLabel(f"Total submissions to be exported: {num_files}")
            
            if progress is not None:
                percent = f"{progress:.2f}"
                progress_int = int(progress)
                self.gauge.SetValue(progress_int)
                self.SetTitle(f"Exporting submissions ({percent}%)")
                
            if cancel_flag and cancel_flag[0]:
                self.update_status("\nExport process canceled by the user.\n")
                self.cancel_button.Enable(False)
                self.exit_button.Enable(True)
                return

    def on_cancel_button(self, event):
        self.GetParent().cancel_flag[0] = True
        self.cancel_button.Enable(False)
        if "Export process canceled by the user." in self.output_text.GetValue():
                    self.exit_button.Enable(True)

    def on_exit_button(self, event):
        self.Close()

app = wx.App()
ExportSubmissionsToPdf()
app.MainLoop()