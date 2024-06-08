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


def export_submissions_to_pdf():

    # Read configuration file
    config = configparser.ConfigParser()
    config.read('config.ini')

    # Get configuration values
    username = config.get('DEFAULT', 'username')
    password = config.get('DEFAULT', 'password')
    token = config.get('DEFAULT', 'token')
    start_date = config.get('DEFAULT', 'start_date')
    end_date = config.get('DEFAULT', 'end_date')
    form_id = config.get('DEFAULT', 'form_id')
    namevar = config.get('DEFAULT', 'namevar')
    status = config.get('DEFAULT', 'status')
    driver_path = config.get('DEFAULT', 'driver_path')
        
    # Set up Chrome options and driver
    chrome_options = Options()
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920x1080")
    chrome_options.add_argument("--log-level=3") # Hide extraneous logging output
    chrome_options.add_argument("--headless")
    driver_path = "C:/Users/Haytham.Alsoufi/Desktop/PDFF/chromedriver.exe"
    service = Service(executable_path=driver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)

    start_time = timeit.default_timer()

    try:
        # Log in to Kobo
        login_start_time = timeit.default_timer()
        driver.get("https://kobonew.ifrc.org/accounts/login/")
        wait = WebDriverWait(driver, 10)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "login")))
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
            wait.until(EC.presence_of_element_located((By.ID, "form-title")))

            driver.switch_to.window(driver.window_handles[-1])
            # Set the desired PDF name
            pdf_name = f"{results[i][namevar]}_{results[i]['_submission_time'][:10].replace('-', '')}_{results[i]['_validation_status']['label']}.pdf"
            pdf_path = os.path.join(os.getcwd(), pdf_name)

            # Calculate the elapsed time
            elapsed_time = timeit.default_timer() - start_time
            total_time += elapsed_time
            print(f"{i+1}. {pdf_name} [in {elapsed_time:.2f} seconds]")

            # Save the current page as PDF
            
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

export_submissions_to_pdf()
