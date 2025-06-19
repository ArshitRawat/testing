import os
import time
import requests
import pandas as pd
from PIL import Image
from io import BytesIO
from flask import Flask, render_template, request, send_file
import pytesseract as pyt
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoAlertPresentException, InvalidSessionIdException, TimeoutException
import tempfile
import logging

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def writeCSV(enroll, name, *args, sgpa, cgpa, remark, filename):
    gradesString = [str(a) + "," for a in args]
    information = [enroll, ",", name, ","
                   ] + gradesString + [sgpa, ",", cgpa, ",", remark, "\n"]
    with open(filename, 'a') as f:
        f.writelines(information)

def makeXslx(filename):
    csvFile = f'{filename}.csv'
    df = pd.read_csv(csvFile)
    excelFile = f'{filename}.xlsx'
    df.index += 1
    df.to_excel(excelFile)
    return excelFile

def downloadImage(url, name):
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        
        file = BytesIO(response.content)
        image = Image.open(file)
        
        temp_dir = tempfile.gettempdir()
        path = os.path.join(temp_dir, name)
        
        with open(path, "wb") as f:
            image.save(f, "JPEG")
        return path
    except Exception as e:
        logger.error(f"Error downloading image: {e}")
        return None

def readFromImage(captchaImage: str) -> str:
    try:
        # Set tesseract path for different environments
        if os.name == 'nt':  # Windows
            pyt.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        else:  # Linux/Unix (Railway uses this)
            # Try common paths
            possible_paths = ['/usr/bin/tesseract', '/usr/local/bin/tesseract', 'tesseract']
            for path in possible_paths:
                if os.path.exists(path) or path == 'tesseract':
                    pyt.pytesseract.tesseract_cmd = path
                    break
        
        image = Image.open(captchaImage)
        captcha = pyt.image_to_string(image, config='--psm 8 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789')
        captcha = captcha.upper().replace(" ", "").replace("\n", "")
        
        # Clean up the image file
        if os.path.exists(captchaImage):
            os.remove(captchaImage)
        
        return captcha
    except Exception as e:
        logger.error(f"Error reading captcha: {e}")
        if os.path.exists(captchaImage):
            os.remove(captchaImage)
        return ""

def create_driver():
    """Create a new Chrome driver instance with proper configuration"""
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--window-size=1920,1080')
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_argument('--user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36')
    options.add_argument('--disable-extensions')
    options.add_argument('--disable-plugins')
    options.add_argument('--disable-images')  # Speed up loading
    options.add_argument('--memory-pressure-off')
    
    # Set memory limits
    options.add_argument('--max_old_space_size=512')
    options.add_argument('--memory-pressure-off')
    
    try:
        driver = webdriver.Chrome(options=options)
        driver.set_page_load_timeout(30)
        driver.implicitly_wait(10)
        return driver
    except Exception as e:
        logger.error(f"Error creating driver: {e}")
        return None

def resultFound(start: int, end: int, branch: str, year: str, sem: int):
    if branch not in ["CS", "IT", "ME", "AI", "DS", "EC", "EX"]:
        logger.error("Wrong Branch Entered")
        return None, []

    noResult = []
    driver = None
    
    firstRow = True
    filename = f'{branch}_sem{sem}_result.csv'
    
    # Use temporary directory
    temp_dir = tempfile.gettempdir()
    filepath = os.path.join(temp_dir, filename)
    
    # Clean up any existing file
    if os.path.exists(filepath):
        os.remove(filepath)
    
    try:
        driver = create_driver()
        if not driver:
            logger.error("Failed to create driver")
            return None, []
        
        wait = WebDriverWait(driver, 20)
        
        logger.info("Loading initial page...")
        driver.get("http://result.rgpv.ac.in/Result/ProgramSelect.aspx")
        
        # Wait for page to load and click program selection
        program_radio = wait.until(EC.element_to_be_clickable((By.ID, "radlstProgram_1")))
        program_radio.click()
        
        consecutive_failures = 0
        max_consecutive_failures = 5
        
        while start <= end:
            try:
                # Check if we've had too many consecutive failures
                if consecutive_failures >= max_consecutive_failures:
                    logger.error(f"Too many consecutive failures ({consecutive_failures}). Stopping.")
                    break
                
                if (start < 10):
                    num = "00" + str(start)
                elif (start < 100):
                    num = "0" + str(start)
                else:
                    num = str(start)

                enroll = f"0105{branch}{year}1{num}"
                logger.info(f"Currently compiling ==> {enroll}")

                # Wait for captcha image to load
                try:
                    img_element = wait.until(EC.presence_of_element_located(
                        (By.XPATH, '//img[contains(@src, "CaptchaImage.axd")]')))
                    img_src = img_element.get_attribute("src")
                    url = f'http://result.rgpv.ac.in/result/{img_src.split("Result/")[-1]}'
                except TimeoutException:
                    logger.error("Captcha image not found, refreshing page...")
                    driver.refresh()
                    time.sleep(3)
                    continue

                captcha_path = downloadImage(url, f"captcha_{start}.jpg")
                if not captcha_path:
                    logger.error("Failed to download captcha image")
                    consecutive_failures += 1
                    continue
                
                captcha = readFromImage(captcha_path)
                if not captcha or len(captcha) < 4:
                    logger.error("Failed to read captcha or captcha too short")
                    consecutive_failures += 1
                    continue

                # Clear previous inputs
                try:
                    captcha_input = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_TextBox1"]')
                    captcha_input.clear()
                    
                    roll_input = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_txtrollno"]')
                    roll_input.clear()
                except:
                    pass

                # Select semester
                try:
                    semester_dropdown = Select(driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_drpSemester"]'))
                    semester_dropdown.select_by_value(str(sem))
                except Exception as e:
                    logger.error(f"Error selecting semester: {e}")

                # Enter captcha and enrollment number
                captcha_input = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_TextBox1"]')
                captcha_input.send_keys(captcha)
                
                time.sleep(1)
                
                roll_input = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_txtrollno"]')
                roll_input.send_keys(enroll)
                
                time.sleep(2)
                
                # Submit form
                submit_button = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnviewresult"]')
                submit_button.click()

                time.sleep(3)
                
                # Handle alert if present
                alert_text = ""
                try:
                    wait_short = WebDriverWait(driver, 5)
                    alert = wait_short.until(EC.alert_is_present())
                    alert_text = alert.text
                    alert.accept()
                except TimeoutException:
                    pass
                except Exception as e:
                    logger.error(f"Error handling alert: {e}")

                # Check if result is found
                if "Total Credit" in driver.page_source:
                    consecutive_failures = 0  # Reset failure counter on success
                    
                    if firstRow:
                        details = []
                        try:
                            rows = driver.find_elements(By.CSS_SELECTOR, "table.gridtable tbody tr")
                            for row in rows:
                                cells = row.find_elements(By.TAG_NAME, "td")
                                if len(cells) >= 4 and '[T]' in cells[0].text:
                                    details.append(cells[0].text.strip('- [T]'))
                            firstRow = False
                            writeCSV("Enrollment No.", "Name", *details, 
                                   sgpa="SGPA", cgpa="CGPA", remark="REMARK", filename=filepath)
                        except Exception as e:
                            logger.error(f"Error processing header row: {e}")
                    
                    # Extract student data
                    try:
                        name = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblNameGrading").text
                        grades = []
                        rows = driver.find_elements(By.CSS_SELECTOR, "table.gridtable tbody tr")
                        for row in rows:
                            cells = row.find_elements(By.TAG_NAME, "td")
                            if len(cells) >= 4 and '[T]' in cells[0].text:
                                grades.append(cells[3].text.strip())
                        
                        sgpa = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblSGPA").text
                        cgpa = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblcgpa").text
                        result = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblResultNewGrading").text

                        result = result.replace(",", " ")
                        name = name.replace("\n", " ").replace(",", " ")
                        
                        writeCSV(enroll, name, *grades, sgpa=sgpa, cgpa=cgpa, remark=result, filename=filepath)
                        logger.info("Compilation Successful")
                    except Exception as e:
                        logger.error(f"Error extracting student data: {e}")

                    # Reset form
                    try:
                        reset_button = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnReset"]')
                        reset_button.click()
                        time.sleep(2)
                    except Exception as e:
                        logger.error(f"Error resetting form: {e}")

                    start += 1
                    
                else:
                    if "Result" in alert_text or "not found" in alert_text.lower():
                        # Enrollment number not found
                        try:
                            reset_button = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnReset"]')
                            reset_button.click()
                            time.sleep(2)
                        except Exception as e:
                            logger.error(f"Error resetting form: {e}")
                        
                        start += 1
                        noResult.append(enroll)
                        logger.info(f"Enrollment NO: {enroll} not found.")
                        consecutive_failures = 0  # Don't count "not found" as failure
                    else:
                        # Wrong captcha or other error
                        logger.info("Wrong Captcha or other error, retrying...")
                        consecutive_failures += 1
                        time.sleep(2)
                        continue

                # Add delay between requests to avoid being blocked
                time.sleep(2)
                
            except Exception as e:
                logger.error(f"Error in iteration for {enroll if 'enroll' in locals() else start}: {e}")
                consecutive_failures += 1
                
                # Try to reset the form
                try:
                    reset_button = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnReset"]')
                    reset_button.click()
                    time.sleep(2)
                except:
                    pass
                
                # If too many consecutive failures, try refreshing the page
                if consecutive_failures >= 3:
                    logger.info("Multiple failures, refreshing page...")
                    try:
                        driver.refresh()
                        time.sleep(5)
                        program_radio = wait.until(EC.element_to_be_clickable((By.ID, "radlstProgram_1")))
                        program_radio.click()
                    except Exception as refresh_error:
                        logger.error(f"Error refreshing page: {refresh_error}")
                        break
                
                time.sleep(3)

        logger.info(f'Enrollment Numbers not found: {noResult}')
        
        if os.path.exists(filepath):
            excel_file = makeXslx(filepath.split(".")[0])
            return excel_file, noResult
        else:
            logger.error("No CSV file was created")
            return None, noResult
        
    except Exception as e:
        logger.error(f"Critical error occurred: {e}")
        return None, noResult
    finally:
        # Always clean up the driver
        if driver:
            try:
                driver.quit()
                logger.info("Driver closed successfully")
            except Exception as e:
                logger.error(f"Error closing driver: {e}")

app = Flask(__name__)

@app.route('/')
def form():
    return render_template('form.html')

@app.route('/submit', methods=['POST'])
def submit():
    try:
        branch = request.form['branch']
        year = request.form['year']
        sem = int(request.form['sem'])
        start = int(request.form['start'])
        end = int(request.form['end'])

        # Validate input ranges
        if end - start > 50:  # Limit batch size to prevent timeouts
            return f"""
            <h2>⚠️ Batch size too large</h2>
            <p>Please limit your range to 50 students or less to prevent timeouts.</p>
            <a href="/">Go Back</a>
            """

        logger.info(f"Starting scraping for {branch} semester {sem}, range {start}-{end}")
        excel_file, no_result = resultFound(start, end, branch.upper(), year, sem)
        
        if excel_file and os.path.exists(excel_file):
            return send_file(excel_file, 
                           as_attachment=True, 
                           download_name=f'{branch.upper()}_sem{sem}_result.xlsx',
                           mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        else:
            return f"""
            <h2>❌ Error occurred while fetching results</h2>
            <p>No results were found or an error occurred during processing.</p>
            <p>Missing enrollment numbers: {no_result}</p>
            <a href="/">Go Back</a>
            """
    except Exception as e:
        logger.error(f"Error in submit route: {e}")
        return f"""
        <h2>❌ Error: {str(e)}</h2>
        <a href="/">Go Back</a>
        """

if __name__ == '__main__':
    # Use environment variable for port (Railway requirement)
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
