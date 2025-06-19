import os
import time
import requests
import pandas as pd
from PIL import Image
from io import BytesIO
from flask import Flask, render_template, request, send_file, jsonify
import pytesseract as pyt
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import NoAlertPresentException, InvalidSessionIdException
import tempfile
import threading
from concurrent.futures import ThreadPoolExecutor
import signal
import sys

# Global variable to track processing
processing_status = {"running": False, "progress": 0, "total": 0, "file_path": None}

def signal_handler(sig, frame):
    """Handle shutdown gracefully"""
    print('Shutting down gracefully...')
    if 'driver' in globals():
        try:
            driver.quit()
        except:
            pass
    sys.exit(0)

signal.signal(signal.SIGTERM, signal_handler)
signal.signal(signal.SIGINT, signal_handler)

def writeCSV(enroll, name, *args, sgpa, cgpa, remark, filename):
    gradesString = [str(a) + "," for a in args]
    information = [enroll, ",", name, ","
                   ] + gradesString + [sgpa, ",", cgpa, ",", remark, "\n"]
    with open(filename, 'a') as f:
        f.writelines(information)

def makeXslx(filename):
    csvFile = f'{filename}.csv'
    if not os.path.exists(csvFile):
        return None
    df = pd.read_csv(csvFile)
    excelFile = f'{filename}.xlsx'
    df.index += 1
    df.to_excel(excelFile)
    return excelFile

def downloadImage(url, name):
    try:
        content = requests.get(url, timeout=10).content
        file = BytesIO(content)
        image = Image.open(file)
        temp_dir = tempfile.gettempdir()
        path = os.path.join(temp_dir, name)
        with open(path, "wb") as f:
            image.save(f, "JPEG")
        return path
    except Exception as e:
        print(f"Error downloading image: {e}")
        return None

def readFromImage(captchaImage: str) -> str:
    try:
        # Tesseract path configuration
        if os.name == 'nt':
            pyt.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        else:
            pyt.pytesseract.tesseract_cmd = '/usr/bin/tesseract'
        
        image = Image.open(captchaImage)
        captcha = pyt.image_to_string(image)
        captcha = captcha.upper().replace(" ", "")
        os.remove(captchaImage)
        return captcha
    except Exception as e:
        print(f"Error reading captcha: {e}")
        if os.path.exists(captchaImage):
            os.remove(captchaImage)
        return ""

def create_chrome_driver():
    """Create Chrome driver with optimized options for Railway"""
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--disable-extensions')
    options.add_argument('--disable-plugins')
    options.add_argument('--disable-images')
    options.add_argument('--disable-javascript')
    options.add_argument('--window-size=1024,768')
    options.add_argument('--memory-pressure-off')
    options.add_argument('--max_old_space_size=4096')
    options.add_argument('--single-process')
    options.add_argument('--disable-background-timer-throttling')
    options.add_argument('--disable-renderer-backgrounding')
    options.add_argument('--disable-backgrounding-occluded-windows')
    
    try:
        driver = webdriver.Chrome(options=options)
        driver.implicitly_wait(2)
        return driver
    except Exception as e:
        print(f"Error creating Chrome driver: {e}")
        return None

def process_single_enrollment(enroll, driver, sem, first_row_data, filepath):
    """Process a single enrollment number"""
    try:
        print(f"Processing ==> {enroll}")
        
        # Get captcha
        img_element = driver.find_element(By.XPATH, '//img[contains(@src, "CaptchaImage.axd")]')
        img_src = img_element.get_attribute("src")
        url = f'http://result.rgpv.ac.in/result/{img_src.split("Result/")[-1]}'
        
        captcha_path = downloadImage(url, f"captcha_{enroll}.jpg")
        if not captcha_path:
            return False, "Failed to download captcha"
        
        captcha = readFromImage(captcha_path)
        if not captcha:
            return False, "Failed to read captcha"
        
        # Fill form
        Select(driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_drpSemester"]')).select_by_value(str(sem))
        driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_TextBox1"]').send_keys(captcha)
        time.sleep(0.5)
        driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_txtrollno"]').send_keys(enroll)
        time.sleep(1)
        driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnviewresult"]').send_keys(Keys.ENTER)
        
        time.sleep(2)
        
        # Handle alert
        try:
            alert = Alert(driver)
            alerttext = alert.text
            alert.accept()
            if "Result" in alerttext:
                driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnReset"]').send_keys(Keys.ENTER)
                return False, "Result not found"
        except NoAlertPresentException:
            pass
        
        # Check if result found
        if "Total Credit" in driver.page_source:
            # Write header if first row
            if first_row_data["is_first"]:
                details = []
                rows = driver.find_elements(By.CSS_SELECTOR, "table.gridtable tbody tr")
                for row in rows:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if len(cells) >= 4 and '[T]' in cells[0].text:
                        details.append(cells[0].text.strip('- [T]'))
                first_row_data["is_first"] = False
                writeCSV("Enrollment No.", "Name", *details, sgpa="SGPA", cgpa="CGPA", remark="REMARK", filename=filepath)
            
            # Extract data
            name = driver.find_element("id", "ctl00_ContentPlaceHolder1_lblNameGrading").text
            grades = []
            rows = driver.find_elements(By.CSS_SELECTOR, "table.gridtable tbody tr")
            for row in rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) >= 4 and '[T]' in cells[0].text:
                    grades.append(cells[3].text.strip())
            
            sgpa = driver.find_element("id", "ctl00_ContentPlaceHolder1_lblSGPA").text
            cgpa = driver.find_element("id", "ctl00_ContentPlaceHolder1_lblcgpa").text
            result = driver.find_element("id", "ctl00_ContentPlaceHolder1_lblResultNewGrading").text
            
            result = result.replace(",", " ")
            name = name.replace("\n", " ")
            writeCSV(enroll, name, *grades, sgpa=sgpa, cgpa=cgpa, remark=result, filename=filepath)
            
            driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnReset"]').send_keys(Keys.ENTER)
            return True, "Success"
        else:
            driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_TextBox1"]').clear()
            driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_txtrollno"]').clear()
            return False, "Wrong captcha"
            
    except Exception as e:
        print(f"Error processing {enroll}: {e}")
        try:
            driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_TextBox1"]').clear()
            driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_txtrollno"]').clear()
        except:
            pass
        return False, str(e)

def resultFound(start: int, end: int, branch: str, year: str, sem: int):
    """Main function to scrape results"""
    global processing_status
    
    if branch not in ["CS", "IT", "ME", "AI", "DS", "EC", "EX"]:
        return None, ["Wrong Branch Entered"]

    processing_status["running"] = True
    processing_status["progress"] = 0
    processing_status["total"] = end - start + 1
    
    noResult = []
    filename = f'{branch}_sem{sem}_result.csv'
    temp_dir = tempfile.gettempdir()
    filepath = os.path.join(temp_dir, filename)
    
    # Remove existing file
    if os.path.exists(filepath):
        os.remove(filepath)
    
    driver = create_chrome_driver()
    if not driver:
        processing_status["running"] = False
        return None, ["Failed to create browser driver"]
    
    try:
        driver.get("http://result.rgpv.ac.in/Result/ProgramSelect.aspx")
        driver.find_element(By.ID, "radlstProgram_1").click()
        
        first_row_data = {"is_first": True}
        current = start
        
        while current <= end and processing_status["running"]:
            # Format enrollment number
            if current < 10:
                num = "00" + str(current)
            elif current < 100:
                num = "0" + str(current)
            else:
                num = str(current)
            
            enroll = f"0105{branch}{year}1{num}"
            
            # Process enrollment with retry logic
            max_retries = 3
            for attempt in range(max_retries):
                success, message = process_single_enrollment(enroll, driver, sem, first_row_data, filepath)
                
                if success:
                    print(f"✓ {enroll} - Success")
                    break
                elif "Result not found" in message:
                    noResult.append(enroll)
                    print(f"✗ {enroll} - Not found")
                    break
                elif attempt < max_retries - 1:
                    print(f"⚠ {enroll} - Retry {attempt + 1}: {message}")
                    time.sleep(1)
                else:
                    print(f"✗ {enroll} - Failed after retries: {message}")
                    noResult.append(enroll)
            
            current += 1
            processing_status["progress"] = current - start
            
            # Memory management - restart driver every 10 enrollments
            if (current - start) % 10 == 0:
                try:
                    driver.quit()
                    time.sleep(2)
                    driver = create_chrome_driver()
                    if driver:
                        driver.get("http://result.rgpv.ac.in/Result/ProgramSelect.aspx")
                        driver.find_element(By.ID, "radlstProgram_1").click()
                except Exception as e:
                    print(f"Error restarting driver: {e}")
                    break
        
        processing_status["running"] = False
        
        if os.path.exists(filepath):
            excel_file = makeXslx(filepath.split(".")[0])
            processing_status["file_path"] = excel_file
        else:
            excel_file = None
            
        return excel_file, noResult
        
    except Exception as e:
        processing_status["running"] = False
        print(f"Error in resultFound: {e}")
        return None, [str(e)]
    finally:
        try:
            driver.quit()
        except:
            pass

app = Flask(__name__)

@app.route('/')
def form():
    return render_template('form.html')

@app.route('/submit', methods=['POST'])
def submit():
    """Start the scraping process"""
    global processing_status
    
    if processing_status["running"]:
        return jsonify({"error": "Another process is already running"}), 400
    
    try:
        branch = request.form['branch'].upper()
        year = request.form['year']
        sem = int(request.form['sem'])
        start = int(request.form['start'])
        end = int(request.form['end'])
        
        # Limit batch size to prevent memory issues
        if end - start > 50:
            return jsonify({"error": "Please limit batch size to 50 enrollments at a time"}), 400
        
        # Start processing in background thread
        def process():
            resultFound(start, end, branch, year, sem)
        
        thread = threading.Thread(target=process)
        thread.daemon = True
        thread.start()
        
        return jsonify({"message": "Processing started", "status": "started"})
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/status')
def status():
    """Check processing status"""
    return jsonify(processing_status)

@app.route('/download')
def download():
    """Download the result file"""
    if processing_status["file_path"] and os.path.exists(processing_status["file_path"]):
        return send_file(processing_status["file_path"], 
                        as_attachment=True, 
                        download_name=f'rgpv_results.xlsx',
                        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    else:
        return jsonify({"error": "No file available"}), 404

@app.route('/health')
def health():
    """Health check endpoint"""
    return jsonify({"status": "healthy", "timestamp": time.time()})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False, threaded=True)
