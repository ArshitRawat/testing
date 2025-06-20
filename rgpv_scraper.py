import os
import time
import requests
import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import NoAlertPresentException, InvalidSessionIdException
import threading
import zipfile
from datetime import datetime

# Create results directory
os.makedirs('results', exist_ok=True)

def writeCSV(enroll, name, *args, sgpa, cgpa, remark, filename):
    """Write student data to CSV file"""
    gradesString = [str(a) + "," for a in args]
    information = [enroll, ",", name, ","] + gradesString + [sgpa, ",", cgpa, ",", remark, "\n"]
    filepath = os.path.join('results', filename)
    with open(filepath, 'a') as f:
        f.writelines(information)

def makeXlsx(filename):
    """Convert CSV to Excel format"""
    csvFile = os.path.join('results', f'{filename}.csv')
    if os.path.exists(csvFile):
        df = pd.read_csv(csvFile)
        excelFile = os.path.join('results', f'{filename}.xlsx')
        df.index += 1
        df.to_excel(excelFile)
        return excelFile
    return None

def readFromImage(url: str) -> str:
    """Extract text from captcha image using OCR"""
    api_key = os.environ.get('OCR_API_KEY', 'K86969399988957')
    
    try:
        image_response = requests.get(url, timeout=10)
        if image_response.status_code != 200:
            print("Failed to fetch image from URL.")
            return ""

        ocr_response = requests.post(
            'https://api.ocr.space/parse/image',
            files={'filename': ('captcha.jpg', image_response.content)},
            data={'apikey': api_key, 'OCREngine': 2},
            timeout=15
        )

        result = ocr_response.json()
        if result.get('IsErroredOnProcessing'):
            return ""
        
        parsed = result.get('ParsedResults')
        if not parsed:
            return ""
        
        text = parsed[0].get('ParsedText', "")
        return text.upper().replace(" ", "").strip()
    
    except Exception as e:
        print(f"OCR Error: {e}")
        return ""

def get_chrome_driver():
    """Setup Chrome driver for Docker environment"""
    chrome_options = Options()
    
    # Docker-specific Chrome options
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-infobars")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--mute-audio")
    chrome_options.add_argument("--disable-background-timer-throttling")
    chrome_options.add_argument("--disable-backgrounding-occluded-windows")
    chrome_options.add_argument("--disable-renderer-backgrounding")
    chrome_options.add_argument("--disable-features=TranslateUI")
    chrome_options.add_argument("--disable-ipc-flooding-protection")
    chrome_options.add_argument("--single-process")
    chrome_options.add_argument("--disable-dev-tools")
    chrome_options.add_argument("--no-zygote")
    chrome_options.add_argument("--remote-debugging-port=9222")
    
    # Set Chrome binary location
    chrome_options.binary_location = "/usr/bin/google-chrome"
    
    # Create service
    service = Service("/usr/local/bin/chromedriver")
    
    try:
        driver = webdriver.Chrome(service=service, options=chrome_options)
        return driver
    except Exception as e:
        print(f"Error creating Chrome driver: {e}")
        raise

# Global variable to track scraping status
scraping_status = {"running": False, "progress": "", "error": None}

def resultFound(start: int, end: int, branch: str, year: str, sem: int):
    """Main scraping function"""
    global scraping_status
    
    if branch not in ["CS", "IT", "ME", "AI", "DS", "EC", "EX"]:
        scraping_status["error"] = "Wrong Branch Entered"
        return
    
    scraping_status["running"] = True
    scraping_status["error"] = None
    noResult = []
    
    try:
        driver = get_chrome_driver()
        firstRow = True
        filename = f'{branch}_sem{sem}_result_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv'
        driver.implicitly_wait(0.5)
        
        scraping_status["progress"] = "Initializing..."
        driver.get("http://result.rgpv.ac.in/Result/ProgramSelect.aspx")
        driver.find_element(By.ID, "radlstProgram_1").click()

        total_students = end - start + 1
        processed = 0

        while start <= end:
            if start < 10:
                num = "00" + str(start)
            elif start < 100:
                num = "0" + str(start)
            else:
                num = str(start)
            
            enroll = f"0105{branch}{year}1{num}"
            processed += 1
            scraping_status["progress"] = f"Processing {enroll} ({processed}/{total_students})"
            print(f"Currently compiling ==> {enroll}")

            try:
                img_element = driver.find_element(By.XPATH, '//img[contains(@src, "CaptchaImage.axd")]')
                img_src = img_element.get_attribute("src")
                url = f'http://result.rgpv.ac.in/result/{img_src.split("Result/")[-1]}'

                captcha = readFromImage(url)
                captcha = captcha.replace(" ", "")

                Select(driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_drpSemester"]')).select_by_value(str(sem))
                driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_TextBox1"]').send_keys(captcha)
                time.sleep(1)
                driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_txtrollno"]').send_keys(enroll)
                time.sleep(2)
                driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnviewresult"]').send_keys(Keys.ENTER)

                time.sleep(2)
                alert = Alert(driver)
                alerttext = ""
                try:
                    alerttext = alert.text
                    alert.accept()
                except NoAlertPresentException:
                    pass

                if "Total Credit" in driver.page_source:
                    if firstRow:
                        details = []
                        rows = driver.find_elements(By.CSS_SELECTOR, "table.gridtable tbody tr")
                        for row in rows:
                            cells = row.find_elements(By.TAG_NAME, "td")
                            if len(cells) >= 4 and '[T]' in cells[0].text:
                                details.append(cells[0].text.strip('- [T]'))
                        firstRow = False
                        writeCSV("Enrollment No.", "Name", *details, sgpa="SGPA", cgpa="CGPA", remark="REMARK", filename=filename)
                    
                    roll_nu = driver.find_element("id", 'ctl00_ContentPlaceHolder1_lblRollNoGrading').text
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
                    writeCSV(enroll, name, *grades, sgpa=sgpa, cgpa=cgpa, remark=result, filename=filename)
                    print("Compilation Successful")

                    driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnReset"]').send_keys(Keys.ENTER)
                    start += 1
                else:
                    if "Result" in alerttext:
                        driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnReset"]').send_keys(Keys.ENTER)
                        start += 1
                        noResult.append(enroll)
                        print(f"Enrollment NO: {enroll} not found.")
                    else:
                        driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_TextBox1"]').clear()
                        driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_txtrollno"]').clear()
                        print("Wrong Captcha Entered")
                        continue

            except Exception as e:
                print(f"Error processing {enroll}: {e}")
                start += 1
                continue

        print(f'Enrollment Numbers not found ====> {noResult}')
        
        # Create Excel file
        makeXlsx(filename.replace('.csv', ''))
        
        scraping_status["progress"] = f"Completed! Processed {total_students} students. File: {filename}"
        
    except Exception as e:
        scraping_status["error"] = str(e)
        print(f"Error during scraping: {e}")
    finally:
        if 'driver' in locals():
            driver.quit()
        scraping_status["running"] = False

# Flask App
app = Flask(__name__)

@app.route('/')
def form():
    return render_template('1st.html')

@app.route('/submit', methods=['POST'])
def submit():
    global scraping_status
    
    if scraping_status["running"]:
        return jsonify({"error": "Scraping is already in progress"}), 400
    
    try:
        branch = request.form['branch'].upper()
        year = request.form['year']
        sem = int(request.form['sem'])
        start = int(request.form['start'])
        end = int(request.form['end'])

        # Start scraping in background thread
        thread = threading.Thread(target=resultFound, args=(start, end, branch, year, sem))
        thread.start()

        return jsonify({"message": "Scraping started", "status": "success"})
    
    except Exception as e:
        return jsonify({"error": str(e)}), 400

@app.route('/status')
def status():
    return jsonify(scraping_status)

@app.route('/results')
def list_results():
    """List all result files"""
    files = []
    if os.path.exists('results'):
        for filename in os.listdir('results'):
            if filename.endswith(('.csv', '.xlsx')):
                filepath = os.path.join('results', filename)
                files.append({
                    'name': filename,
                    'size': os.path.getsize(filepath),
                    'modified': datetime.fromtimestamp(os.path.getmtime(filepath)).strftime('%Y-%m-%d %H:%M:%S')
                })
    return jsonify(files)

@app.route('/download/<filename>')
def download_file(filename):
    """Download result file"""
    filepath = os.path.join('results', filename)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    return "File not found", 404

@app.route('/health')
def health():
    return jsonify({'status': 'healthy', 'scraping': scraping_status["running"]})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
