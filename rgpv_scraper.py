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
from selenium.common.exceptions import NoAlertPresentException, InvalidSessionIdException
import tempfile


def writeCSV(enroll, name, *args, sgpa, cgpa, remark, filename):
    gradesString = [str(a) + "," for a in args]
    information = [enroll, ",", name, ","
                   ] + gradesString + [sgpa, ",", cgpa, ",", remark, "\n"]
    with open(filename, 'a') as f:
        f.writelines(information)
        f.close()


def makeXslx(filename):
    csvFile = f'{filename}.csv'
    df = pd.read_csv(csvFile)
    excelFile = f'{filename}.xlsx'
    df.index += 1
    df.to_excel(excelFile)
    return excelFile


def downloadImage(url, name):
    content = requests.get(url).content
    file = BytesIO(content)
    image = Image.open(file)
    # Use temporary directory instead of hardcoded path
    temp_dir = tempfile.gettempdir()
    path = os.path.join(temp_dir, name)
    with open(path, "wb") as f:
        image.save(f, "JPEG")
    return path

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


def readFromImage(captchaImage: str) -> str:
    # Set tesseract path for different environments
    if os.name == 'nt':  # Windows
        pyt.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    else:  # Linux/Unix (Railway uses this)
        pyt.pytesseract.tesseract_cmd = '/usr/bin/tesseract'
    
    image = Image.open(captchaImage)
    captcha = pyt.image_to_string(image)
    captcha = captcha.upper().replace(" ", "")
    os.remove(captchaImage)
    return captcha


def resultFound(start: int, end: int, branch: str, year: str, sem: int):
    """Main function to scrape results"""
    
    if branch not in ["CS", "IT", "ME", "AI", "DS", "EC", "EX"]:
        return None, ["Wrong Branch Entered"]
    
    noResult = []
    filename = f'{branch}_sem{sem}_result.csv'
    temp_dir = tempfile.gettempdir()
    filepath = os.path.join(temp_dir, filename)
    
    # Remove existing file
    if os.path.exists(filepath):
        os.remove(filepath)
    
    driver = create_chrome_driver()
    if not driver:
        return None, ["Failed to create browser driver"]
    
    try:
        driver.get("http://result.rgpv.ac.in/Result/ProgramSelect.aspx")
        driver.find_element(By.ID, "radlstProgram_1").click()
        
        first_row_data = {"is_first": True}
        current = start
        
        while current <= end:
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
        
        
        if os.path.exists(filepath):
            excel_file = makeXslx(filepath.split(".")[0])
        else:
            excel_file = None
            
        return excel_file, noResult
        
    except Exception as e:
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
    try:
        branch = request.form['branch']
        year = request.form['year']
        sem = int(request.form['sem'])
        start = int(request.form['start'])
        end = int(request.form['end'])

        excel_file, no_result = resultFound(start, end, branch.upper(), year, sem)
        
        if excel_file and os.path.exists(excel_file):
            return send_file(excel_file, 
                           as_attachment=True, 
                           download_name=f'{branch.upper()}_sem{sem}_result.xlsx',
                           mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        else:
            return f"""
            <h2>❌ Error occurred while fetching results</h2>
            <p>Please try again or check your parameters.</p>
            <a href="/">Go Back</a>
            """
    except Exception as e:
        return f"""
        <h2>❌ Error: {str(e)}</h2>
        <a href="/">Go Back</a>
        """


if __name__ == '__main__':
    # Use environment variable for port (Railway requirement)
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
