

rgpv_scraper.py
231 lines

py

this is my project file, how do i make it run on a webserver like rsilway

Looking at your Python Flask application for scraping RGPV results, I can help you deploy it to Railway (or similar web hosting platforms). Here are the key changes and steps needed:


Modified rgpv_scraper.py for Railway Deployment
Code 
Now, let me create the necessary configuration files for Railway deployment:


requirements.txt
Code 

Dockerfile
Code 

templates/form.html
Code 

Procfile
Code 
Steps to Deploy on Railway:
1. Prepare Your Project Structure
Create this folder structure:

your-project/
├── rgpv_scraper.py (updated version from above)
├── requirements.txt
├── Dockerfile
├── Procfile
└── templates/
    └── form.html
2. Deploy to Railway
Option A: Using GitHub (Recommended)

Create a new GitHub repository
Upload all the files above to your repository
Go to Railway.app
Sign up/Login with GitHub
Click "New Project" → "Deploy from GitHub repo"
Select your repository
Option B: Using Railway CLI

Install Railway CLI: npm install -g @railway/cli
Login: railway login
In your project folder: railway init
Deploy: railway up
3. Key Changes Made for Web Server Deployment
Headless Chrome: Added Chrome options for server environment
File Handling: Using temporary directories instead of hardcoded paths
Port Configuration: Using environment variable for port (Railway requirement)
Error Handling: Better error handling for web deployment
File Download: Added file download functionality for results
Dependencies: Added all required packages in requirements.txt
4. Important Notes
Tesseract OCR: The Dockerfile installs Tesseract for CAPTCHA reading
Chrome & ChromeDriver: Automatically installed in the container
File Storage: Results are temporarily stored and downloaded immediately
Performance: May be slower on free tier due to resource limitations
5. Environment Variables (if needed)
You can set these in Railway dashboard:

PORT (automatically set by Railway)
Any other custom configurations
6. Potential Issues & Solutions
Memory Limits: Free tiers have memory restrictions; consider optimizing batch sizes
Timeout: Long scraping sessions might timeout; consider implementing progress tracking
CAPTCHA Recognition: OCR accuracy may vary; you might need to improve the image processing
The application will be accessible at your Railway-provided URL (something like https://your-app-name.up.railway.app) once deployed successfully.





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
    if branch not in ["CS", "IT", "ME", "AI", "DS", "EC", "EX"]:
        print("Wrong Branch Entered")
        return None, []

    noResult = []
    
    # Configure Chrome options for headless mode (required for Railway)
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--window-size=1920,1080')
    
    firstRow = True
    filename = f'{branch}_sem{sem}_result.csv'
    
    # Use temporary directory
    temp_dir = tempfile.gettempdir()
    filepath = os.path.join(temp_dir, filename)
    
    try:
        driver = webdriver.Chrome(options=options)
        driver.implicitly_wait(0.5)
        driver.get("http://result.rgpv.ac.in/Result/ProgramSelect.aspx")
        driver.find_element(By.ID, "radlstProgram_1").click()

        while start <= end:
            if (start < 10):
                num = "00" + str(start)
            elif (start < 100):
                num = "0" + str(start)
            else:
                num = str(start)

            enroll = f"0105{branch}{year}1{num}"
            print(f"Currently compiling ==> {enroll}")

            img_element = driver.find_element(
                By.XPATH, '//img[contains(@src, "CaptchaImage.axd")]')
            img_src = img_element.get_attribute("src")
            url = f'http://result.rgpv.ac.in/result/{img_src.split("Result/")[-1]}'

            captcha_path = downloadImage(url, "captcha.jpg")
            captcha = readFromImage(captcha_path)
            captcha = captcha.replace(" ", "")

            Select(
                driver.find_element(
                    By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_drpSemester"]')
            ).select_by_value(str(sem))
            driver.find_element(
                By.XPATH,
                '//*[@id="ctl00_ContentPlaceHolder1_TextBox1"]').send_keys(captcha)
            time.sleep(1)
            driver.find_element(
                By.XPATH,
                '//*[@id="ctl00_ContentPlaceHolder1_txtrollno"]').send_keys(enroll)
            time.sleep(2)
            driver.find_element(
                By.XPATH,
                '//*[@id="ctl00_ContentPlaceHolder1_btnviewresult"]').send_keys(
                    Keys.ENTER)

            time.sleep(2)
            alert = Alert(driver)
            alerttext = ""
            try:
                alerttext = alert.text
                alert.accept()
            except NoAlertPresentException:
                pass
            except InvalidSessionIdException:
                pass

            if "Total Credit" in driver.page_source:
                if (firstRow):
                    details = []
                    rows = driver.find_elements(By.CSS_SELECTOR,
                                                "table.gridtable tbody tr")
                    for row in rows:
                        cells = row.find_elements(By.TAG_NAME, "td")
                        if len(cells) >= 4 and '[T]' in cells[0].text:
                            details.append(cells[0].text.strip('- [T]'))
                    firstRow = False
                    writeCSV("Enrollment No.",
                             "Name",
                             *details,
                             sgpa="SGPA",
                             cgpa="CGPA",
                             remark="REMARK",
                             filename=filepath)
                name = driver.find_element(
                    "id", "ctl00_ContentPlaceHolder1_lblNameGrading").text
                grades = []
                rows = driver.find_elements(By.CSS_SELECTOR,
                                            "table.gridtable tbody tr")
                for row in rows:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if len(cells) >= 4 and '[T]' in cells[0].text:
                        grades.append(cells[3].text.strip())
                sgpa = driver.find_element(
                    "id", "ctl00_ContentPlaceHolder1_lblSGPA").text
                cgpa = driver.find_element(
                    "id", "ctl00_ContentPlaceHolder1_lblcgpa").text
                result = driver.find_element(
                    "id", "ctl00_ContentPlaceHolder1_lblResultNewGrading").text

                result = result.replace(",", " ")
                name = name.replace("\n", " ")
                writeCSV(enroll,
                         name,
                         *grades,
                         sgpa=sgpa,
                         cgpa=cgpa,
                         remark=result,
                         filename=filepath)
                print("Compilation Successful")

                driver.find_element(
                    By.XPATH,
                    '//*[@id="ctl00_ContentPlaceHolder1_btnReset"]').send_keys(
                        Keys.ENTER)

                start = start + 1
            else:
                if "Result" in alerttext:  # when enrollment number is not found
                    driver.find_element(
                        By.XPATH,
                        '//*[@id="ctl00_ContentPlaceHolder1_btnReset"]').send_keys(
                            Keys.ENTER)
                    start = start + 1
                    noResult.append(enroll)
                    print(f"Enrollment NO: {enroll} not found.")
                else:  # when captcha is wrong
                    driver.find_element(
                        By.XPATH,
                        '//*[@id="ctl00_ContentPlaceHolder1_TextBox1"]').clear()
                    driver.find_element(
                        By.XPATH,
                        '//*[@id="ctl00_ContentPlaceHolder1_txtrollno"]').clear()
                    print("Wrong Captcha Entered")
                    continue

        print(f'Enrollment Numbers not found ====> {noResult}')
        excel_file = makeXslx(filepath.split(".")[0])
        driver.quit()
        return excel_file, noResult
        
    except Exception as e:
        print(f"Error occurred: {e}")
        if 'driver' in locals():
            driver.quit()
        return None, noResult


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

