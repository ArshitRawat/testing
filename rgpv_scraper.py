import os
import time
import requests
import pandas as pd
from PIL import Image
from io import BytesIO
from flask import Flask, render_template, request, send_file, jsonify
import base64
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import NoAlertPresentException, InvalidSessionIdException
import tempfile
import gc
import threading
import queue
import json


class MemoryOptimizedScraper:
    def __init__(self, ocr_api_key):
        self.driver = None
        self.temp_files = []
        self.ocr_api_key = ocr_api_key
        
    def get_chrome_options(self):
        """Optimized Chrome options for minimal memory usage"""
        options = Options()
        
        # Essential headless options
        options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        
        # Memory optimization options
        options.add_argument('--disable-gpu')
        options.add_argument('--disable-software-rasterizer')
        options.add_argument('--disable-background-timer-throttling')
        options.add_argument('--disable-backgrounding-occluded-windows')
        options.add_argument('--disable-renderer-backgrounding')
        options.add_argument('--disable-features=TranslateUI')
        options.add_argument('--disable-ipc-flooding-protection')
        options.add_argument('--disable-default-apps')
        options.add_argument('--disable-extensions')
        options.add_argument('--disable-plugins')
        options.add_argument('--disable-images')  # Don't load images except captcha
        options.add_argument('--disable-javascript')  # Minimal JS needed
        
        # Reduce memory footprint
        options.add_argument('--memory-pressure-off')
        options.add_argument('--max_old_space_size=512')
        options.add_argument('--window-size=800,600')  # Smaller window
        
        # Disable unnecessary features
        options.add_argument('--disable-web-security')
        options.add_argument('--disable-features=VizDisplayCompositor')
        
        return options

    def cleanup_temp_files(self):
        """Clean up temporary files"""
        for file_path in self.temp_files:
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
            except:
                pass
        self.temp_files.clear()

    def download_image_optimized(self, url, name):
        """Download and optimize captcha image"""
        try:
            response = requests.get(url, timeout=10, stream=True)
            response.raise_for_status()
            
            # Process image in memory first
            image = Image.open(BytesIO(response.content))
            
            # Optimize image for OCR
            image = image.convert('L')  # Convert to grayscale
            image = image.point(lambda x: 0 if x < 128 else 255, '1')  # Binary threshold
            
            # Save to temp file
            temp_dir = tempfile.gettempdir()
            path = os.path.join(temp_dir, f"{name}_{int(time.time())}.jpg")
            image.save(path, "JPEG", optimize=True, quality=85)
            
            self.temp_files.append(path)
            return path
        except Exception as e:
            print(f"Error downloading image: {e}")
            return None

    def read_captcha_with_ocrspace(self, captcha_image: str) -> str:
        """Use OCR.space API for CAPTCHA reading"""
        try:
            # Read image and convert to base64
            with open(captcha_image, 'rb') as f:
                image_data = f.read()
            
            # Convert to base64
            base64_image = base64.b64encode(image_data).decode('utf-8')
            
            # OCR.space API endpoint
            url = 'https://api.ocr.space/parse/image'
            
            payload = {
                'apikey': self.ocr_api_key,
                'base64Image': f'data:image/jpeg;base64,{base64_image}',
                'language': 'eng',
                'isOverlayRequired': False,
                'detectOrientation': False,
                'isCreateSearchablePdf': False,
                'isSearchablePdfHideTextLayer': False,
                'scale': True,
                'isTable': False,
                'OCREngine': 2  # Use engine 2 for better accuracy
            }
            
            # Make API request with timeout
            response = requests.post(url, data=payload, timeout=10)
            
            if response.status_code == 200:
                result = response.json()
                if result.get('IsErroredOnProcessing', True) == False:
                    parsed_text = result.get('ParsedResults', [{}])[0].get('ParsedText', '')
                    captcha = parsed_text.upper().replace(" ", "").replace("\n", "").replace("\r", "")
                    
                    # Clean up image file immediately
                    try:
                        os.remove(captcha_image)
                        if captcha_image in self.temp_files:
                            self.temp_files.remove(captcha_image)
                    except:
                        pass
                    
                    print(f"OCR.space result: '{captcha}'")
                    return captcha
                else:
                    print(f"OCR.space error: {result.get('ErrorMessage', 'Unknown error')}")
            else:
                print(f"OCR.space API error: {response.status_code}")
                
        except Exception as e:
            print(f"OCR.space API Error: {e}")
        
        # Fallback: Clean up file and return empty
        try:
            os.remove(captcha_image)
            if captcha_image in self.temp_files:
                self.temp_files.remove(captcha_image)
        except:
            pass
        
        return ""

    def write_to_csv_batch(self, data_batch, filename):
        """Write data in batches to reduce memory usage"""
        try:
            df = pd.DataFrame(data_batch)
            if not os.path.exists(filename):
                df.to_csv(filename, index=False)
            else:
                df.to_csv(filename, mode='a', header=False, index=False)
            return True
        except Exception as e:
            print(f"Error writing CSV: {e}")
            return False

    def scrape_batch(self, start: int, end: int, branch: str, year: str, sem: int, batch_size: int = 5):
        """Scrape in small batches to manage memory"""
        if branch not in ["CS", "IT", "ME", "AI", "DS", "EC", "EX"]:
            return None, ["Invalid branch"]

        no_result = []
        all_data = []
        current_batch = []
        first_row = True
        headers = None
        
        # Create temporary CSV file
        temp_dir = tempfile.gettempdir()
        csv_filename = os.path.join(temp_dir, f'{branch}_sem{sem}_result_{int(time.time())}.csv')
        
        try:
            # Initialize driver
            self.driver = webdriver.Chrome(options=self.get_chrome_options())
            self.driver.implicitly_wait(2)
            self.driver.get("http://result.rgpv.ac.in/Result/ProgramSelect.aspx")
            self.driver.find_element(By.ID, "radlstProgram_1").click()

            current = start
            while current <= end:
                # Process in batches
                batch_end = min(current + batch_size - 1, end)
                print(f"Processing batch: {current} to {batch_end}")
                
                for num in range(current, batch_end + 1):
                    try:
                        result_data = self.process_single_enrollment(num, branch, year, sem)
                        
                        if result_data:
                            if first_row and result_data.get('is_valid_result'):
                                headers = result_data['headers']
                                # Write headers
                                header_row = ['Enrollment No.', 'Name'] + headers + ['SGPA', 'CGPA', 'REMARK']
                                pd.DataFrame([header_row]).to_csv(csv_filename, index=False, header=False)
                                first_row = False
                            
                            if result_data.get('is_valid_result'):
                                row_data = ([result_data['enrollment'], result_data['name']] + 
                                          result_data['grades'] + 
                                          [result_data['sgpa'], result_data['cgpa'], result_data['remark']])
                                current_batch.append(row_data)
                            else:
                                no_result.append(result_data['enrollment'])
                        else:
                            no_result.append(f"0105{branch}{year}1{num:03d}")
                            
                    except Exception as e:
                        print(f"Error processing enrollment {num}: {e}")
                        no_result.append(f"0105{branch}{year}1{num:03d}")
                
                # Write batch to file and clear memory
                if current_batch:
                    if headers:  # Only write if we have headers
                        df_batch = pd.DataFrame(current_batch)
                        df_batch.to_csv(csv_filename, mode='a', header=False, index=False)
                    current_batch.clear()
                
                # Force garbage collection after each batch
                gc.collect()
                current = batch_end + 1
                
                # Small pause to prevent overwhelming the server
                time.sleep(1)
            
            # Convert to Excel
            if os.path.exists(csv_filename) and os.path.getsize(csv_filename) > 0:
                excel_filename = csv_filename.replace('.csv', '.xlsx')
                df = pd.read_csv(csv_filename)
                df.index += 1
                df.to_excel(excel_filename, index=True)
                
                # Clean up CSV
                try:
                    os.remove(csv_filename)
                except:
                    pass
                    
                return excel_filename, no_result
            else:
                return None, no_result
                
        except Exception as e:
            print(f"Batch processing error: {e}")
            return None, no_result
        finally:
            self.cleanup_driver()
            self.cleanup_temp_files()

    def process_single_enrollment(self, num: int, branch: str, year: str, sem: int):
        """Process a single enrollment number"""
        try:
            enroll = f"0105{branch}{year}1{num:03d}"
            print(f"Processing: {enroll}")

            # Get captcha
            img_element = self.driver.find_element(By.XPATH, '//img[contains(@src, "CaptchaImage.axd")]')
            img_src = img_element.get_attribute("src")
            url = f'http://result.rgpv.ac.in/result/{img_src.split("Result/")[-1]}'

            captcha_path = self.download_image_optimized(url, "captcha")
            if not captcha_path:
                return None
                
            # Use OCR.space API
            captcha = self.read_captcha_with_ocrspace(captcha_path)
                
            if not captcha:
                print("Failed to read CAPTCHA")
                return None

            # Fill form
            Select(self.driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_drpSemester"]')).select_by_value(str(sem))
            self.driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_TextBox1"]').send_keys(captcha)
            time.sleep(0.5)
            self.driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_txtrollno"]').send_keys(enroll)
            time.sleep(1)
            self.driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnviewresult"]').send_keys(Keys.ENTER)

            time.sleep(2)
            
            # Handle alert
            try:
                alert = Alert(self.driver)
                alert_text = alert.text
                alert.accept()
                if "Result" in alert_text:
                    self.reset_form()
                    return {'enrollment': enroll, 'is_valid_result': False}
            except NoAlertPresentException:
                pass

            # Check if result found
            if "Total Credit" in self.driver.page_source:
                # Extract data
                name = self.driver.find_element("id", "ctl00_ContentPlaceHolder1_lblNameGrading").text
                name = name.replace("\n", " ").replace(",", " ")
                
                # Get subject headers and grades
                headers = []
                grades = []
                rows = self.driver.find_elements(By.CSS_SELECTOR, "table.gridtable tbody tr")
                for row in rows:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if len(cells) >= 4 and '[T]' in cells[0].text:
                        headers.append(cells[0].text.strip('- [T]'))
                        grades.append(cells[3].text.strip())
                
                sgpa = self.driver.find_element("id", "ctl00_ContentPlaceHolder1_lblSGPA").text
                cgpa = self.driver.find_element("id", "ctl00_ContentPlaceHolder1_lblcgpa").text
                result = self.driver.find_element("id", "ctl00_ContentPlaceHolder1_lblResultNewGrading").text.replace(",", " ")

                self.reset_form()
                
                return {
                    'enrollment': enroll,
                    'name': name,
                    'headers': headers,
                    'grades': grades,
                    'sgpa': sgpa,
                    'cgpa': cgpa,
                    'remark': result,
                    'is_valid_result': True
                }
            else:
                self.reset_form()
                return {'enrollment': enroll, 'is_valid_result': False}
                
        except Exception as e:
            print(f"Error processing {enroll}: {e}")
            try:
                self.reset_form()
            except:
                pass
            return None

    def reset_form(self):
        """Reset the form"""
        try:
            self.driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnReset"]').send_keys(Keys.ENTER)
            time.sleep(0.5)
        except:
            pass
            
    def cleanup_driver(self):
        """Clean up driver resources"""
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
            self.driver = None


app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Global progress tracking
progress_data = {}

# Get OCR API key from environment variable
OCR_API_KEY = os.environ.get('OCR_API_KEY')

if not OCR_API_KEY:
    print("WARNING: OCR_API_KEY environment variable not set!")

@app.route('/')
def form():
    return render_template('form.html')

@app.route('/submit', methods=['POST'])
def submit():
    try:
        # Check if API key is configured
        if not OCR_API_KEY:
            return f"""
            <h2>⚠️ Service Configuration Error</h2>
            <p>OCR API key is not configured on the server. Please contact the administrator.</p>
            <a href="/">Go Back</a>
            """
            
        branch = request.form['branch'].upper()
        year = request.form['year']
        sem = int(request.form['sem'])
        start = int(request.form['start'])
        end = int(request.form['end'])
        
        # Validate range to prevent memory issues
        if end - start > 100:
            return f"""
            <h2>⚠️ Range too large</h2>
            <p>Please limit your range to maximum 100 enrollments to prevent timeout issues.</p>
            <p>Current range: {end - start + 1} enrollments</p>
            <a href="/">Go Back</a>
            """

        # Create scraper instance with the server's API key
        scraper = MemoryOptimizedScraper(OCR_API_KEY)
        
        # Process in batches
        batch_size = min(15, end - start + 1)
        excel_file, no_result = scraper.scrape_batch(start, end, branch, year, sem, batch_size)
        
        if excel_file and os.path.exists(excel_file):
            def cleanup_file():
                time.sleep(60)  # Keep file for 1 minute
                try:
                    os.remove(excel_file)
                except:
                    pass
            
            # Schedule cleanup
            threading.Thread(target=cleanup_file, daemon=True).start()
            
            return send_file(excel_file, 
                           as_attachment=True, 
                           download_name=f'{branch}_sem{sem}_result.xlsx',
                           mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        else:
            return f"""
            <h2>ℹ️ No results found or error occurred</h2>
            <p>Enrollments not found: {len(no_result)}</p>
            <p>Please check your parameters and try again.</p>
            <a href="/">Go Back</a>
            """
            
    except Exception as e:
        return f"""
        <h2>❌ Error: {str(e)}</h2>
        <p>Please try with a smaller range or check your parameters.</p>
        <a href="/">Go Back</a>
        """

@app.route('/health')
def health():
    """Health check endpoint"""
    api_status = "configured" if OCR_API_KEY else "missing"
    return jsonify({
        'status': 'healthy', 
        'memory': f'{os.getpid()}',
        'ocr_api_status': api_status
    })

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False, threaded=True)
