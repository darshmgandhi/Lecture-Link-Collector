from __future__ import print_function

import os
import io
import sys
import json
import calendar
import subprocess
import pkg_resources
from getpass import getpass
from traceback import print_exc, format_exc
from datetime import date, datetime, timedelta

required_modules = {'selenium', 'google-api-python-client', 'google-auth-httplib2', 'google-auth-oauthlib'}
installed_modules = {pkg.key for pkg in pkg_resources.working_set}
missing_modules = required_modules - installed_modules

if missing_modules:
    print('To run this program properly, some packages have to be installed.\nThose packages are:', *['  ' + i for i in missing_modules], sep = '\n')
    while True:
        install_choice = input('Do you want to proceed with the installation(y/n)? ').lower()
        if install_choice == 'y':
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', '--upgrade', *missing_modules])
            print('Running Lecture Links Collector...\n')
            break
        elif install_choice == 'n':
            print('This program cannot run without installing the listed packages. Quitting the program...')
            sys.exit()
        else:
            print('Please choose a valid option.')

from selenium import webdriver
from selenium.webdriver.common.by import By 
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, NoSuchWindowException, WebDriverException

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google.auth.exceptions import RefreshError

#Getting all the class links of a particular course on a given day
def get_class_link(course_name):
    course_link = []
    course_title = driver.find_elements_by_xpath("//div[contains(@title, '" + course_name + "')][contains(@class, 'rbc-event')]")
    course_title = [k.get_attribute("title") for k in course_title] 
    if len(course_title) > 0:
        class_date = driver.find_element_by_class_name('rbc-toolbar-label').text
        for i in range(len(course_title)):
            try:
                course_recording = driver.find_element_by_xpath("//div[@title = '" + course_title[i] + "']//a[@class = 'recordingbtn']").get_attribute("vurl").replace('preview', 'view')
            except NoSuchElementException:
                course_recording = 'Class Cancelled' 
            class_type = driver.find_element_by_xpath("//div[@title = '" + course_title[i] + "']//h6")
            course_link.append([class_type.text[-2]])
            course_link[-1].extend(class_date.split())
            course_link[-1].append(course_recording)
            course_link[-1][2] = datetime.strptime(course_link[-1][2], '%b').strftime('%B')
    return course_link        

def convert_to_ist(datetime_utc):
    return (datetime.fromisoformat(datetime_utc) + timedelta(hours = 5, minutes = 30)).strftime('%d %b %Y %I:%M:%S %p')

def save(data):
    emails = []
    SCOPES = ['https://www.googleapis.com/auth/drive.file']

    print('\nPlease select a method to save your lecture links:')
    if os.path.exists('token.json'):
        with open('token.json', 'r') as preferences:
            users = json.load(preferences)
        emails = users["emails"]
        for user in range(len(emails)):
            print(f'{user + 1}. Upload to Google Drive as a Spreadsheet ({emails[user]})')
    choice = int(input(f'{len(emails) + 1}. Upload to Google Drive as a Spreadsheet{" (Use another account)" if emails else ""}\n{len(emails) + 2}. Save locally as a CSV file\nChoose: '))

    if choice <= len(emails):
        creds = Credentials.from_authorized_user_info(users['tokens'][choice - 1], SCOPES)
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
            users['tokens'][choice - 1] = json.loads(creds.to_json())
            with open('token.json', 'w') as token:
                token.write(json.dumps(users))
        service = build('drive', 'v3', credentials = creds)    
    elif choice == 1 or choice == (len(emails) + 1):
        print('Opening authorization window...')
        client_credentials = {"installed":{"client_id":"292011036249-a6r9097a7q4216d0copqhdbkon3flsa5.apps.googleusercontent.com",
                       "project_id":"lecture-links-collector", "auth_uri":"https://accounts.google.com/o/oauth2/auth", 
                       "token_uri":"https://oauth2.googleapis.com/token", 
                       "auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs",
                       "client_secret":"qRsAEyh_9ddbzoLGMIM989_m", "redirect_uris":["urn:ietf:wg:oauth:2.0:oob","http://localhost"]}}
        flow = InstalledAppFlow.from_client_config(client_credentials, SCOPES)
        creds = flow.run_local_server(port=0, authorization_prompt_message = 'Please authorize Lecture Links Collector at this URL: {url}',
                                      success_message = 'Thank you for the authorization. Please close this window and return to the program.')
        print('Authorization Completed')
        service = build('drive', 'v3', credentials = creds)
        email_address = service.about().get(fields = 'user(emailAddress)').execute()['user']['emailAddress']
        if emails:
            users['emails'].append(email_address)
            users['tokens'].append(json.loads(creds.to_json()))
        else:
            users = {'tokens': [json.loads(creds.to_json())], 'emails': [email_address]}
        with open('token.json', 'w') as token:
            token.write(json.dumps(users))
    else:
        filename = input(f"Save As(Enter filename/path without extension): ").strip().replace('\\\\', '\\')
        with open(filename + '.csv', 'w') as f:
            f.write(data)
        print(f'File \'{filename}\' has been successfully locally saved as a CSV File.')
        return
    print('NOTE: You will now be asked to enter a desired name for your lecture links file. If you want to update an older file instead of creating a new file, just write the name of that old file.')
    filename = input("Save As(Enter filename without extension): ").strip()
    media = MediaIoBaseUpload(io.BytesIO(data.encode('utf-8')), mimetype = 'text/csv')
    files_on_drive = service.files().list(q = f"name = '{filename}'", 
                                        fields = 'files(id, name, webViewLink, createdTime, viewedByMeTime, owners(emailAddress))') \
                                        .execute().get('files')
    if files_on_drive:
        print('Some existing files with the same name have been found in your drive. Enter the corresponding number to the file description' ,
               'to update that particular file or create a new file.')
        for f in range(len(files_on_drive)):
            print((f'{f + 1}. NAME: {files_on_drive[f]["name"]}  OWNER: {files_on_drive[f]["owners"][0]["emailAddress"]}\n'
                   f'   CREATED ON: {convert_to_ist(files_on_drive[f]["createdTime"].rstrip("Z"))}  '
                   f'LAST VIEWED BY YOU ON: {convert_to_ist(files_on_drive[f]["viewedByMeTime"].rstrip("Z"))}\n'
                   f'   LINK: {files_on_drive[f]["webViewLink"]}'))
        choice = int(input(f'{len(files_on_drive) + 1}. Create new file\nChoose: '))
        if choice != (len(files_on_drive) + 1):
            print('Updating file...')
            result = service.files().update(fileId = files_on_drive[choice - 1]['id'], media_body = media,
                                            fields = 'id, name, createdTime, webViewLink').execute()
            print(f'UPDATE SUCCESSFUL\nFile \'{result["name"]}\' that was created on {convert_to_ist(result["createdTime"].rstrip("Z"))} has been successfully updated on your drive account. \
                    \nLink to the file: {result["webViewLink"]}')
            return
    print('Uploading file...')
    metadata = {'name': filename, 'mimeType': 'application/vnd.google-apps.spreadsheet'}
    result = service.files().create(body = metadata, media_body = media, 
                                    fields = 'id, name, webViewLink, createdTime').execute()  
    print(f'UPLOAD SUCCESSFUL\nFile \'{result["name"]}\' has been successfully created on your drive account on {convert_to_ist(result["createdTime"].rstrip("Z"))}. \
        \nLink to the file: {result["webViewLink"]}')  

def save_as_file(data):
    try:
        save(data)
    except (RefreshError, ValueError, json.JSONDecodeError) as e:
        os.remove('token.json')
        with open('errors.log', 'a') as error:
            error.write(format_exc())
        print('ATTENTION: Your token.json file that stored all your user data has gone corrupt.', 
              'Manually changing the contents of the file could be one of the reasons. The corrupt file has been deleted.',
              'All your previously saved logins are now erased. You will have to do the Sign In process again. Inconvenience is regretted.')
        save(data)
        return
    except:
        with open('errors.log', 'a') as error:
            error.write(format_exc())
        print('Unfortunately, an unknown error has occurred. Exiting the program...')

if __name__ == "__main__":
    print("Lecture Link Collector(for learn.niituniversity.in)")
    #Taking user credentials and course details as input
    user_data = {}
    email_id = None
    course_preferences = []
    date_preferences = []
    if os.path.exists('paths.json'):
        try:
            with open('paths.json') as user_file:
                user_data = json.load(user_file)
            email_id = user_data['email']
            course_preferences = user_data['courses']
            date_preferences = user_data['dates']
        except json.JSONDecodeError:
            os.remove('paths.json')
            print('ATTENTION: Your paths.json file that stored your data and the path to your chromedriver/geckodriver has gone corrupt.', 
                  'Manually changing the contents of the file could be one of the reasons. The corrupt file has been deleted.',
                  'The driver paths that you had entered as input are now erased. You will have to input them again. Inconvenience is regretted.')
        except KeyError:
            pass
    if not email_id or input(f'Continue with {email_id}@st.niituniversity.in(y/n)? ').strip().lower() != 'y':
        email_id = input("Email Id(Without @st.niituniversity.in): ").strip()
        user_data['email'] = email_id
    password = getpass().strip()
    if course_preferences:
        print('Choose a saved course by entering the corresponding number given with it\'s name or enter a new course name.\nSaved Courses:')
        for index, course in enumerate(course_preferences, start = 1):
            print(f'{index}. {course}')
    course_name = input(f"Enter any keyword related to course(Ex. Algorithms from Introduction to Algorithms, Data Structures from Data Structures)(Case Sensitive){' or enter the course number' if course_preferences else ''}: ").strip()
    if course_name.isnumeric() and int(course_name) <= len(course_preferences):
        course_name = course_preferences[int(course_name) - 1]
    else:
        if course_preferences:
            user_data['courses'].append(course_name)
        else:
            user_data['courses'] = [course_name]
    if date_preferences:
        print('Choose a saved date by entering the corresponding number given with it or enter a new date.\nSaved Dates:')
        for index, date in enumerate(date_preferences, start = 1):
            date_split = date.split('/')
            print(f'{index}. {calendar.month_name[int(date_split[1])]} {date_split[0]}, {date_split[2]}')
    course_start_date = input(f"Enter the starting date from which the links have to be collected(DD/MM/YYYY){' or enter the date number' if date_preferences else ''}: ").strip()
    if course_start_date.isnumeric() and int(course_start_date) <= len(date_preferences):
        course_start_date = date_preferences[int(course_start_date) - 1]
    else:
        if date_preferences:
            user_data['dates'].append(course_start_date)
        else:
            user_data['dates'] = [course_start_date]
    course_start_date = datetime.strptime(course_start_date, '%d/%m/%Y').date()
    week_start = course_start_date.today()
    course_data = []
    #Opening Website using desired browser
    while True:
        browser = int(input("Choose your browser:\n1. Chrome\n2. Firefox\nChoose: "))
        if browser == 1:
            if 'chromedriver' not in user_data:
                print('Please enter the path to your chromedriver(including .exe extension). This is to be done only once. The path will be saved for future convenience.')
                user_data['chromedriver'] = input('Enter path: ').strip().replace('\\\\', '\\')
            with open('paths.json', 'w') as paths:
                paths.write(json.dumps(user_data))
            try:
                driver = webdriver.Chrome(user_data['chromedriver'])
            except WebDriverException:
                print('ERROR: Cannot open browser window.\nPlease verify whether the version of Chrome you are using matches the version of the', 
                      'chromedriver that has been installed.\nCheck your chrome version: https://www.google.com/chrome/update/ \nDownload chromedriver: https://chromedriver.chromium.org/downloads', 
                      '\nIf the problem persists, verify whether the path that had been entered for the chromedriver was correct. The path has been deleted from your data.', 
                      'Please choose Chrome again and re-enter the path. Inconvenience is regretted.')
                user_data.pop('chromedriver')
                continue
            break
        elif browser == 2:
            if not user_data or 'geckodriver' not in user_data:
                print('Please enter the path to your geckodriver(including .exe extension). This is to be done only once. The path will be saved for future convenience.')
                user_data['geckodriver'] = input('Enter path: ').strip().replace('\\\\', '\\')
            with open('paths.json', 'w') as paths:
                paths.write(json.dumps(user_data))
            try:
                driver = webdriver.Firefox(executable_path = user_data['geckodriver'])
            except WebDriverException:
                print('ERROR: Cannot open browser window.\nPlease verify whether the version of geckodriver you have installed is suitable for the version of the', 
                      'Firefox that you are using. Download the latest version of geckodriver.\nCheck your Firefox version: https://support.mozilla.org/en-US/kb/find-what-version-firefox-you-are-using \nDownload geckodriver: https://github.com/mozilla/geckodriver/releases', 
                      '\nIf the problem persists, verify whether the path that had been entered for the geckodriver was correct. The path has been deleted from your data.', 
                      'Please choose Firefox again and re-enter the path. Inconvenience is regretted.')
                user_data.pop('geckodriver')
                continue
            break
        else:
            print("Please choose a valid option.")
            continue
    print("Opening Website...")
    driver.get("https://learn.niituniversity.in")
    driver.maximize_window()
    #Logging In with Email ID and Password
    driver.find_element_by_class_name('login-btn').click()
    WebDriverWait(driver, 300).until(
        EC.presence_of_element_located((By.ID, 'identifierId'))
    ).send_keys(email_id)
    while True:
        try:       
            driver.find_element_by_class_name('FliLIb').click()  
            break
        except (NoSuchWindowException, KeyboardInterrupt):
            sys.exit()
        except:
            continue
    while True:
        try:       
            WebDriverWait(driver, 300).until(
                EC.presence_of_element_located((By.NAME, 'password'))
            ).send_keys(password)
            break
        except TimeoutException:
            print('The webpage is taking too much time to load. Please check your internet connection and try again.')
            sys.exit()
        except (NoSuchWindowException, KeyboardInterrupt):
            sys.exit()
        except:
            continue 
    while True:
        try:       
            driver.find_element_by_class_name('FliLIb').click()  
            break
        except (NoSuchWindowException, KeyboardInterrupt):
            sys.exit()
        except:
            continue
    #Navigating to the course calendar
    while True:
        try:
            WebDriverWait(driver, 300).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'calendaricon'))
            ).click()
            break
        except TimeoutException:
            print('The webpage is taking too much time to load. Please check your internet connection and try again.')
            sys.exit()
        except (NoSuchWindowException, KeyboardInterrupt):
            sys.exit()
        except:
            continue
    WebDriverWait(driver, 300).until(
        EC.presence_of_element_located((By.XPATH, "//option[@value = 'day']"))
    ).click()
    WebDriverWait(driver, 300).until(
        EC.presence_of_element_located((By.CLASS_NAME, 'rbc-toolbar-label'))
    )
    #Collecting all class links from the given start date
    print("Collecting Data...")
    course_link = get_class_link(course_name)
    if len(course_link) > 0:
        course_data = course_link + course_data
    while week_start > course_start_date:
        try:
            WebDriverWait(driver, 300).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'nextp-btn'))
            ).click()
        except TimeoutException:
            print('The webpage is taking too much time to load. Please check your internet connection and try again.')
            sys.exit()
        except (NoSuchWindowException, KeyboardInterrupt):
            print_exc()
            sys.exit()
        except:
            continue
        while True:                
            try:        
                course_link = get_class_link(course_name)
                if len(course_link) > 0:
                    course_data = course_link + course_data
                break    
            except (NoSuchWindowException, KeyboardInterrupt):
                print_exc()
                sys.exit()
            except:
                continue  
        week_start -= timedelta(days = 1)
    driver.quit()    
    print("Data Collected")
    #Saving all data to a .csv file with filename taken as input
    count = {'T': 1, 'P': 1, 'L': 1}
    file_data = "Class Number,Date,Day,Link\n"
    for i in course_data:
        file_data += i[0].upper() + " " + str(count[i[0].upper()]) + "," + i[2] + " " + i[3] + "," + i[1] + "," + i[-1] + "\n"
        count[i[0].upper()] += 1
    file_data += ("\nLegend:" + ((count['L'] > 1) * "   L - Lecture") + ((count['P'] > 1) * "   P - Practical") + ((count['T'] > 1) * "   T - Tutorial"))   
    save_as_file(file_data)