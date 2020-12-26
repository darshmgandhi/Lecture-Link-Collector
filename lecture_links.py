import time
import traceback
from getpass import getpass
from selenium import webdriver
from selenium.webdriver.common.by import By 
from datetime import date, datetime, timedelta
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException 

#Getting all the class links of a particular course on a given day
def get_class_link(course_name):
    course_link = []
    course_schedule = driver.find_elements_by_xpath("//div[contains(@title, '" + course_name + "')]//a[@class = 'recordingbtn']")
    if len(course_schedule) > 0:
        class_type = driver.find_elements_by_xpath("//div[contains(@title, '" + course_name + "')]//h6")
        class_date = driver.find_element_by_class_name('rbc-toolbar-label').text
        for i in range(len(course_schedule)):
            course_link.append([class_type[i].text[-2]])
            course_link[-1].extend(class_date.split())
            course_link[-1].append(course_schedule[i].get_attribute("vurl").replace('preview', 'view'))
            course_link[-1][2] = datetime.strptime(course_link[-1][2], '%b').strftime('%B')
    return course_link        

if __name__ == "__main__":
    print("Lecture Link Collector(URL: learn.niituniversity.in, Output File Format: .csv)")
    #Taking user credentials and course details as input
    email_id = input("Email Id(Without @st.niituniversity.in): ").strip()
    password = getpass().strip()
    course_name = input("Enter any keyword related to course(Ex. Algorithms from Introduction to Algorithms, Data Structures from Data Structures): ").strip()
    course_start_date = datetime.strptime(input("Enter starting date of the course(DD/MM/YYYY): ").strip(), '%d/%m/%Y').date()
    week_start = date.today()
    course_data = []
    #Opening Website
    print("Opening Website...")
    driver = webdriver.Chrome("E:\Programming\drivers\chromedriver")
    driver.get("https://learn.niituniversity.in")
    driver.maximize_window()
    #Logging In with Email ID and Password
    driver.find_element_by_class_name('login-btn').click()
    driver.find_element_by_id('identifierId').send_keys(email_id)
    driver.find_element_by_class_name('VfPpkd-RLmnJb').click()
    while True:
        try:       
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, 'password'))
            ).send_keys(password)
            break
        except:
            continue        
    while True:
        try:       
            driver.find_element_by_class_name('VfPpkd-RLmnJb').click()  
            break
        except:
            continue
    #Navigating to the course calendar        
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, 'calendaricon'))
    ).click()
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//option[@value = 'day']"))
    ).click()
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, 'rbc-toolbar-label'))
    )
    #Collecting all class links from the given start date
    print("Collecting Data...")
    course_link = get_class_link(course_name)
    if len(course_link) > 0:
        course_data = course_link + course_data
    while week_start > course_start_date:
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'nextp-btn'))
            ).click()
        except StaleElementReferenceException:
            continue
        while True:                
            try:        
                course_link = get_class_link(course_name)
                if len(course_link) > 0:
                    course_data = course_link + course_data
                break    
            except StaleElementReferenceException:
                continue                
        week_start -= timedelta(days = 1)
    print("Data Collected")
    #Saving all data to a .csv file with filename taken as input
    filename = input("Save As(Enter filename/path without extension): ").strip()
    count = {'T': 1, 'P': 1, 'L': 1}
    f = open(filename + ".csv", "w")
    f.write("Class Number,Date,Day,Link\n")
    for i in course_data:
        f.write((i[0].upper() + " " + str(count[i[0].upper()]) + "," + i[2] + " " + i[3] + "," + i[1] + "," + i[-1] + "\n"))
        count[i[0].upper()] += 1
    f.write(("\nLegend:" + ((count['L'] > 1) * "   L - Lecture") + ((count['P'] > 1) * "   P - Practical") + ((count['T'] > 1) * "   T - Tutorial")))    
    f.close()    