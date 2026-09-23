from playwright.sync_api import Playwright, sync_playwright, expect
from bs4 import BeautifulSoup
import re
import xlsxwriter

username = input("Enter username: ")
password = input("Enter password: ")


#launch playwright
with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
#go to webtop and login
    page.goto("https://webtop.smartschool.co.il/account/login?stateUrl=U2FsdGVkX1%2FJ2NiuxYHTfvC5qMRV9UjY0JeI0TchZFc%3D")
    page.get_by_role("button", name="אשר cookies").click()
    page.get_by_role("button", name="הזדהות משרד החינוך").click()
#fill in username
    page.wait_for_load_state('networkidle')
    page.locator("#blocker").click()
    page.get_by_placeholder("קוד משתמש").click()
    page.get_by_placeholder("קוד משתמש").fill(username)
#fill in password
    page.get_by_placeholder("סיסמה").click()
    page.get_by_placeholder("סיסמה").fill(password)
#log in
    page.get_by_role("button", name="כניסה").click()
    page.wait_for_load_state('networkidle')
#go to grade list and sort by class
    page.get_by_role("link", name="כרטיס תלמיד").click()
    page.get_by_role("link", name="ציונים שוטפים").click()
    page.get_by_text("מחצית ב`").click()
    page.get_by_text("תקופה א`").click()
    page.get_by_label("אפשרויות מיון").click()
    page.get_by_role("menuitem", name="לפי מקצוע").click()
#scrape html code for the table of the grades
    page.wait_for_load_state('networkidle')
    html = page.inner_html('#mainBody > app-root > div > app-app-nav > div:nth-child(2) > mat-sidenav-container > mat-sidenav-content > app-pupil-card > mat-sidenav-container > mat-sidenav-content > app-multi-cards-view > div > div > div > app-grades > div')
    soup = BeautifulSoup(html, 'html.parser')
#define html code for grades, class names, and weight of the grades
    gradesraw = soup.find_all("span", {"class": "col l2 s6 mobile-left-align ng-star-inserted"})
    classesraw = soup.find_all("span", {"class": "col l3 s12"})
    percentraw = soup.find_all("span", {"class": "col l1 s6"})
#create list of classes
    classes = []
    class1 = ''
    for item in classesraw:
        class1 = item.text #"text" finds the the text value of the html code of the names of the classes (finds what actually appears in the website)
        if class1: #if class1 is not empty
            classes.append(class1)


    grades = []
    num = 0
    str1 = ''
    str2 = ''

    for item in gradesraw:
        str1 = item.text #finds the value of the html code of the grades
        if str1.strip(): #finds if its not empty after removing spaces (the function "strip" removes any spaces)
            str2 = re.sub(r'[^0-9]', '', str1) #removes non digits
        if str2: #if str2 is not empty
            num = int(str2)
            grades.append(num)  
        else:
            grades.append(str1)


    percents = []
    percent1 = ''
    percent2 = ''
    for item in percentraw:
        percent1 = item.text
        if percent1.strip():
            percent2 = re.sub(r'[^0-9%]', '', percent1)
            percents.append(percent2)
        else:
            percents.append(percent1)


    class_grade = []
    for class_name, percent, grade in zip(classes, grades, percents): #combines all data to a list
        class_grade.append([class_name, grade, percent])


    clean_class_grade = []
    current_class = ''
#creates a list of lists for each class. every list for a class, include a list of grade and weight of that grade (i know its weird)
    for item in class_grade:
        if current_class != item[0]:
            clean_class_grade.append([item[0], [item[1], item[2]]])
        else:
            clean_class_grade[-1].append([item[1], item[2]])
            
        current_class = item[0]
    print(clean_class_grade)


#converts database to an excel file
    workbook = xlsxwriter.Workbook("Grades.xlsx")
    worksheet = workbook.add_worksheet("WebTop")
    worksheet.right_to_left()

    worksheet.write(0, 0, 'שיעור')
    worksheet.write(0, 1, 'ציון')
    worksheet.write(0, 2, 'משקל')

    worksheet.write_column(1, 0, classes)
    worksheet.write_column(1, 1, grades)
    worksheet.write_column(1, 2, percents)

    workbook.close()
#converts weight to integer, so it can be used for calculations later.
    weighed_class_grade =  clean_class_grade
    for item in weighed_class_grade:
        for x in item[1:]:
            if isinstance(x[0], str) and '%' in x[0]:
                x[0] = int(x[0].replace('%', ''))
            else:
                x[0] = 3

    print('')
    print(weighed_class_grade)

    percentint = 0
    sum1 = 0
    avrages = []
    avg = 0
    sum2 = 0
    gensum = 0
    for item in weighed_class_grade:
        sum1 = 0
        sum2 = 0
        for x in item:
            if isinstance(x, str):
                avrages.append([x])
            else:
                percentint = x[0]/100
                sum1 += x[1] * percentint
                sum2 += x[0]
        
        avrages[-1].append(sum1/sum2*100)
        gensum += sum1/sum2*100

    print('')
    print(avrages)
    print('General avrage:', gensum/len(avrages))