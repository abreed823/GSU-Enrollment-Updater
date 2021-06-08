from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import time
import scrape_courses

base_url = 'https://www.gosolar.gsu.edu/bprod/bwckschd.p_disp_dyn_sched'
driver = webdriver.Chrome('/Users/aprilbreedlove/Downloads/PythonGSU/chromedriver')
driver.maximize_window()
driver.get(base_url)
driver.implicitly_wait(10)


# Functions that operate on the landing page
class PageOne:
    def select_term(self):
        semester_menu = driver.find_element_by_name('p_term')
        semester = Select(semester_menu)

        # Selects the most current semester which is always at index 0
        semester.select_by_index('1')

        time.sleep(0.5)

        # Navigates to next page once semester is selected
        next_page = driver.find_element_by_xpath("//input[@type='submit' and @value='Submit']")
        next_page.click()


# Functions that operate on the second page
class PageTwo:
    # All time.sleep() commands are for visual purposes only so that I can physically see the program work slowly

    # Selects associates degree
    def select_degree(self):
        degree_menu = driver.find_element_by_id('levl_id')
        degree = Select(degree_menu)
        # "Associates Degree" is at index 0
        degree.select_by_index('0')

    # Selects all subjects
    def select_subjects(self):
        # Defines first and last elements to be clicked during shift-click
        first_subject = driver.find_element_by_xpath("//option[@value='ACCT']")
        last_subject = driver.find_element_by_xpath("//option[@value='WLC']")
        select_all = ActionChains(driver)

        # Performs shift-click operation
        select_all\
            .click(first_subject)\
            .key_down(Keys.SHIFT)\
            .click(last_subject)\
            .key_up(Keys.SHIFT)\
            .perform()

    # Selects the six campuses
    def select_campus(self):
        campus_menu = driver.find_element_by_id('camp_id')
        campus = Select(campus_menu)

        # Deselects 'All" option which is automatically selected by the website
        campus.deselect_by_index('0')
        # Alpharetta
        campus.select_by_value('PA')
        # Clarkston
        campus.select_by_value('PC')
        # Decatur
        campus.select_by_value('PS')
        # Dunwoody
        campus.select_by_value('PN')
        # Newton
        campus.select_by_value('PE')
        # Online
        campus.select_by_value('PF')

    # Selects just honors as course type
    def select_course_type(self):
        course_type_menu = driver.find_element_by_id('attr_id')
        course_type = Select(course_type_menu)

        # Deselects 'All" option which is automatically selected by the website
        course_type.deselect_by_index('0')
        # Selects Honors College courses
        course_type.select_by_value('HP')

    # Automatically selects appropriate criteria to filter class list
    def filter_classes(self, file):
        self.select_degree()
        self.select_subjects()
        self.select_campus()
        self.select_course_type()

        # Navigates to next page once criteria are selected
        next_page = driver.find_element_by_xpath("//input[@type='submit' and @value='Class Search']")
        next_page.click()

        # Defines html source code and passes it to scrape_courses.py which creates soup object
        html = driver.page_source
        page_3 = scrape_courses.Courses()
        page_3.create_data_frame(html, file)

