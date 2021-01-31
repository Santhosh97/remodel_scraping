from selenium.webdriver import Chrome,ChromeOptions
from selenium.webdriver.support.select import Select
from time import sleep
from random import randint
import csv
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
chome_driver_path = r'./ChromeDriver/chromedriver.exe'
csv_columns = ['State','NameOfFirm','URL','PrimaryContact','Address','Phone','ListOfStaff','Website','Specialities']
csv_file = "output.csv"
opts =  ChromeOptions()

# Argument for maximizing the screensize, It's necessary so that, all required data is displayed on the browser.
opts.add_argument("start-maximized")
#Creating output File Writing Header
try:
    f = open(csv_file)
    f.close()
except IOError:
    try:
        with open(csv_file, 'w', encoding='utf-8', newline='') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=csv_columns, delimiter=',')
            writer.writeheader()
    except IOError:
        print("I/O error")

#Open Chrome Browser
print('Opening Chrome Browser Window...')
driver = Chrome(executable_path=chome_driver_path,options=opts)
driver.get('https://remodelingdoneright.nari.org/remodelers/')

# Get the list of all states from the dropdown menu
states = driver.find_elements_by_xpath("//*[@id='maincontent']/div/div[1]/div/div[2]/div[1]/div[1]/select/option")
states_names = [s.text for s in states]

def Search_State(d,s):
    print(f"Searching for {s}")

    #Select the state from the dropdown menu
    dropdown = WebDriverWait(d, 60).until(EC.presence_of_element_located((By.XPATH, "//label[text()='State']/parent::div/select")))
    dropdown_menu = Select(dropdown)
    dropdown_menu.select_by_visible_text(s)

    #Select the range from dropdown menu

    range = WebDriverWait(d, 60).until(EC.presence_of_element_located((By.XPATH, "//div[@class='filter-item select'][2]/select")))
    range_menu = Select(range)
    range_menu.select_by_visible_text('100 miles')

    #Click Search Button
    search_b = WebDriverWait(d, 60).until(EC.presence_of_element_located((By.XPATH, "//input[@value='Search']")))
    search_b.click()


def Get_Remodelers_Data(d,s):

    #Getting the the sections that contains remodeler data

    remodeler_cards = WebDriverWait(d, 60).until(
            EC.presence_of_all_elements_located((By.XPATH, "//div[@class='find-remodeler-card']"))
        )

    # remodeler_cards = d.find_elements_by_xpath("//div[@class='find-remodeler-card']")
    for card in remodeler_cards:
        final_data = dict()
        final_data['State'] = s
        try:
            final_data['NameOfFirm'] = card.find_element_by_xpath(".//h3[@class='find-remodeler-title']/a").text
        except:
            final_data['NameOfFirm'] = 'n.a'

        try:
            final_data['URL'] = card.find_element_by_xpath(".//h3[@class='find-remodeler-title']/a").get_attribute('href')
        except:
            final_data['URL'] = 'n.a'

        try:
            final_data['PrimaryContact'] = card.find_element_by_xpath(".//p[text()='Primary Contact']/parent::h3/following::p").text
        except:
            final_data['PrimaryContact'] = 'n.a'

        try:
            final_data['Address'] = card.find_element_by_xpath(".//h3[text()='Address']/parent::div/following::p[1]").text
        except:
            final_data['Address'] = 'n.a'
        try:
            final_data['Phone'] = card.find_element_by_xpath(".//h3[text()='Phone']/parent::div/following::p[1]").text
        except:
            final_data['Phone'] = 'n.a'
        try:
            pss = card.find_elements_by_xpath(".//div[@class='inner-col-middle']/p")
            final_data['ListOfStaff'] = '; '.join([p.text for p in pss])
        except:
            final_data['ListOfStaff'] = 'n.a'

        try:
            final_data['Website'] = card.find_element_by_xpath(".//div[@class='inner-col-middle']/a").get_attribute('href')
        except:
            final_data['Website'] = 'n.a'

        try:
            sps = card.find_element_by_xpath(".//*[@id='column-wrap-first']/div[2]/ul").text
            final_data['Specialities'] = sps.replace('\n','; ')
        except:
            final_data['Specialities'] = 'n.a'

        for k,v in final_data.items():
            print(f'{k} - {v}')
        print('\n')
        try:
            with open(csv_file, 'a', encoding='utf-8', newline='') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=csv_columns, delimiter=',')
                writer.writerow(final_data)
        except IOError:
            print("I/O error")


for state in states_names[1:]:
    #Search
    Search_State(driver,state)
    sleep(randint(5,10))

    last = False
    try:
        #Check the number of last page of search result
        last = int(driver.find_element_by_xpath("//div[@class='page-filter__right']/div/span[last()]").text.strip('of').strip())
    except:
        #If no value is present for last number of page, Then the search result is empty
        print(f'No Company Found for {state}')
    #If search result is not empty
    if(last):

        #Looping through all the search result pages
        for i in range(0,last):
            print(f'Page no {str(i+1)} of {str(last)}')

            Get_Remodelers_Data(driver,state)

            # Crawling to the next page
            next_page = driver.find_element_by_xpath("//div[@class='page-filter__right']/div/a[last()]")
            next_page.click()
            sleep(randint(5,10))
    sleep(randint(5,10))




