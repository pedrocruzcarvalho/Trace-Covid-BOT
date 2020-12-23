import selenium
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
import time
import xlrd


def check_exists_by_xpath(xpath):
    try:
        web.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True

   
start_time=time.time()
email_login = input("email : ")
pw = input("password : ")
i2 = input("ID de inicio: ")
i = int(i2)
web=selenium.webdriver.Chrome('chromedriver.exe')


#Entrar no site
web.maximize_window()
time.sleep(0.5)
web.get('https://tracecovid19.min-saude.pt/')
web.find_element_by_link_text('Entrar').click()
time.sleep(7)
web.find_element_by_id("i0116").send_keys(email_login)
web.find_element_by_id("i0116").send_keys(Keys.ENTER)
time.sleep(7)                                        
web.find_element_by_id("i0118").send_keys(pw)
web.find_element_by_id("i0118").send_keys(Keys.ENTER)
time.sleep(7)
web.switch_to.active_element.send_keys(Keys.ENTER)
time.sleep(2)
web.find_element_by_link_text('Trace Covid-19').click()
web.find_element_by_link_text('Pessoas').click()
time.sleep(1)
web.find_element_by_xpath("//div[@class='toggle btn btn-primary']").click()
time.sleep(1)


path = 'Base_Mestre.xlsx'
wb = xlrd.open_workbook(path)
sheet = wb.sheet_by_index(0)

i-=1
curados=0
inserir=0
sns_errado=0
sobreativo=0
action=0

while(sheet.cell_value(i,4) != "" ):
    i +=1
    tempsns = int(sheet.cell_value (i,4))
    str1 = str(tempsns)
    if(len(str1) == 9 and str1[0] != '0'):
        web.find_element_by_id("PatientNumber").send_keys(9 * Keys.BACKSPACE)
        web.find_element_by_id("PatientNumber").send_keys(str1)
        web.find_element_by_id("filterSearch").click()
        time.sleep(1.5)
        while check_exists_by_xpath("//td[@class='table-results waitMe_container']")==True:  
            time.sleep(0.5)
        if check_exists_by_xpath("//td[@class='dataTables_empty']")==False:
            x= web.find_element_by_xpath("//table[@id='tablePerson']//td[9]").text
            with open('telefone.txt', 'a') as g: 
                g.write(str1 + ';' + x + '\n')

                
