import selenium
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
import time
import xlrd
from datetime import datetime
import time
import xlwt
from xlutils.copy import copy
import pandas as pd
import xlsxwriter

def check_exists_by_xpath(xpath):
    try:
        web.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True



if __name__ == "__main__":
 
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
    time.sleep(3)
    web.find_element_by_id("i0116").send_keys(email_login)
    web.find_element_by_id("i0116").send_keys(Keys.ENTER)
    time.sleep(3)                                        
    web.find_element_by_id("i0118").send_keys(pw)
    web.find_element_by_id("i0118").send_keys(Keys.ENTER)
    time.sleep(3)
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
        if(sheet.cell_value(i,4) != "" and sheet.cell_value(i,12)== "LIVRE"):
            tempsns = int(sheet.cell_value (i,4))
            str1 = str(tempsns)
            if(len(str1) != 9 or str1[0] == '0'):
                sns_errado +=1
                with open('textfile.txt', 'a') as g:
                    g.write('SNS errado %s\n' %(str1))
                continue
            else:
                web.find_element_by_id("PatientNumber").send_keys(9 * Keys.BACKSPACE)
                web.find_element_by_id("PatientNumber").send_keys(str1)
                web.find_element_by_id("filterSearch").click()
                time.sleep(1.5)
                while check_exists_by_xpath("//td[@class='table-results waitMe_container']")==True:
                    time.sleep(0.5)  
                if check_exists_by_xpath("//td[@class='dataTables_empty']")==True:
                    inserir +=1
                    with open('textfile.txt', 'a') as g:
                        g.write('Não foi encontrado. Inserir  %s\n' %(str1))  
                    continue
                if check_exists_by_xpath("//table[@id='tablePerson']//td[contains(text(), 'Curado')]")==True:
                    curados+=1
                    x = web.find_element_by_xpath("//table[@id='tablePerson']//td[9]").text
                    with open('textfile.txt', 'a') as g:
                        g.write('Curado: %s' %(str1) + " " +'; Telefone: '+(x) + '\n')
                    continue
                if check_exists_by_xpath("//table[@id='tablePerson']//td[contains(text(), 'Vigilância Sobreativa (MGF)')]")==True and check_exists_by_xpath("//table[@id='tablePerson']//td[contains(text(), 'Positivo')]")==True :
                    sobreativo +=1
                else:
                    action +=1
                    x = web.find_element_by_xpath("//table[@id='tablePerson']//td[9]").text
                    with open('textfile.txt', 'a') as g: 
                        g.write('É preciso acao humana: %s' %(str1) + " " +'; Telefone: '+(x) + '\n')
                    continue

   
    df = pd.read_excel('dados.xlsx', sheet_name='Data')
    df.loc[0,'Número total']= int(df.loc[0,'Número total'])+ sobreativo
    df.loc[1,'Número total']= int(df.loc[1,'Número total'])+ curados
    df.loc[2,'Número total']= int(df.loc[2,'Número total'])+ sns_errado
    df.loc[3,'Número total']= int(df.loc[3,'Número total'])+ action
    df.loc[4,'Número total']= int(df.loc[4,'Número total'])+ inserir

    writer = pd.ExcelWriter('dados.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Data')
    writer.save()



    print("demorou " + str(time.time()-start_time) + " segundos")
