from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
import time
import xlrd
#xlrd==1.2.0

def abrir_arquivo(arquivo):
    data = []
    #Abra e leia um arquivo Excel
    workbook = xlrd.open_workbook(arquivo ,   on_demand = True)
    # obter a primeira planilha
    worksheet = workbook.sheet_by_index(0)
    for row in range(0, worksheet.nrows):
        #obtem a informação que tem numa linha
        data.append(worksheet.row_values(row))
    return data, worksheet

def start():
    web = webdriver.Chrome()
    web.get('https://docs.google.com/forms/d/e/1FAIpQLSdKkB7bu0WXQSRQXrFBNaibhNKybd2a29dASpKyIWqup8IRPg/viewform')
    time.sleep(2)
    
    return web 

def preecher(web,data, worksheet):
    #actionChains = ActionChains(web)

    contratante = web.find_element_by_xpath('//*[@id="mG61Hd"]/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(str(data[1][0]))
    objeto = web.find_element_by_xpath('//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(str(data[1][1]))
    #time.sleep(2)
    Contratante = web.find_element_by_xpath('//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div[2]/textarea').send_keys(str(data[1][2]))
    cpfcnpj = web.find_element_by_xpath('//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(str(data[1][3]))
    Nprocesso = web.find_element_by_xpath('//*[@id="mG61Hd"]/div[2]/div/div[2]/div[5]/div/div/div[2]/div/div[1]/div[2]/textarea').send_keys(str(data[1][4]))
    Local = web.find_element_by_xpath('//*[@id="mG61Hd"]/div[2]/div/div[2]/div[6]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(str(data[1][5]))
    Situacao = web.find_element_by_xpath('//*[@id="mG61Hd"]/div[2]/div/div[2]/div[7]/div/div/div[2]/div/div[1]/div[1]/div[1]').click()
    actions = ActionChains(web)
    actions.send_keys(Keys.ARROW_DOWN)
    actions.send_keys(Keys.ENTER)
    actions.perform()
    Modalidade = web.find_element_by_xpath('//*[@id="mG61Hd"]/div[2]/div/div[2]/div[8]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(str(data[1][7]))
    Fundamento = web.find_element_by_xpath('//*[@id="mG61Hd"]/div[2]/div/div[2]/div[9]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(str(data[1][8]))
    Fonte = web.find_element_by_xpath('//*[@id="mG61Hd"]/div[2]/div/div[2]/div[10]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(str(data[1][9]))
    Ncontrato = web.find_element_by_xpath('//*[@id="mG61Hd"]/div[2]/div/div[2]/div[11]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(str(data[1][10]))
    Link = web.find_element_by_xpath('//*[@id="mG61Hd"]/div[2]/div/div[2]/div[12]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(str(data[1][10]))
    Data = web.find_element_by_xpath('//*[@id="mG61Hd"]/div[2]/div/div[2]/div[13]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(str(data[1][11]))
    Vigencia = web.find_element_by_xpath('//*[@id="mG61Hd"]/div[2]/div/div[2]/div[14]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(str(data[1][12]))
    Vcontrato = web.find_element_by_xpath('//*[@id="mG61Hd"]/div[2]/div/div[2]/div[15]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(str(data[1][13]))
    Vpago = web.find_element_by_xpath('//*[@id="mG61Hd"]/div[2]/div/div[2]/div[16]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(str(data[1][14]))
    Vanulado = web.find_element_by_xpath('//*[@id="mG61Hd"]/div[2]/div/div[2]/div[17]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(str(data[1][15]))
    Submit = web.find_element_by_xpath('//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div/div/span/span').click()
    #time.sleep(20)

    return 0

if __name__ == '__main__':
    data = []
    data, worksheet = abrir_arquivo('COVID-19 - CONSUTAS TRANSPARÊNCIA - alteração leiaute - Copia.xlsx')

    #print(data)
    print(data[0][5])
    web = start()
    preecher(web, data, worksheet)