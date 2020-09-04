from selenium import webdriver
import pyautogui
import time



class OnedriveBot:
    def __init__(self, a, b):
        self.a = "email"      #coloca aqui o seu email
        self.b = "passe"      #coloca aqui a sua palavra passe
        PATH = "C:\Program Files (x86)\chromedriver.exe"     #localização do drive do chrome
        self.driver = webdriver.Chrome(PATH)            #Iniciar o Chorme quando rodar

    def login(self):
        driver = self.driver
        driver.get(
            "https://login.live.com/login.srf?wa=wsignin1.0&rpsnv=13&ct=1598850431&rver=7.3.6962.0&wp=MBI_SSL_SHARED&wreply=https:%2F%2Fonedrive.live.com%3Fgologin%3D1%26mkt%3Dpt-BR&lc=1046&id=250206&cbcxt=sky&mkt=pt-BR&lw=1&fl=easi2&username=" + self.a)
        driver.find_element_by_id('i0118').send_keys(self.b)  # inser password
        time.sleep(2)
        driver.find_element_by_id("idSIButton9").click()  # confirm
        time.sleep(2)
        driver.find_element_by_id("KmsiCheckboxField").click()  # checkbox
        time.sleep(2)
        driver.find_element_by_id("idSIButton9").click()  # 'SIM'

    def getdados(self):
        driver = self.driver
        time.sleep(4)
        # //span[@class='ms-Tile-label label_9c46d494']         ======info:: Anexos de email, Pasta, Modificado 12 de mai. de 2017, 4,13 KB, 1 item, Particular
        principal = driver.find_element_by_class_name('od-TilesList')
        # print('principal: '+str(principal))      print element
        name = driver.find_elements_by_xpath("//span[@data-automationid='name']")
        items = principal.find_elements_by_tag_name('a')
        i = 0
        for item in items:
            href = item.get_attribute('href')
            print('NAME:  ' + name[i].text)
            print('HREF:  ' + href + '\n')
            i = i + 1
        time.sleep(3)
        driver.get("https://onedrive.live.com/?id=root&cid=5C6B00A7D5013266#id=5C6B00A7D5013266%21102&cid=5C6B00A7D5013266")

    def upload(self):
        driver = self.driver
        time.sleep(3)
        driver.find_element_by_xpath("//button[@name ='Carregar']").click()
        time.sleep(2)
        driver.find_element_by_xpath("//button[@name ='Arquivos']").click()

    def move(self):
        driver = self.driver
        pyautogui.press(['tab', 'tab', 'tab', 'tab', 'tab'])
        pyautogui.press(['enter'])
        time.sleep(3)
        pyautogui.write('Diretório da planilha') #C:\Users\...
        pyautogui.press(['enter'])
        time.sleep(3)
        pyautogui.write('nome do ficheiro') # Exemplo nome.xlsx
        pyautogui.press(['enter'])
        time.sleep(2)
        driver.find_element_by_xpath("//button[@title ='Mostrar erro']").click()
        time.sleep(2)
        driver.find_element_by_class_name("od-Button-label").click()
        time.sleep(5)
        driver.quit()


bot = OnedriveBot("LOGIN", "SENHA")
bot.login()
bot.getdados()
bot.upload()
bot.move()