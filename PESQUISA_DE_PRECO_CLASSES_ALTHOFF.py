from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
import openpyxl


    



    
lista_produtos = list()
lista_preco = list()
lista_desconto = list()


s = Service("C:\Program Files (x86)\chromedriver.exe")
driver = webdriver.Chrome(service=s)
wait = WebDriverWait(driver, 10)

class PesqPrice:  
    
    def pesquisar_url(url): 
        driver.get(f"{url}")


    def local(XPATH):
        driver.maximize_window()
        element = wait.until(EC.element_to_be_clickable((By.XPATH, f'{XPATH}')))

        local = driver.find_element(By.XPATH, f'{XPATH}')
        local.click()


        element = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="body"]/div[2]/div[3]/div/div[4]/div/button')))
        confirmar = driver.find_element(By.XPATH,'//*[@id="body"]/div[2]/div[3]/div/div[4]/div/button')
        confirmar.click()


        element = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="root"]/div[2]/div/div/button')))
        cookie_button = driver.find_element(By.XPATH, '//*[@id="root"]/div[2]/div/div/button')
        cookie_button.click()


        sleep(2)
 

    def categ_althoff():
        categs = driver.find_elements(By.CSS_SELECTOR, '.corridor-content-wrapper a')

        cont = 0
        lista_link = list()
        for a in categs:
            cont += 1
            link = a.get_attribute('href')
            teste = link.split('categorias/')[1]
            if cont == len(categs):
                lista_link.append(link)
         
            elif teste not in categs[cont].get_attribute('href'):
                lista_link.append(link)

        for link in lista_link:
            driver.get(link)
            cont_prod = 0
            lista_id_prod = list()
            
            while True:
                cont_prod += 1
            
                
                
                try:
                    id_preco = driver.find_element(By.XPATH, f'//*[@id="page-scroll-element"]/div[2]/div[1]/div/div/div[3]/div/div/div/div/div/div[1]/div[{cont_prod}]/div').get_attribute('data-product-id')
                    lista_id_prod.append(id_preco)
                    if str(id_preco) != 'None':
                        preco = driver.find_element(By.XPATH, f'//*[@id="product-{id_preco}-price"]/div[1]').text
                        descricao = driver.find_element(By.XPATH, f'//*[@id="product-{id_preco}-price"]/div[3]/span').text
                        print(descricao)
                        
                        try:
                            preco_desconto = driver.find_element(By.XPATH, f'//*[@id="product-{id_preco}-price"]/a/span').text
                            
                        except:
                            
                            preco_desconto = driver.find_element(By.XPATH, f'//*[@id="product-{id_preco}-price"]/div[2]/span').text
                                
                                
                            if len(preco_desconto.split('R$')) == 1:
                                preco_desconto = ''
                                
                                
                        
                        lista_produtos.append(descricao)
                        lista_preco.append(preco)
                        lista_desconto.append(preco_desconto)
            
                                
                                
                            
                            
                    
                except:
                    if cont_prod > 10:
                        cont_prod -= 1
                        
                        descer_pag = driver.find_element(By.XPATH, f'//*[@id="add-{lista_id_prod[len(lista_id_prod)-1]}-to-cart-btn"]')
                        descer_pag.send_keys(Keys.PAGE_DOWN)
                        sleep(1)
                        subir_pag = driver.find_element(By.XPATH, f'//*[@id="add-{lista_id_prod[len(lista_id_prod)-3]}-to-cart-btn"]')
                        subir_pag.send_keys(Keys.PAGE_UP)
                        try:
                            id_preco_novo = driver.find_element(By.XPATH, f'//*[@id="page-scroll-element"]/div[2]/div[1]/div/div/div[3]/div/div/div/div/div/div[1]/div[{cont_prod+3}]/div').get_attribute('data-product-id')
                            print(id_preco_novo)
                            element = wait.until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="add-{id_preco_novo}-to-cart-btn"]')))
                            
                        except:
                            break
                        
                    else:
                        break
                
                                
                
                    
 

                        
                    
                        
                        
    def add_excel(nome_planilha):
        book = openpyxl.Workbook()
        mercadorias = book['Sheet']
        print(len(lista_produtos))
        mercadorias["A1"] = "ITEM"
        mercadorias["B1"] = "PREÇO"
        mercadorias["C1"] = "PREÇO ANTERIOR"
        for c in range(0,len(lista_produtos)):
            #print(c)
            d = c + 2
            produto_excel = lista_produtos[c]
            preco_excel = lista_preco[c]
            desconto_excel = lista_desconto[c]
            mercadorias[f"A{d}"] = produto_excel
            mercadorias[f"B{d}"] = preco_excel
            mercadorias[f"C{d}"] = desconto_excel
            
        book.save(f'{nome_planilha}')
        print("ACABou ")
