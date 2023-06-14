import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep


class PesqPrice:
    def __init__(self, driver_path):
        self.driver = self._create_driver(driver_path)
        self.wait = WebDriverWait(self.driver, 2)
        self.lista_produtos = []
        self.lista_preco = []
        self.lista_desconto = []

    def _create_driver(self, driver_path):
        service = Service(driver_path)
        return webdriver.Chrome(service=service)

    def pesquisar_url(self, url):
        self.driver.get(url)

    def local(self, xpath):
        self.driver.maximize_window()
        element = self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        local = self.driver.find_element(By.XPATH, xpath)
        local.click()

        element = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="body"]/div[2]/div[3]/div/div[4]/div/button')))
        confirmar = self.driver.find_element(By.XPATH, '//*[@id="body"]/div[2]/div[3]/div/div[4]/div/button')
        confirmar.click()

        element = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="root"]/div[2]/div/div/button')))
        cookie_button = self.driver.find_element(By.XPATH, '//*[@id="root"]/div[2]/div/div/button')
        cookie_button.click()

        sleep(2)

    def categ_althoff(self):
        categs = self.driver.find_elements(By.CSS_SELECTOR, '.corridor-content-wrapper a')

        lista_link = []
        for cont, a in enumerate(categs, start=1):
            link = a.get_attribute('href')
            teste = link.split('categorias/')[1]
            if cont == len(categs) or teste not in categs[cont].get_attribute('href'):
                lista_link.append(link)

        for link in lista_link:
            self.driver.get(link)
            lista_id_prod = []
            cont_prod = 0
            teste = 0

            while True:
                cont_prod += 1

                try:
                    if teste == 1:
                        cont_prod.split('da')

                    element = self.wait.until(EC.element_to_be_clickable((By.XPATH,
                                                                           '//*[@id="page-scroll-element"]/div[2]/div[1]/div/div/div[3]/div/div/div/div/div/div[1]/div[1]/div/div/div[2]/button/i[1]')))
                    teste = 0

                    try:
                        id_preco = self.driver.find_element(By.XPATH,
                                                            f'//*[@id="page-scroll-element"]/div[2]/div[1]/div/div/div[3]/div/div/div/div/div/div[1]/div[{cont_prod}]/div').get_attribute(
                            'data-product-id')
                        lista_id_prod.append(id_preco)
                        if id_preco is not None:
                            preco = self.driver.find_element(By.XPATH,
                                                             f'//*[@id="product-{id_preco}-price"]/div[1]').text
                            descricao = self.driver.find_element(By.XPATH,
                                                                 f'//*[@id="page-scroll-element"]/div[2]/div[1]/div/div/div[3]/div/div/div/div/div/div[1]/div[{cont_prod}]/div/div/div[3]/img').get_attribute(
                                'alt')

                            try:
                                preco_desconto = self.driver.find_element(By.XPATH,
                                                                         f'//*[@id="product-{id_preco}-price"]/a/span').text

                            except:
                                preco_desconto = self.driver.find_element(By.XPATH,
                                                                         f'//*[@id="product-{id_preco}-price"]/div[2]/span').text

                                if len(preco_desconto.split('R$')) == 1:
                                    preco_desconto = ''

                            self.lista_produtos.append(descricao)
                            self.lista_preco.append(preco)
                            self.lista_desconto.append(preco_desconto)
                            print(descricao, preco_desconto, preco)

                    except:
                        if cont_prod > 10:
                            cont_prod -= 1
                            descer_pag = self.driver.find_element(By.XPATH,
                                                                  f'//*[@id="add-{lista_id_prod[len(lista_id_prod) - 1]}-to-cart-btn"]')
                            descer_pag.send_keys(Keys.PAGE_DOWN)
                            sleep(1)
                            subir_pag = self.driver.find_element(By.XPATH,
                                                                 f'//*[@id="add-{lista_id_prod[len(lista_id_prod) - 3]}-to-cart-btn"]')
                            subir_pag.send_keys(Keys.PAGE_UP)

                            try:
                                id_preco_novo = self.driver.find_element(By.XPATH,
                                                                        f'//*[@id="page-scroll-element"]/div[2]/div[1]/div/div/div[3]/div/div/div/div/div/div[1]/div[{cont_prod + 3}]/div').get_attribute(
                                    'data-product-id')
                                element = self.wait.until(EC.element_to_be_clickable(
                                    (By.XPATH, f'//*[@id="add-{id_preco_novo}-to-cart-btn"]')))

                            except:
                                break

                        else:
                            break

                except:
                    try:
                        element = self.wait.until(EC.element_to_be_clickable((By.XPATH,
                                                                               '//*[@id="page-scroll-element"]/div[2]/div[2]/div/div/div[3]/div/div/div/div/div/div[1]/div[1]/div/div/div[2]/button/i[1]')))
                        teste = 1

                        try:
                            id_preco = self.driver.find_element(By.XPATH,
                                                                f'//*[@id="page-scroll-element"]/div[2]/div[2]/div/div/div[3]/div/div/div/div/div/div[1]/div[{cont_prod}]/div').get_attribute(
                                'data-product-id')
                            lista_id_prod.append(id_preco)
                            if id_preco is not None:
                                preco = self.driver.find_element(By.XPATH,
                                                                 f'//*[@id="product-{id_preco}-price"]/div[1]').text
                                descricao = self.driver.find_element(By.XPATH,
                                                                     f'//*[@id="page-scroll-element"]/div[2]/div[2]/div/div/div[3]/div/div/div/div/div/div[1]/div[{cont_prod}]/div/div/div[3]/img').get_attribute(
                                    'alt')

                                try:
                                    preco_desconto = self.driver.find_element(By.XPATH,
                                                                             f'//*[@id="product-{id_preco}-price"]/a/span').text

                                except:
                                    preco_desconto = self.driver.find_element(By.XPATH,
                                                                             f'//*[@id="product-{id_preco}-price"]/div[2]/span').text

                                    if len(preco_desconto.split('R$')) == 1:
                                        preco_desconto = ''

                                self.lista_produtos.append(descricao)
                                self.lista_preco.append(preco)
                                self.lista_desconto.append(preco_desconto)
                                print(descricao, preco_desconto, preco)

                        except:
                            if cont_prod > 10:
                                cont_prod -= 1
                                descer_pag = self.driver.find_element(By.XPATH,
                                                                      f'//*[@id="add-{lista_id_prod[len(lista_id_prod) - 1]}-to-cart-btn"]')
                                descer_pag.send_keys(Keys.PAGE_DOWN)
                                sleep(1)
                                subir_pag = self.driver.find_element(By.XPATH,
                                                                     f'//*[@id="add-{lista_id_prod[len(lista_id_prod) - 3]}-to-cart-btn"]')
                                subir_pag.send_keys(Keys.PAGE_UP)

                                try:
                                    id_preco_novo = self.driver.find_element(By.XPATH,
                                                                            f'//*[@id="page-scroll-element"]/div[2]/div[2]/div/div/div[3]/div/div/div/div/div/div[1]/div[{cont_prod + 3}]/div').get_attribute(
                                        'data-product-id')
                                    element = self.wait.until(EC.element_to_be_clickable(
                                        (By.XPATH, f'//*[@id="add-{id_preco_novo}-to-cart-btn"]')))

                                except:
                                    break

                            else:
                                break

                    except:
                        break

    def add_excel(self,nome_planilha):
        book = openpyxl.Workbook()
        mercadorias = book['Sheet']
        print(len(self.lista_produtos))
        mercadorias["A1"] = "ITEM"
        mercadorias["B1"] = "PREÇO"
        mercadorias["C1"] = "PREÇO ANTERIOR"
        for c in range(0, len(self.lista_produtos)):
            # print(c)
            d = c + 2
            produto_excel = self.lista_produtos[c]
            preco_excel = self.lista_preco[c]
            desconto_excel = self.lista_desconto[c]
            mercadorias[f"A{d}"] = produto_excel
            mercadorias[f"B{d}"] = preco_excel
            mercadorias[f"C{d}"] = desconto_excel

        book.save(f'{nome_planilha}')
        print("ACABou ")


def main():
    driver_path = 'C:\Program Files (x86)\chromedriver.exe'

    url = 'https://emcasa.althoff.com.br/categorias'

    pesq = PesqPrice(driver_path)
    pesq.pesquisar_url(url)
    pesq.local('//*[@id="body"]/div[2]/div[3]/div/div[3]/div/div[2]/div[4]/div[2]')

    pesq.categ_althoff()

    # Faça o processamento dos dados conforme necessário

    # Adicione os dados a um arquivo Excel
    pesq.add_excel('dados.xlsx')

    # Feche o driver ao finalizar
    pesq.driver.quit()


if __name__ == '__main__':
    main()