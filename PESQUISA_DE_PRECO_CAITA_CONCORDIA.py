from PESQUISA_DE_PRECO_CLASSES_OSUPER import PesqPrice
from datetime import date
from time import sleep
import time

start_time = time.time()
driver_path = 'C:\Program Files (x86)\chromedriver.exe'

url = 'https://caitasupermercados.com.br/categorias'

pesq = PesqPrice(driver_path)
pesq.pesquisar_url(url)
pesq.local('//*[@id="body"]/div[2]/div[3]/div/div[3]/div/div[2]/div[3]/div[2]')


pesq.categ_althoff()


pesq.add_excel(f'PLANILHA ALTHOFF-LAGUNAs-{date.today()}.csv')

pesq.driver.quit()

end_time = time.time()

execution_time = end_time - start_time

print("Tempo de execução: ", execution_time, "segundos ")

