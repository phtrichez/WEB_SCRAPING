from PESQUISA_DE_PRECO_CLASSES_ALTHOFF import PesqPrice 
from datetime import date
from time import sleep
import time

start_time = time.time()






PesqPrice.pesquisar_url("https://emcasa.althoff.com.br/categorias")


PesqPrice.local('//*[@id="body"]/div[2]/div[3]/div/div[3]/div/div[2]/div[4]/div[2]')


PesqPrice.categ_althoff()


PesqPrice.add_excel(f'PLANILHA ALTHOFF-LAGUNAs-{date.today()}.csv')

end_time = time.time()

execution_time = end_time - start_time

print("Tempo de execução: ", execution_time, "segundos ")

