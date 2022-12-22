from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By


data = ""

class execute:
    def __init__(self, text:str):
        global data
        navegador = webdriver.Chrome()
        navegador.minimize_window()
        navegador.get('http://cadsinc.sefaz.al.gov.br/ConsultarDadosContribuinte.do?action=VisualizarDadosContribuinte.do')
        navegador.find_element(By.XPATH, '//*[@id="idTipoDocumento"]').click()
        navegador.find_element(By.XPATH, '/html/body/form/table/tbody/tr[1]/td[1]/select/option[4]').click()
        element = navegador.find_element(By.XPATH, '//*[@id="valor"]')
        element.send_keys(f'{text}')
        navegador.find_element(By.XPATH, '/html/body/form/table/tbody/tr[2]/td/input').click()
        navegador.find_element(By.XPATH, '//*[@id="item"]/tbody/tr/td[6]/a').click()
        nome = navegador.find_element(By.XPATH, '/html/body/table[3]/tbody/tr/td/fieldset[1]/table[1]/tbody/tr[4]/td/b')
        empresa = navegador.find_element(By.XPATH, '/html/body/table[3]/tbody/tr/td/fieldset[1]/table[1]/tbody/tr[6]/td')
        cidade = navegador.find_element(By.XPATH, '/html/body/table[3]/tbody/tr/td/fieldset[2]/table[2]/tbody/tr[2]/td[3]')
        if cidade.text == 'LIMOEIRO DE ANADIA':
            data = [nome.text, empresa.text, cidade.text]

        else:
            data = [nome.text, empresa.text, cidade.text]
def datareturn():
    return data
if __name__ == "__main__":
    execute

