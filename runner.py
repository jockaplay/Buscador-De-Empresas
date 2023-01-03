from selenium import webdriver
from selenium.webdriver.common.by import By

data = []
ok = True


class Execute:
    def acao(self, codes: list, cidades: list, tipo: int):
        global data
        for text in codes:
            self.navegador.get('http://cadsinc.sefaz.al.gov.br/ConsultarDadosContribuinte.do?action'
                               '=VisualizarDadosContribuinte.do')
            if tipo:
                self.navegador.find_element(By.XPATH, '//*[@id="idTipoDocumento"]').click()
                self.navegador.find_element(By.XPATH, '/html/body/form/table/tbody/tr[1]/td[1]/select/option[4]').click()
            element = self.navegador.find_element(By.XPATH, '//*[@id="valor"]')
            element.send_keys(f'{text}')
            try:
                self.navegador.find_element(By.XPATH, '/html/body/form/table/tbody/tr[2]/td/input').click()
                self.navegador.find_element(By.XPATH, '//*[@id="item"]/tbody/tr/td[6]/a').click()
                nome = self.navegador.find_element(By.XPATH, '/html/body/table[3]/tbody/tr/td/fieldset[1]/table['
                                                             '1]/tbody/tr[4]/td/b')
                empresa = self.navegador.find_element(By.XPATH, '/html/body/table[3]/tbody/tr/td/fieldset[1]/table['
                                                                '1]/tbody/tr[6]/td')
                cidade = self.navegador.find_element(By.XPATH, '/html/body/table[3]/tbody/tr/td/fieldset[2]/table['
                                                               '2]/tbody/tr[2]/td[3]')
                for i in cidades:
                    if cidade.text == i:
                        data.append([nome.text, empresa.text, cidade.text, "Local"])
            except:
                print("CNPJ inválido.")

    def __init__(self, codes: list, cidades: list, brows: int, tipo: int):
        global ok
        ok = True
        try:
            if brows == 0:
                self.navegador = webdriver.Firefox()
            else:
                self.navegador = webdriver.Chrome()
            self.navegador.minimize_window()
            self.acao(codes, cidades, tipo)
            self.navegador.close()
        except:
            ok = False
            return


def wipedata():
    global data
    data = []


if __name__ == "__main__":
    __ = Execute
