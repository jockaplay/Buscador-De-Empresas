from selenium import webdriver
from selenium.webdriver.common.by import By

data = list()


class Execute:
    def acao(self, codes: list, cidades: list, tipodebusca: int):
        result = list()
        for text in codes:
            # 0 == CNPJ
            # 1 == CACEAL
            match tipodebusca:
                case 0:
                    for cidade in cidades:
                        resultado = self.searchCNPJ(text)
                        if cidade == resultado[2]:
                            result.append(resultado)
                case 1:
                    for cidade in cidades:
                        resultado = self.searchCACEAL(text)
                        if cidade == resultado[2]:
                            result.append(resultado)
        return result

    def searchCNPJ(self, text):
        nome, rasaosocial, local, contato1 = "Nome", "Razão social", "Cidade", "Email"
        try:
            self.navegador.get(
                f'https://cadsinc.sefaz.al.gov.br/VisualizarDadosContribuinte.do?opcao=raizcnpj&valor={text}')
            self.navegador.find_element(By.XPATH, '/html/body/div[2]/table/tbody/tr/td[6]/a').click()
            try:
                nome = self.navegador.find_element(By.XPATH,
                                                   '/html/body/table[3]/tbody/tr/td/fieldset[1]/table[1]/tbody/tr[4]/td/b').text
            except:
                pass
            try:
                rasaosocial = self.navegador.find_element(By.XPATH,
                                                          '/html/body/table[3]/tbody/tr/td/fieldset[1]/table[1]/tbody/tr[6]/td').text
            except:
                pass
            try:
                local = self.navegador.find_element(By.XPATH,
                                                    '/html/body/table[3]/tbody/tr/td/fieldset[2]/table[2]/tbody/tr[2]/td[3]').text
            except:
                pass
            try:
                contato1 = self.navegador.find_element(By.XPATH,
                                                       '/html/body/table[3]/tbody/tr/td/fieldset[2]/table[3]/tbody/tr[2]/td[3]').text
            except:
                pass
        except:
            pass
        return [nome, rasaosocial, local, contato1, text]

    def searchCACEAL(self, text):
        nome, rasaosocial, local, contato1 = "Nome", "Razão social", "Cidade", "Contato"
        try:
            self.navegador.get(f'http://cadsinc.sefaz.al.gov.br/VisualizarDadosContribuinte.do?opcao=caceal&valor={text}')
            try:
                nome = self.navegador.find_element(By.XPATH, '/html/body/table[3]/tbody/tr/td/fieldset[1]/table[1]/tbody/tr[4]/td/b').text
            except:
                pass
            try:
                rasaosocial = self.navegador.find_element(By.XPATH, '/html/body/table[3]/tbody/tr/td/fieldset[1]/table[1]/tbody/tr[6]/td').text
            except:
                pass
            try:
                local = self.navegador.find_element(By.XPATH, '/html/body/table[3]/tbody/tr/td/fieldset[2]/table[2]/tbody/tr[2]/td[3]').text
            except:
                pass
            try:
                contato1 = self.navegador.find_element(By.XPATH, '/html/body/table[3]/tbody/tr/td/fieldset[2]/table[3]/tbody/tr[2]/td[3]').text
            except:
                pass
        except:
            pass
        return [nome, rasaosocial, local, contato1, text]

    def __init__(self, codes: list, cidades: list, witchNav: int, tipo: int):
        global data
        try:
            if witchNav == 0:
                self.navegador = webdriver.Firefox()
            else:
                self.navegador = webdriver.Chrome()
            self.navegador.minimize_window()
            data = self.acao(codes, cidades, tipo)
            self.navegador.close()
        except:
            pass


def wipedata():
    global data
    data = []


if __name__ == "__main__":
    __ = Execute
