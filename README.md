# selettoBuscador
Buscador de empresas via raíz de CNPJ ou CACEAL

- [x] Retorna o nome da empresa e local
- [x] Mais de uma empresa por vez
- [x] gerar uma planílha com as empresas
- [x] busca em um arquivo excel

| Cidades suportadas |       status       |
|:------------------:|:------------------:|
| Limoeiro de Anadia | :white_check_mark: |
|     Taquarana      |        :x:         |
|     Arapiraca      |        :x:         |
| Giral do Ponciano  |        :x:         |

### Pacotes utilizados:
```
$ pip list

Package                 Version
-------------------------------
openpyxl                3.0.10
pandas                  1.5.2
selenium                4.7.2
ttkbootstrap            1.10.0
webdriver-manager       3.8.5
XlsxWriter              3.0.4
```

### Como usar

- Antes de tudo, selecione as cidades que deseja buscar.
- Ao clicar em importar, selecione um arquivo no formato xlsx contendo, na primeira coluna, a raíz do CNPJ das empresas que deseja procurar
- Logo após, pressione o botão buscar para que o programa comece a fazer o processo, que pode ser observado no navegador que será aberto pelo programa
- Quando o programa finalizar, abrirá uma janela de exportação para escolher onde e o nome que vai ser salvo o arquivo gerado.


>Observação, para utilizar este programa é necessário instalar um driver de automação.

|downloads|
|---|
|[<img src="https://pnggrid.com/wp-content/uploads/2021/04/Google-Chrome-Logo-2048x2048.png" width="20" /> Chrome](https://chromedriver.storage.googleapis.com/index.html?path=108.0.5359.71/)|
|[<img src="https://th.bing.com/th/id/R.27d319b45926552180640e6c91290e5e?rik=AAxCF1FvPM3t4Q&pid=ImgRaw&r=0" width="20" height="20" /> Firefox](https://github.com/mozilla/geckodriver/releases/tag/v0.32.0)|

>Muito simples de se usar, mas qualquer dúvida, pode me mandar nos issues
