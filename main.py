import time
from selenium import webdriver
from bs4 import BeautifulSoup as bs
import openpyxl




navegador = webdriver.Chrome()
url = 'https://www.visaoimobiliariaparacatu.com.br/buscar?order=neighborhood&availability=buy&city=Paracatu'
navegador.get(url)
time.sleep(5)


book = openpyxl.Workbook()
book.create_sheet('Imoveis')
addImoveis = book['Imoveis']

cont = 1
n = 1

while True:

    try:
        nome = navegador.find_element('xpath', f'//*[@id="search"]/section[2]/imobzi-property-list/section/mat-card[{cont}]/mat-card-content/mat-card-title').text
        bairro = navegador.find_element('xpath', f'//*[@id="search"]/section[2]/imobzi-property-list/section/mat-card[{cont}]/mat-card-content/mat-card-subtitle[1]').text
        try:
            valor = navegador.find_element('css selector', f'#search > section.gd-xs-1-12.gd-sm-1-12.gd-md-5-12.gd-lg-4-12.gd-xl-4-12 > imobzi-property-list > section > mat-card:nth-child({cont}) > mat-card-content > mat-card-subtitle.mat-card-subtitle.h3.color-title.bold > div').text
        except:
            valor = 'Consulte'



        element = navegador.find_element('xpath', f'//*[@id="search"]/section[2]/imobzi-property-list/section/mat-card[{cont}]/mat-card-header/imobzi-property-gallery/div[1]/a')
        html_content = element.get_attribute('outerHTML')
        soup = bs(html_content, 'html.parser')

        atributo = {"class": "swiper-wrapper"}

        for url in soup.find_all("a", attrs=atributo):
            urlHref = url['href']
            link = f'https://www.visaoimobiliariaparacatu.com.br{urlHref}'

        print(f'========== {n} ==========')
        print(f'Nome: {nome}')
        print(f'Bairro: {bairro}')
        print(f'Valor: {valor}')
        print(f"Link: {link}\n\n")

        addImoveis.append([nome, bairro, valor, link])

        cont += 1
        n += 1


    except:
        try:

            try:
                button = navegador.find_element('css selector', '#search > section.gd-xs-1-12.gd-sm-1-12.gd-md-5-12.gd-lg-4-12.gd-xl-4-12 > imobzi-property-list > imobzi-pagination > section > button.btn.btn-md.btn-disabled-2').text
                print(button)
                if button == 'Próximo':
                    break
                else:
                    navegador.find_element('css selector', '#search > section.gd-xs-1-12.gd-sm-1-12.gd-md-5-12.gd-lg-4-12.gd-xl-4-12 > imobzi-property-list > imobzi-pagination > section > button:nth-child(3)').click()
                    print('PRÓXIMA PÁGINA')
                    time.sleep(10)
                    cont = 1
            except:
                navegador.find_element('css selector', '#search > section.gd-xs-1-12.gd-sm-1-12.gd-md-5-12.gd-lg-4-12.gd-xl-4-12 > imobzi-property-list > imobzi-pagination > section > button:nth-child(3)').click()
                print('PRÓXIMA PÁGINA')
                time.sleep(10)
                cont = 1



        except:
            print('ACABOU')
            break

book.save('Imoveis.xlsx')