from selenium import webdriver as web
from selenium.webdriver.common.by import By
import openpyxl as excel
from tkinter import *
from tkinter import filedialog
import threading

#Projeto de automação
#com este código, entra-se no site do exemplo 
#É feita uma varredura do nome de todos os itens(jogos), e seus respectivos preços
#Após esse processo, todos os dados são organizados em uma tabela do excel

# Variável global para controlar a execução do scraping
executando_scraping = False
#Função para ser chamada pelo botão da interface
def web_scraping():
    global executando_scraping
    executando_scraping = True
    
    driver = web.Chrome() # Inicia-se a instância do Chrome Webdriver 
    driver.get('https://www.r10gamer.com.br/')
   
    #Faz a varredura dos dados
    nomes_jogos = driver.find_elements(By.XPATH,"//h2[@class='woocommerce-loop-product__title']")
    precos = driver.find_elements(By.XPATH,"//span[@class='original-price']")

    #Criação da Planinha
    planilha = excel.Workbook()
    planilha.create_sheet('Jogos') #Cria uma folha(sheet) específica para os dados: Jogos
    planilha_jogos = planilha['Jogos']
    planilha_jogos['A1'].value = 'Jogo'
    planilha_jogos['B1'].value = 'Preço'

    #Laço de repetição para varrer os dados e colocá-los na tabela 
    for jogos, preco in zip(nomes_jogos, precos):
        if not executando_scraping:
            break
        planilha_jogos.append([jogos.text,preco.text]) #Adiciona os dados na tabela 
        texto_jogos.insert(END,f'  Jogo: {jogos.text}  //  Preço: {preco.text}') #Coloca os dados na interface 
        texto_jogos.see(END) #Coloca os dados ao final da lista, atualizando-a
    
    #Esta parte do código serve para abrir o explorador de arquivos para que o usuário escolha oonde e como salvar
    arquivo = filedialog.asksaveasfilename(defaultextension=' .xlsx', filetypes=[('Planilha Excel', '*.xlsx')])
    planilha.save(arquivo)
    planilha.close()
    driver.quit()

    executando_scraping = False  # Reset da variável global ao finalizar

#Função usada para chamar a função Web_scraping. Foi a única maneira do código funcionar 
def inicar_scraping():
    threading.Thread(target=web_scraping).start()

def encerrar_scraping():
    global executando_scraping
    executando_scraping = False

#Criar o layout do app
janela = Tk()
janela.configure(bg='black')
janela.title('Web Scraping')
janela.geometry('600x600')
janela.grid_columnconfigure(0, weight=1)
janela.rowconfigure(3, weight=1)

#Elementos do layout
apresentacao = Label(janela, text='Automação de Coleta de Dados.\nClique no botão abaixo para iniciar.', font=('Helvetica', 20, 'bold'), bg='black', fg='white')
apresentacao.grid(column=0, row=0, padx=10, pady=10)

botao_style = ('Helvetica', 12, 'bold')
botao_coletar = Button(janela, text='Coletar Dados', command=inicar_scraping, font=botao_style, bg='#1E415D', fg='white', padx=20, pady=10) #Botão para chamar a função que contém o código
botao_coletar.grid(column=0, row=1, padx=10, pady=10)

botao_encerrar = Button(janela, text='Encerrar Coleta', command=encerrar_scraping, font=botao_style, bg='#FFD43B', fg='white', padx=20, pady=10)
botao_encerrar.grid(column=0, row=2, padx=10, pady=10)

texto_jogos = Listbox(janela, font=('Helvetica',12, 'bold'), bg='#1D1D1D', fg='white', selectbackground='white', selectforeground='black')
texto_jogos.grid(column=0, row=3, sticky='nsew', padx=10, pady=10)

nome_empresa = Label(janela, text='Desenvolvido por Front Dev Studio@', font=('Helvetica', 12, 'bold'), bg='black', fg='white')
nome_empresa.grid(column=0, row=4, padx=10, pady=10)

itens = texto_jogos.size()
if itens > 10:
    texto_jogos.config(height=10)
else:
    texto_jogos.config(height=itens)

janela.mainloop()