import customtkinter as ctk
import pandas as pd
import matplotlib.pyplot as plt
from PIL import Image


# ------------------------------------------
# ---Criação das janelas e customização
# ------------------------------------------
# Criando a janela principal TKinter
janela = ctk.CTk()

janela.title('Calculadora de Estatística Aplicada - Spoke Studio')

# Set tamanho da janela
janela.geometry('1280x720')
janela.resizable(width=False, height=False)
# janela.iconify() #FECHA A JANELA
# janela.deiconify() #RECUPERA A JANELA FECHADA

# Customização do tema do app
janela._set_appearance_mode('system')

img1 = ctk.CTkImage(light_image=Image.open("./fotoicon.png"),
                    dark_image=Image.open("./fotoicon.png"), size=(180, 180))
ctk.CTkLabel(janela, text=None, image=img1).place(x=80, y=20)

# ------------------------------------------
# ---Funções
# ------------------------------------------


def nova_tela():
    nova_janela = ctk.CTkToplevel(janela)
    nova_janela.geometry('720x720')
    janela.resizable(width=False, height=False)


def importar_excel():
    dialog_import = ctk.CTkInputDialog(
        title='Importar Banco de Dados EXCEL', text='Digite o caminho do arquivo abaixo:')
    caminho = (dialog_import.get_input())
    df = pd.read_excel(f'{caminho}')


def add_erro():
    dialog_erro = ctk.CTkInputDialog(
        title='Adicionar Dados - Erro', text='Digite o nome do dado a adicionar ao banco de dados EXCEL')
    print(dialog_erro.get_input())
    erro = (f'{dialog_erro.get_input}')

    dialog_quantidade = ctk.CTkInputDialog(
        title='Adicionar Dados - Quantidade', text='Digite a quantidade do erro informado')
    print(dialog_quantidade.get_input())
    # Obtenha o valor da entrada de quantidade como uma string
    quantidade_str = (f'{dialog_quantidade.get_input()}')

    try:
        quantidade = int(quantidade_str)  # Converta a quantidade em um inteiro
    except ValueError:
        print("Erro: A quantidade deve ser um número inteiro válido.")
        return

    novas_linhas = []
    for _ in range(quantidade):
        nova_linha = {'Erro': erro, 'Quantidade': 1}
        novas_linhas.append(nova_linha)

    df = pd.read_excel(f'{caminho}')
    df = pd.concat([df, pd.DataFrame(novas_linhas)], ignore_index=True)

    return df

    # df = pd.read_excel('C:/Users/Ryzen/Documents/GitHub/FATEC-EST/BD_erros.xlsx')
# ------------------------------------------
# ---Frames
# ------------------------------------------
# frame1 = ctk.CTkFrame(master=janela, width=300, height=680).place(x=20, y=20)


# txtbox = ctk.CTkTextbox(master=janela, width=100, height=100)
# txtbox.pack()
# frame2 = ctk.CTkFrame(master=janela, width=920, height=680).place(x=340, y=20)
# ------------------------------------------
# ---Criação dos botões
# ------------------------------------------
# Label
ctk.CTkLabel(janela, text='Gerenciamento do Banco de Dados',
             font=('arial bold', 20)).place(x=20, y=200)
# Set buttons CRUD
btn_importBD = ctk.CTkButton(
    janela, text='Importar Banco de Dados EXCEL', width=315, height=30, command=importar_excel)
btn_importBD.place(x=20, y=240)

btn_addCrud = ctk.CTkButton(
    janela, text='Adicionar dado', width=315, height=30, command=add_erro)
btn_addCrud.place(x=20, y=280)

btn_removeCrud = ctk.CTkButton(
    janela, text='Excluir dado', width=315, height=30)
btn_removeCrud.place(x=20, y=320)

btn_altCrud = ctk.CTkButton(
    janela, text='Alterar dado', width=315, height=30)
btn_altCrud.place(x=20, y=360)

# Label
ctk.CTkLabel(janela, text='Análises e Medidas', font=(
    'arial bold', 20)).place(x=100, y=400)

# Set buttons pareto
btn_TbPareto = ctk.CTkButton(
    janela, text='Tabela de Distribuição de Frequência', width=315, height=30)
btn_TbPareto.place(x=20, y=440)

btn_diaPareto = ctk.CTkButton(
    janela, text='Diagrama de Análise de Pareto', width=315, height=30)
btn_diaPareto.place(x=20, y=480)

btn_medFrequencias = ctk.CTkButton(
    janela, text='Tabela de Medidas', width=315, height=30)
btn_medFrequencias.place(x=20, y=520)

btn_Histograma = ctk.CTkButton(janela, text='Histograma', width=315, height=30)
btn_Histograma.place(x=20, y=560)


# ------------------------------------------
# ---Executa a janela
# ------------------------------------------
janela.mainloop()
