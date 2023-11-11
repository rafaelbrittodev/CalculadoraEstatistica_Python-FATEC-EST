# Importação de Bibliotecas
import os
import pandas as pd
import matplotlib.pyplot as plt
import datetime

# Funções


def escolher_caminho_banco_de_dados():
    # Solicita ao usuário o caminho do banco de dados
    caminho_banco_de_dados = input("Digite o caminho do banco de dados: ")

    # Verifica se o caminho é válido
    if not os.path.exists(caminho_banco_de_dados):
        print("Caminho inválido. Digite novamente.")
        return escolher_caminho_banco_de_dados()

    # Verifica se é um caminho vazio
    if not os.listdir(caminho_banco_de_dados):
        print("Caminho inválido. O caminho deve apontar para um arquivo.")
        return escolher_caminho_banco_de_dados()

    # Verifica se o arquivo é um arquivo Excel
    if not caminho_banco_de_dados.endswith(".xlsx"):
        print("Arquivo inválido. O arquivo deve ser um arquivo Excel.")
        return escolher_caminho_banco_de_dados()

    return caminho_banco_de_dados


def adicionar_erro(df):
    erro = input("Digite o nome do novo erro: ")
    quantidade = input("Digite a quantidade desse erro: ")

    # Tenta converter a quantidade para um int
    try:
        quantidade = int(quantidade)
    except ValueError:
        # Se a conversão falhar, imprime uma mensagem de erro
        print("Quantidade inválida. Digite um número inteiro.")
        return adicionar_erro(df)

    # Adiciona o erro à lista de novas linhas
    novas_linhas = []
    for _ in range(quantidade):
        nova_linha = {'Erro': erro, 'Quantidade': 1}
        novas_linhas.append(nova_linha)

    # Concatena a lista de novas linhas ao DataFrame
    df = pd.concat([df, pd.DataFrame(novas_linhas)], ignore_index=True)

    return df


def excluir_erro(df):
    # Solicita ao usuário o erro a ser excluído
    erro_a_excluir = input("Digite o nome do erro a ser excluído: ")

    # Verifica se o erro existe no DataFrame
    if erro_a_excluir not in df['Erro'].to_list():
        print("Erro não encontrado. Digite novamente.")

        # Solicita novamente o nome do erro
        erro_a_excluir = input("Digite o nome do erro a ser excluído: ")

    # Remove o erro do DataFrame
    df = df[df['Erro'] != erro_a_excluir]

    return df


def alterar_erro(df):
    # Solicita ao usuário o erro a ser alterado
    erro_a_alterar = input("Digite o nome do erro a ser alterado: ")

    # Solicita ao usuário as novas informações do erro
    novo_nome = input("Digite o novo nome do erro: ")
    try:
        nova_quantidade = int(input("Digite a nova quantidade do erro: "))
    except ValueError:
        # Se a conversão falhar, imprime uma mensagem de erro
        print("Quantidade inválida. Digite um número inteiro.")
        return alterar_erro(df)

    # Verifica se o erro existe no DataFrame
    if erro_a_alterar not in df['Erro'].to_list():
        print("Erro não encontrado. Digite novamente.")
        return alterar_erro(df)

    # Altera as informações do erro no DataFrame
    df = df[df['Erro'] != erro_a_alterar]

    erro = novo_nome
    quantidade = nova_quantidade

    # Adiciona o erro à lista de novas linhas
    novas_linhas = []
    for _ in range(quantidade):
        nova_linha = {'Erro': erro, 'Quantidade': 1}
        novas_linhas.append(nova_linha)

    # Concatena a lista de novas linhas ao DataFrame
    df = pd.concat([df, pd.DataFrame(novas_linhas)], ignore_index=True)

    return df


def print_erros(df):
    # Ordena o DataFrame pelo número de ocorrências
    df = df.sort_values(by='Quantidade', ascending=False)

    # Soma a quantidade de erros por nome
    quantidades = df.groupby('Erro')['Quantidade'].sum()

    # Print dos nomes dos erros e suas respectivas quantidades
    for erro, quantidade in quantidades.items():
        print(f'\n{erro}: {quantidade}')


def salvar_banco_de_dados(df, nome_banco_de_dados):
    # Obtém a data e hora atuais
    agora = datetime.datetime.now()

    # Cria o nome do novo arquivo
    nome_arquivo = f"{nome_banco_de_dados}_{agora.strftime('%d_%m_%Y_%H_%M_%S')}.xlsx"

    # Salva o banco de dados no novo arquivo
    df.to_excel(nome_arquivo, index=False)

    return nome_arquivo


def verificar_banco_de_dados(df):
    # Verifique se o banco de dados foi informado
    if df is None:
        # Solicite ao usuário o caminho do banco de dados
        caminho_banco_de_dados = input("Digite o caminho do banco de dados: ")
        # Verifica se o caminho é válido
        if not os.path.exists(caminho_banco_de_dados):
            print("Caminho inválido. Digite novamente.")
            return escolher_caminho_banco_de_dados()

        # Verifica se é um caminho vazio
        if not os.listdir(caminho_banco_de_dados):
            print("Caminho inválido. O caminho deve apontar para um arquivo.")
            return escolher_caminho_banco_de_dados()

        # Verifica se o arquivo é um arquivo Excel
        if not caminho_banco_de_dados.endswith(".xlsx"):
            print("Arquivo inválido. O arquivo deve ser um arquivo Excel.")
            return escolher_caminho_banco_de_dados()
        # Verifique se o banco de dados foi informado
        if df is None:
            return None

        return df

        # Carrega o banco de dados
        df = pd.read_excel(caminho_banco_de_dados)

        # Verifica se o banco de dados está vazio
        if df.empty:
            print("Banco de dados vazio.")
            return None

    return df


def analise_pareto(df):
    # Separar grupo por grupo de Erros/TítuloColuna
    grupos = df.groupby('Erro')

    # Contagem individual
    contagem_individual = grupos.size().reset_index(name='Quantidade')

    # Formatação em ordem decrescente
    contagem_individual = contagem_individual.sort_values(
        by='Quantidade', ascending=False)

    # Calcular o total de erros
    total_erro = contagem_individual['Quantidade'].sum()

    # Frequência em %
    contagem_individual['Frequência(%)'] = (
        contagem_individual['Quantidade'] / total_erro * 100)

    # Formatação da coluna Frequência(%) em duas casas decimais
    contagem_individual['Frequência(%)'] = contagem_individual['Frequência(%)'].map(
        '{:.2f}'.format)

    # Calcular a frequência acumulada + formatação
    frequencia_acumulada = contagem_individual['Frequência(%)'].astype(
        float).cumsum()
    frequencia_acumulada = frequencia_acumulada.map('{:.2f}'.format)

    # Arredonde a última frequência acumulada para 100%
    frequencia_acumulada.iloc[-1] = '100.00'

    # Atualize a coluna 'Frequência Acumulada(%) Arredondada' no DataFrame
    contagem_individual['Frequência Acumulada(%)'] = frequencia_acumulada

    # Criação do gráfico de Pareto
    plt.figure(figsize=(10, 9))
    ax = plt.gca()
    bars = ax.bar(contagem_individual['Erro'],
                  contagem_individual['Quantidade'], color='royalblue', alpha=0.7)
    ax.set_ylabel('Quantidade', color='royalblue')
    plt.xticks(rotation=15)
    plt.title('Análise de Pareto')

    # Adicionar rótulos nas barras com os valores
    for bar in bars:
        height = bar.get_height()
        ax.annotate(f'{height}',
                    xy=(bar.get_x() + bar.get_width() / 2, height),
                    xytext=(0, 3),  # 3 pontos acima da barra
                    textcoords='offset points',
                    ha='center', va='bottom')

    # Configura o eixo Y direito para a frequência acumulada para maior visibilidade
    ax2 = ax.twinx()
    ax2.plot(contagem_individual['Erro'], contagem_individual['Frequência Acumulada(%)'],
             color='red', marker='o', markersize=8)
    ax2.set_ylim(ax.get_ylim())
    ax2.set_ylabel('Frequência Acumulada(%)', color='red')

    # Apresenta a Análise de Pareto
    plt.show()
    print(contagem_individual[['Erro', 'Quantidade',
                               'Frequência(%)', 'Frequência Acumulada(%)']])


def boxplot_pareto(df):

    # Separar grupo por grupo de Erros/TítuloColuna
    grupos = df.groupby('Erro')

    # Contagem individual
    contagem_individual = grupos.size().reset_index(name='Quantidade')

    # Formatação em ordem decrescente
    contagem_individual = contagem_individual.sort_values(
        by='Quantidade', ascending=False)

    # Calcular o total de erros
    total_erro = contagem_individual['Quantidade'].sum()

    # Frequência em %
    contagem_individual['Frequência(%)'] = (
        contagem_individual['Quantidade'] / total_erro * 100)

    # Formatação da coluna Frequência(%) em duas casas decimais
    contagem_individual['Frequência(%)'] = contagem_individual['Frequência(%)'].map(
        '{:.2f}'.format)

    # Calcular a frequência acumulada + formatação
    frequencia_acumulada = contagem_individual['Frequência(%)'].astype(
        float).cumsum()
    frequencia_acumulada = frequencia_acumulada.map('{:.2f}'.format)

    # Arredonde a última frequência acumulada para 100%
    frequencia_acumulada.iloc[-1] = '100.00'

    # Atualize a coluna 'Frequência Acumulada(%) Arredondada' no DataFrame
    contagem_individual['Frequência Acumulada(%)'] = frequencia_acumulada

    # Criação do gráfico de Pareto
    plt.figure(figsize=(10, 9))
    ax = plt.gca()
    bars = ax.bar(contagem_individual['Erro'],
                  contagem_individual['Quantidade'], color='royalblue', alpha=0.7)
    ax.set_ylabel('Quantidade', color='royalblue')
    plt.xticks(rotation=15)
    plt.title('Análise de Pareto')

    # Adicionar rótulos nas barras com os valores
    for bar in bars:
        height = bar.get_height()
        ax.annotate(f'{height}',
                    xy=(bar.get_x() + bar.get_width() / 2, height),
                    xytext=(0, 3),  # 3 pontos acima da barra
                    textcoords='offset points',
                    ha='center', va='bottom')

    # Configura o eixo Y direito para a frequência acumulada para maior visibilidade
    ax2 = ax.twinx()
    ax2.plot(contagem_individual['Erro'], contagem_individual['Frequência Acumulada(%)'],
             color='red', marker='o', markersize=8)
    ax2.set_ylim(ax.get_ylim())
    ax2.set_ylabel('Frequência Acumulada(%)', color='red')

    # Adiciona um Box Plot no eixo X
    ax2.boxplot(contagem_individual['Quantidade'],
                positions=contagem_individual['Erro'],
                vert=False,
                widths=0.2,
                showfliers=False,
                notch=True)

    # Apresenta a Análise de Pareto
    plt.show()


# Menu


def Menu_Principal():
    # Leia o DataFrame Padrão
    caminho_banco_de_dados = 'BD_erros.xlsx'
    df = pd.read_excel(caminho_banco_de_dados)
    while True:
        print('\n-------------------- Menu --------------------')
        print('[1]  Escolher banco de dados')
        print('[2]  Exibir banco de dados')
        print('[3]  Adicionar um novo erro')
        print('[4]  Alterar um erro existente no Banco de Dados')
        print('[5]  Excluir um erro existente no Banco de Dados')
        print('[6]  Exibir Análise de Pareto')
        print('[7]*  Exibir Box Plot')
        print('[8]*  Exibir Histograma')
        print('[9]*  Medidas de Tendência Central')
        print('[10]* Tabela de Distribuição de Frequência')
        print('[11] Salvar Banco de Dados.')
        print('[12] Sair')

        escolha = input("\nEscolha uma opção: ")

        if escolha == '1':
            caminho_banco_de_dados = escolher_caminho_banco_de_dados()
            df = pd.read_excel(caminho_banco_de_dados)
            print_erros(df)

        elif escolha == '2':
            verificar_banco_de_dados(df)
            print_erros(df)

        elif escolha == '3':
            df = adicionar_erro(df)
            print_erros(df)
        elif escolha == '4':
            df = alterar_erro(df)
            print_erros(df)

        elif escolha == '5':
            df = excluir_erro(df)
            print_erros(df)

        elif escolha == '6':
            analise_pareto(df)

        elif escolha == '11':
            print('\nBanco de dados salvo com sucesso!')
            caminho_arquivo = salvar_banco_de_dados(df, 'BD_erros')
            print(caminho_arquivo)

        elif escolha == '12':
            print("\nFinalizando aplicação...")
            break
        else:
            print("Opção inválida. Por favor, escolha uma opção válida.")

    print("Programa encerrado.\n")


# Executar programa
Menu_Principal()
