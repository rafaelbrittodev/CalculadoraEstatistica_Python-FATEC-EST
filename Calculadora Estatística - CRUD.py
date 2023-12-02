# Importação de Bibliotecas
import os
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import datetime
from collections import Counter
import math


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

    # Criação do gráfico de Box Plot
    plt.figure(figsize=(12, 9))
    ax = plt.gca()

    # Adiciona um Box Plot no eixo Y
    ax.boxplot([contagem_individual['Quantidade']],
               vert=False,
               widths=0.5,
               showfliers=True,
               patch_artist=True,
               boxprops=dict(facecolor='lightblue', color='black'),
               capprops=dict(color='black'),
               whiskerprops=dict(color='black'),
               medianprops=dict(color='red'),
               flierprops=dict(markerfacecolor='red', marker='o', markersize=8)
               )

    # Configura o eixo Y para os erros
    ax.set_yticks([1])
    ax.set_yticklabels(['Erros'])
    ax.invert_yaxis()

    # Adiciona rótulos nas caixas com os valores
    for i, value in enumerate(contagem_individual['Quantidade']):
        ax.text(value, 1, f'{value}', va='center')

    # Configurações adicionais
    ax.set_xlim([0, contagem_individual['Quantidade'].max() + 1])
    plt.title('Box Plot dos Erros')
    plt.xlabel('Quantidade')

    # Apresenta o Box Plot
    plt.show()
    print(contagem_individual[['Erro', 'Quantidade']])


def histograma(df):
    # Agrupa por erro e calcula a soma da quantidade
    dados_agrupados = df.groupby('Erro')['Quantidade'].sum().reset_index()

    # Criação do histograma para cada tipo de erro
    plt.figure(figsize=(12, 8))
    ax = plt.gca()

    # Verifica se a coluna 'Erro' é categórica antes de criar o histograma
    if pd.api.types.is_categorical_dtype(dados_agrupados['Erro']):
        sns.barplot(data=dados_agrupados, x='Erro',
                    y='Quantidade', palette='viridis')
    else:
        sns.barplot(data=dados_agrupados, x='Erro',
                    y='Quantidade', palette='viridis')

    # Configurações adicionais
    plt.title('Histograma de Quantidade de Erros')
    plt.xlabel('Erro')
    plt.ylabel('Frequência')
    plt.xticks(rotation=45, ha='right')  # Rotaciona os rótulos do eixo X
    plt.grid(axis='y', linestyle='--', alpha=0.7)

    ax.set_axisbelow(True)

    # Apresenta o histograma
    plt.show()


def medida_de_tendencia_central(df):
    # Agrupa o DataFrame pelo nome do erro e soma as quantidades
    erros_agrupados = df.groupby('Erro')['Quantidade'].sum()

    # Converte a série agrupada de volta para um DataFrame
    df_agrupado = erros_agrupados.reset_index()

    # Calcula a média, mediana e moda
    media = df_agrupado['Quantidade'].mean()
    mediana = df_agrupado['Quantidade'].median()

    # Encontra a moda usando a biblioteca Counter para lidar com possíveis multimodalidades
    moda_contador = Counter(df_agrupado['Quantidade'])
    moda_quantidade = moda_contador.most_common(1)[0][0]
    moda = df_agrupado[df_agrupado['Quantidade']
                       == moda_quantidade]['Erro'].values

    # Cálculo dos quartis e desvio padrão
    quartis = df_agrupado['Quantidade'].quantile([0.25, 0.5, 0.75])
    desvio_padrao = df_agrupado['Quantidade'].std()

    # Criação da Tabela de Distribuição de Frequência
    tabela_frequencia = pd.DataFrame({
        'Erro': df_agrupado['Erro'],
        'Quantidade': df_agrupado['Quantidade']
    })

    # Adiciona colunas
    tabela_frequencia['Média'] = media
    tabela_frequencia['Mediana'] = mediana
    tabela_frequencia['Moda'] = str(moda)
    tabela_frequencia['Q1'] = quartis[0.25]
    tabela_frequencia['Q2 (Mediana)'] = quartis[0.5]
    tabela_frequencia['Q3'] = quartis[0.75]
    tabela_frequencia['Desvio Padrão'] = desvio_padrao

    # Exibe a Tabela de Distribuição de Frequência
    print('\nTabela de Distribuição de Frequência com Estatísticas:')
    print(tabela_frequencia)


def calculadora_binomial(df):
    print("\n*** Calculadora Probabilidade Binomial ***")

    # Verifica se o DataFrame está vazio
    if df.empty:
        print("O banco de dados está vazio. Adicione erros antes de usar a calculadora binomial.")
        return

    # Obtém a frequência relativa acumulada do DataFrame
    df['Frequência Acumulada(%)'] = df['Quantidade'].cumsum(
    ) / df['Quantidade'].sum() * 100

    # Solicita ao usuário o número de amostras observadas (n)
    try:
        n = int(input("Digite o número de amostras observadas (n): "))
    except ValueError:
        print("Entrada inválida. Certifique-se de digitar um número inteiro.")
        return

    # Solicita ao usuário a quantidade de vezes que deseja observar a probabilidade
    try:
        qtd_observacoes = int(
            input("Digite a quantidade de vezes que deseja observar a probabilidade: "))
    except ValueError:
        print("Entrada inválida. Certifique-se de digitar um número inteiro.")
        return

    # Calcula a probabilidade binomial
    probabilidade_binomial = {}
    for i, erro in enumerate(df['Erro']):
        p = df.at[i, 'Frequência Acumulada(%)'] / 100
        prob_binomial = math.comb(
            n, qtd_observacoes) * (p ** qtd_observacoes) * ((1 - p) ** (n - qtd_observacoes))
        probabilidade_binomial[erro] = probabilidade_binomial.get(
            erro, 0) + prob_binomial

    # Imprime os resultados
    print("\nProbabilidade Binomial para cada tipo de erro:")
    for erro, prob in probabilidade_binomial.items():
        print(f"{erro}: {prob:.4f}")


def login():
    # Número máximo de tentativas de login
    max_tentativas = 3
    tentativas = 0

    while tentativas < max_tentativas:
        try:
            # Solicitação de entrada do usuário
            usuario = input("\nDigite o nome de usuário: ")
            senha = input("Digite a senha: ")

            # Verifica as credenciais
            if usuario == 'admin' and senha == 'admin':
                print("Login bem-sucedido!")
                # Chama o Menu_Principal se o login for aprovado
                Menu_Principal()
                return

            else:
                raise ValueError("Credenciais inválidas. Tente novamente.")

        except ValueError as e:
            # Captura a exceção e imprime a mensagem de erro
            print(f"\nErro 404 {e}")
            tentativas += 1
            print(f"Tentativas restantes: {max_tentativas - tentativas}")

    print("Número máximo de tentativas atingido. Bloqueando o acesso.\n")
    # Encerra o programa se o número máximo de tentativas for atingido
    return

# Menu


def Menu_Principal():
    try:
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
            print('[7]  Exibir Box Plot')
            print('[8]  Exibir Histograma')
            print('[9]  Medidas de Tendência Central')
            print('[10] Calculadora Probabilidade Binomial')
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

            elif escolha == '7':
                boxplot_pareto(df)

            elif escolha == '8':
                histograma(df)

            elif escolha == '9':
                medida_de_tendencia_central(df)

            elif escolha == '10':
                calculadora_binomial(df)

            elif escolha == '11':
                print('\nBanco de dados salvo com sucesso!')
                caminho_arquivo = salvar_banco_de_dados(df, 'BD_erros')
                print(caminho_arquivo)

            elif escolha == '12':
                print("\nFinalizando aplicação...")
                break
            else:
                print("Opção inválida. Por favor, escolha uma opção válida.")

    except Exception as e:
        print(f"Ocorreu um erro: {e}")

    print("Programa encerrado.\n")


# Executar programa
login()
