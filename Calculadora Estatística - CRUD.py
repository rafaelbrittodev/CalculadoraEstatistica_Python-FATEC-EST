# Importação de Bibliotecas
import os
import pandas as pd
import matplotlib.pyplot as plt

# Função para adicionar um novo erro ao DataFrame
def adicionar_erro(df):
    erro = input("Digite o nome do novo erro: ")
    quantidade = int(input("Digite a quantidade desse erro: "))
    
    # Adiciona o erro à lista de novas linhas
    novas_linhas = []
    for _ in range(quantidade):
        nova_linha = {'Erro': erro, 'Quantidade': 1}
        novas_linhas.append(nova_linha)
    
    # Concatena a lista de novas linhas ao DataFrame
    df = pd.concat([df, pd.DataFrame(novas_linhas)], ignore_index=True)
    
    return df

# Função para realizar a análise de Pareto
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

# Leia o DataFrame
df = pd.read_excel('C:/Users/Ryzen/Documents/GitHub/FATEC-EST/BD_erros.xlsx')

# Menu
while True:
    print("\nMenu:")
    print("1. Adicionar um novo erro")
    print("2. Exibir análise de Pareto")
    print("3. Sair e salvar Banco de Dados.")
    
    escolha = input("Escolha uma opção: ")
    
    if escolha == '1':
        df = adicionar_erro(df)
    elif escolha == '2':
        analise_pareto(df)  # Chamada da função para análise de Pareto
    elif escolha == '3':
        # Salve o DataFrame atualizado e saia do programa
        df.to_excel('C:/Users/Ryzen/Documents/GitHub/FATEC-EST/BD_erros.xlsx', index=False)
        break
    else:
        print("Opção inválida. Por favor, escolha uma opção válida.")

print("Programa encerrado.")
