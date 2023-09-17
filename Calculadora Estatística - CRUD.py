########################################################
#EST: Análise de Pareto, Medidas de Tendência Central e Histograma
#Atividade avaliativa
#Disciplina de Estatistica Aplicada
#Programador : Rafael da Silva Brito
#Professor: Luiz Carlos dos Santos Filho
########################################################

#Importação de Bibliotecas
import os
import pandas as pd
import matplotlib.pyplot as plt

# Leia o DataFrame
df = pd.read_excel('C:/Users/Ryzen/Documents/GitHub/FATEC-EST/BD_erros.xlsx')

# Separar grupo por grupo de Erros/TítuloColuna
grupos = df.groupby('Erro')

# Contagem individual
contagem_individual = grupos.size().reset_index(name='Quantidade')

# Formatação em ordem decrescente
contagem_individual = contagem_individual.sort_values(by='Quantidade', ascending=False)

# Contagem Total de erros
total_erro = contagem_individual['Quantidade'].sum()

# Frequência em %
contagem_individual['Frequência(%)'] = (contagem_individual['Quantidade'] / total_erro * 100)

# Formatação da coluna Frequência(%) em duas casas decimais
contagem_individual['Frequência(%)'] = contagem_individual['Frequência(%)'].map('{:.2f}'.format)

# Calcular a frequência acumulada + formatação
frequencia_acumulada = contagem_individual['Frequência(%)'].astype(float).cumsum()
frequencia_acumulada = frequencia_acumulada.map('{:.2f}'.format)

# Arredonde a última frequência acumulada para 100%
frequencia_acumulada.iloc[-1] = '100.00'

# Atualize a coluna 'Frequência Acumulada(%) Arredondada' no DataFrame
contagem_individual['Frequência Acumulada(%)'] = frequencia_acumulada

# Criação do gráfico de Pareto
plt.figure(figsize=(10, 9))
ax = plt.gca()
bars = ax.bar(contagem_individual['Erro'], contagem_individual['Quantidade'], color='royalblue', alpha=0.7)
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
ax2.plot(contagem_individual['Erro'], contagem_individual['Frequência Acumulada(%)'], color='red', marker='o', markersize=8)
ax2.set_ylim(ax.get_ylim())
ax2.set_ylabel('Frequência Acumulada(%)', color='red')

#Apresenta a Análise de Pareto
plt.show()
print(contagem_individual[['Erro', 'Quantidade', 'Frequência(%)', 'Frequência Acumulada(%)']])
