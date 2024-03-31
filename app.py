from dash import Dash, html, dcc
import plotly.graph_objects as go
import pandas as pd
import requests
from io import BytesIO

# Inicialização da aplicação Dash
app = Dash(__name__)

todas_planilhas = pd.read_excel('data/questions_sheet.xlsx', sheet_name=None)

# # URL de onde o arquivo Excel será obtido (substitua pela URL correta)
# url_excel = 'http://localhost:3001/arquivo_excel'

# # Fazendo a requisição GET para obter o conteúdo do arquivo Excel
# response = requests.get(url_excel)

# # Verifica se a requisição foi bem-sucedida
# if response.status_code == 200:
#     # Lendo o conteúdo da resposta como um arquivo Excel
#     todas_planilhas = pd.read_excel(BytesIO(response.content), sheet_name=None)
# else:
#     print("Erro ao obter o arquivo Excel.")


# Definição das planilhas a serem ignoradas
planilhas_excluidas = ['root', 'links']

# Extração dos nomes das planilhas relevantes
nomes_planilhas_relevantes = [nome for nome in todas_planilhas.keys() if nome not in planilhas_excluidas]

# Verificação e seleção da coluna 'nome_da_questao' em cada DataFrame
listas_nomes_questoes = []

for nome_planilha, df_planilha in todas_planilhas.items():
    if nome_planilha not in planilhas_excluidas:
        if 'nome_da_questao' in df_planilha.columns:
            listas_nomes_questoes.append(df_planilha[['nome_da_questao']])
        else:
            print(f"A coluna 'nome_da_questao' não existe na planilha '{nome_planilha}'. Esta planilha será ignorada.")

# Concatenação dos DataFrames
df_questoes_unificado = pd.concat(listas_nomes_questoes, ignore_index=True)

# Remoção de linhas onde a coluna 'nome_da_questao' é "NAN"
df_questoes_unificado = df_questoes_unificado.dropna(subset=['nome_da_questao'])

# Cálculo da quantidade de questões de cada planilha
quantidades_questoes_por_planilha = []

for nome_planilha, df_planilha in todas_planilhas.items():
    if nome_planilha not in planilhas_excluidas:
        if 'nome_da_questao' in df_planilha.columns:
            quantidades_questoes_por_planilha.append(df_planilha['nome_da_questao'].dropna().shape[0])
        else:
            quantidades_questoes_por_planilha.append(0)
            print(f"A coluna 'nome_da_questao' não existe na planilha '{nome_planilha}'. Esta planilha será ignorada.")

# Criação do gráfico de barras para visualizar a quantidade de questões por planilha
figura_questoes = go.Figure(data=go.Bar(x=nomes_planilhas_relevantes, y=quantidades_questoes_por_planilha))

# Atualização do layout do gráfico
figura_questoes.update_layout(
    title='Quantidade de Questões por Planilha',
    xaxis_title='Planilhas',
    yaxis_title='Quantidade de questões'
)

# Cálculo do total de questões
total_questoes = df_questoes_unificado.shape[0]

# Criação do layout da aplicação Dash
app.layout = html.Div([
    html.H1(children='Visualização de Dados'),
    
    # Componente para exibir o gráfico
    html.Div([
        dcc.Graph(figure=figura_questoes)
    ]),
    
    # Componente para exibir o total de questões
    html.Div([
        html.P(f'Total de Questões: {total_questoes}')
    ])
])

# Execução da aplicação Dash
if __name__ == '__main__':
    app.run_server(debug=True)
