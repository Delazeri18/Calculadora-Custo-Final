import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QMessageBox
import pandas as pd

# Função para calcular o preço do anúncio com base na margem final desejada
def calcular_preco_anuncio(margem_desejada):
    try:
        file_path = 'ANÁLISE_PRECIFICAÇÃO.xlsx'
        # Carregar o arquivo Excel
        df = pd.read_excel(file_path)

        # Verificar se a coluna 'CUSTO PRODUTO' existe no DataFrame
        if 'CUSTO PRODUTO' in df.columns:
            # Calcular os preços do anúncio para cada linha
            for index, row in df.iterrows():
                custo_produto = row['CUSTO PRODUTO']
                preco_anuncio = custo_produto  # Inicia com o custo do produto

                # Loop para encontrar o preço do anúncio com margem próxima de 10%
                while True:
                    comissao = 0.50 * preco_anuncio
                    imposto = 0.78 * preco_anuncio
                    logistica = 0.45 * preco_anuncio
                    frete = 80.22 if preco_anuncio > 79.77 else 17
                    lucro = preco_anuncio - (custo_produto + comissao + imposto + logistica + frete)
                    margem_atual = lucro / custo_produto

                    # Verificando se a margem atual está próxima da desejada
                    if abs(margem_atual - margem_desejada) <= 0.01: # Tolerância de 1%
                        break  # Sai do loop se a margem estiver próxima da desejada

                    # Incrementando o preço do anúncio para a próxima iteração
                    preco_anuncio += 0.01

                # Atualizar o valor na coluna 'PREÇO ANÚNCIO'
                df.at[index, 'PREÇO ANÚNCIO'] = preco_anuncio

            # Salvar o DataFrame de volta no arquivo Excel
            df.to_excel(file_path, index=False)
            QMessageBox.information(None, 'Concluído', f'Coluna "PREÇO ANÚNCIO" adicionada ao arquivo {file_path} com sucesso!')
        else:
            QMessageBox.critical(None, 'Erro', 'A coluna "CUSTO PRODUTO" não foi encontrada no arquivo Excel.')
    except Exception as e:
        QMessageBox.critical(None, 'Erro', f'Ocorreu um erro ao processar o arquivo Excel: {str(e)}')

# Função para lidar com o botão "Calcular e Gerar Excel"
def handle_calcular():
    try:
        margem_desejada = float(entry_margem.text().replace(',', '.'))
        calcular_preco_anuncio(margem_desejada)
    except ValueError:
        QMessageBox.critical(None, 'Erro', 'Por favor, insira um valor válido para a margem final.')

# Criar aplicação PyQt
app = QApplication(sys.argv)

# Criar janela principal
window = QWidget()
window.setWindowTitle('Calculadora de Preço do Anúncio')

# Criar widgets
label_margem = QLabel('Escolha a margem final desejada:', window)
label_margem.move(20, 20)

entry_margem = QLineEdit(window)
entry_margem.move(20, 45)

button_calcular = QPushButton('Calcular e Gerar Excel', window)
button_calcular.move(20, 80)
button_calcular.clicked.connect(handle_calcular)

button_sair = QPushButton('Sair', window)
button_sair.move(200, 80)
button_sair.clicked.connect(app.quit)

# Ajustar tamanho da janela
window.resize(600, 120)

# Mostrar janela
window.show()

# Executar a aplicação
sys.exit(app.exec_())