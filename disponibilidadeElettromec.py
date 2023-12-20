import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

def calcular_data_chegada(previsao, data_atual):
    quinzena = 15
    meses_previsao = {
       "JANEIRO": 1, "FEVEREIRO": 2, "MARÇO": 3, "ABRIL": 4, "MAIO": 5, "JUNHO": 6,
        "JULHO": 7, "AGOSTO": 8, "SETEMBRO": 9, "OUTUBRO": 10, "NOVEMBRO": 11, "DEZEMBRO": 12
    }

    if previsao:
        partes = previsao.split(' ')
        if len(partes) == 3:
            periodo_previsao, mes_previsao = partes[0], partes[2]
            mes_numero = meses_previsao.get(mes_previsao.upper(), 0)

            if mes_numero:
                diferenca_meses = mes_numero - data_atual.month
                diferenca_meses = diferenca_meses + 12 if diferenca_meses < 0 else diferenca_meses

                dias_ate_mes_previsao = diferenca_meses * 30
                dias_ate_mes_previsao += quinzena if periodo_previsao == "2ª" else 0

                total_dias = dias_ate_mes_previsao + quinzena
                return total_dias + 30
    return None
# Carregar as planilhas do Excel
kouzinaTray = pd.read_excel("KouzinaElettromec20_12.xlsx")
posicaoEstoque = pd.read_excel("posicaoEstoque.xlsx")

data_atual = datetime.now()
modelos_nao_identificados = []
sugestoes_alteracao = []
alertas_revisao = []
produtos_sem_alteracao = []
produtos_disponibilidade_zero = []
produtos_disponibilidade_imediata = []

# Contadores
total_sem_alteracao = 0
total_precisa_alteracao = 0
total_analisados = 0  # Contador para o total de produtos analisados

for index, row in kouzinaTray.iterrows():
    total_analisados += 1
    modelo_kouzina = row['Modelo']
    disponibilidade_kouzina = str(row['Disponibilidade'])

    correspondente = posicaoEstoque[posicaoEstoque['Modelo'] == modelo_kouzina]
    if not correspondente.empty:
        disponibilidade_estoque = correspondente.iloc[0]['Disponibilidade']
        previsao_chegada = correspondente.iloc[0]['PREVISÃO DE CHEGADA']

        if disponibilidade_estoque in ["INDISPONIVEL", "SOB CONSULTA"]:
            dias_para_chegada = calcular_data_chegada(previsao_chegada, data_atual) if previsao_chegada else None
            if "dias úteis" in disponibilidade_kouzina:
                try:
                    dias_uteis_kouzina = int(disponibilidade_kouzina.split(' ')[2])
                except ValueError:
                    dias_uteis_kouzina = None

                if dias_uteis_kouzina is not None and dias_uteis_kouzina != dias_para_chegada:
                    total_precisa_alteracao += 1
                    alertas_revisao.append(f"{modelo_kouzina} ({disponibilidade_estoque}): Alterar de {dias_uteis_kouzina} dias para {dias_para_chegada} dias")

        elif disponibilidade_kouzina not in ["0", "Imediata", ""] and "dias úteis" in disponibilidade_kouzina:
            try:
                dias_uteis_kouzina = int(disponibilidade_kouzina.split(' ')[2])
            except ValueError:
                dias_uteis_kouzina = None

            if dias_uteis_kouzina is not None:
                produtos_sem_alteracao.append(modelo_kouzina)
                total_sem_alteracao += 1

        # Adicionar produtos com disponibilidade "0" ou "Imediato" nas listas correspondentes
        if disponibilidade_kouzina == "0":
            produtos_disponibilidade_zero.append(modelo_kouzina)
        elif disponibilidade_kouzina.lower() == "imediata":
            produtos_disponibilidade_imediata.append(modelo_kouzina)

    else:
        modelos_nao_identificados.append(modelo_kouzina)



# Ordenando as listas
sugestoes_alteracao.sort()
alertas_revisao.sort()

# Escrever no arquivo de log
with open('relatorio_disponibilidade.log', 'w', encoding='utf-8') as log_file:
    log_file.write(f"Relatório de Disponibilidade - {data_atual.strftime('%d/%m/%Y %H:%M:%S')}\n")
    log_file.write(f"Total de produtos analisados: {total_analisados}\n")
    log_file.write(f"Total de produtos sem necessidade de alteração: {total_sem_alteracao}\n")
    log_file.write(f"Total de produtos que precisam de alteração: {total_precisa_alteracao}\n")
    log_file.write(f"Total de modelos não identificados na kouzinaTray: {len(modelos_nao_identificados)}\n")
    log_file.write("\nModelos não identificados na kouzinaTray:\n")
    for modelo in modelos_nao_identificados:
        log_file.write(f"- {modelo}\n")
        
    log_file.write("\nSugestões de Alteração:\n")
    for modelo in sugestoes_alteracao:
        log_file.write(f"- {modelo}\n")
        
    log_file.write("\nAlertas de Revisão:\n")
    for alerta in alertas_revisao:
        log_file.write(f"- {alerta}\n")
        
    log_file.write("\nProdutos sem necessidade de alteração:\n")
    for produto in produtos_sem_alteracao:
        log_file.write(f"- {produto}\n")
        
    log_file.write("\nProdutos com disponibilidade '0':\n")
    for produto in produtos_disponibilidade_zero:
        log_file.write(f"- {produto}\n")

    # Escrever a lista de produtos com disponibilidade "Imediato"
    log_file.write("\nProdutos com disponibilidade 'Imediata':\n")
    for produto in produtos_disponibilidade_imediata:
        log_file.write(f"- {produto}\n")

    # Escrever o total de produtos analisados
    log_file.write(f"\nTotal de produtos analisados: {total_analisados}\n")
        

print("Relatório gerado: relatorio_disponibilidade.log")
