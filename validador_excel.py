import pandas as pd
import re
from datetime import datetime

# SUPORTE 

def validar_cpf(cpf):
    cpf = re.sub(r'\D', '', str(cpf))
    if len(cpf) != 11 or cpf == cpf[0] * 11:
        return False
    for i in range(9, 11):
        soma = sum(int(cpf[num]) * ((i + 1) - num) for num in range(i))
        digito = (soma * 10 % 11) % 10
        if digito != int(cpf[i]):
            return False
    return True

def validar_data(data):
    try:
        dt = pd.to_datetime(data)
        return dt <= datetime.now()
    except:
        return False

def validar_uf(uf):
    ufs = ['AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 'MT', 'MS',
           'MG', 'PA', 'PB', 'PR', 'PE', 'PI', 'RJ', 'RN', 'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO']
    return str(uf).upper() in ufs

# Verificação

def executar_verificação(caminho):
    relatorio = []
    try:
        df = pd.read_excel(caminho)
        print(f"Arquivo '{caminho}' carregado com sucesso!")

        for idx, row in df.iterrows():
            erros = []

            # Identidade (CPF e Nome)
            if not validar_cpf(row.get('CPF')):
                erros.append("CPF Inválido")
            if pd.isna(row.get('Nome')):
                erros.append("Nome Ausente")

            # Contato (Email e Telefone)
            if not re.match(r"[^@]+@[^@]+\.[^@]+", str(row.get('Email'))):
                erros.append("E-mail Inválido")
            tel = re.sub(r'\D', '', str(row.get('Telefone')))
            if not (10 <= len(tel) <= 11):
                erros.append("Telefone Fora do Padrão")

            # Regras de Negócio (Data e Salário)
            if not validar_data(row.get('Data_Admissao')):
                erros.append("Data Inválida/Futura")
            try:
                sal_raw = str(row.get('Salario')).replace('R$', '').replace('.', '').replace(',', '.').strip()
                sal = float(sal_raw)
                if sal <= 0:
                    erros.append("Salário Negativo ou Zero")
            except:
                erros.append("Salário não numérico")

            # Geográfica (UF e CEP)
            if not validar_uf(row.get('UF')):
                erros.append("UF Inexistente")
            if len(re.sub(r'\D', '', str(row.get('CEP')))) != 8:
                erros.append("CEP Incorreto")

            if erros:
                relatorio.append({
                    "Linha_Excel": idx + 2,
                    "ID_Funcionario": row.get('ID', 'N/A'),
                    "Inconsistencias": " | ".join(erros)
                })

        print("Verificação de linhas concluída.")
        
        # Exportação
        
        if relatorio:
            pd.DataFrame(relatorio).to_excel("relatorio_verificação_final.xlsx", index=False)
            print(f"verificação finalizada: {len(relatorio)} registros com erro.")
        else:
            print("Base 100% íntegra.")

    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")

# CALL

caminho_alvo = "" 
executar_verificação(caminho_alvo)