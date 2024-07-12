from cnpj import CNPJClient, CNPJError
from datetime import datetime
import re


def consultar_cnpj(cnpj):
    cnpj = re.sub(r'[\.\-\/]', '', cnpj)

    if not validar_cnpj(cnpj):
        return "CNPJ inválido"

    cnpj_client = CNPJClient()

    try:
        dados_cnpj = cnpj_client.cnpj(cnpj)
        return dados_cnpj
    except CNPJError as erro:
        return erro

def validar_cnpj(cnpj):
    if len(cnpj) != 14:
        return False
    
    # Pesos para o cálculo dos dígitos verificadores
    peso1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    peso2 = [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]

    # Calcula o primeiro dígito verificador
    digito1 = calcular_digito(cnpj, peso1)
    if cnpj[12] != digito1:
        return False
    
    # Calcula o segundo dígito verificador
    digito2 = calcular_digito(cnpj, peso2)
    if cnpj[13] != digito2:
        return False
    
    return True

def calcular_digito(cnpj, peso):
    soma = 0
    for i in range(len(peso)):
        soma += int(cnpj[i]) * peso[i]
    resto = soma % 11
    return '0' if resto < 2 else str(11 - resto)

def processar_consulta(result):
    if isinstance(result, str) and result == "CNPJ inválido":
        return "CNPJ inválido"
    elif isinstance(result, Exception):
        return "Consulta Não Realizada"
    else:
        return "Consulta Realizada"

from datetime import datetime

def converter_data(data_str):
    try:
        return datetime.strptime(data_str, "%Y-%m-%d").date()
    except ValueError:
        return None

def extrair_dados(result):
    if isinstance(result, str) or isinstance(result, Exception):
        return {}

    data_inicio_atividade = converter_data(result.get("data_inicio_atividade"))
    data_situacao_cadastral = converter_data(result.get("data_situacao_cadastral"))

    dados = {
        "Razão Social": result.get("razao_social"),
        "Nome Fantasia": result.get("nome_fantasia"),
        "Natureza Jurídica": result.get("natureza_juridica"),
        "Porte": result.get("porte"),
        "Data de início das atividades": data_inicio_atividade,
        "Situação Cadastral": result.get("situacao_cadastral"),
        "Data Situação Cadastral": data_situacao_cadastral,
        "Endereco Completo": f"{result['endereco']['tipo_logradouro']} {result['endereco']['logradouro']}, {result['endereco']['numero']}, {result['endereco']['bairro']}, {result['endereco']['municipio']}-{result['endereco']['uf']}, CEP {result['endereco']['cep']}",
        "CNAE Principal - Código": result["cnae_fiscal_principal"]["codigo"],
        "CNAE Principal - Nome": result["cnae_fiscal_principal"]["nome"]
    }
    return dados


def extrair_cnaes_secundarios(result):
    if isinstance(result, str) or isinstance(result, Exception) or "cnae_fiscal_secundaria" not in result:
        return []
    
    cnaes_secundarios = result["cnae_fiscal_secundaria"]
    return cnaes_secundarios
