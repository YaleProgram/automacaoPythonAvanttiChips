import json
import os
import re
import pandas as pd
import google.generativeai as genai
from datetime import datetime

# Configuração do Google Generative AI
GOOGLE_API_KEY = "AIzaSyAbCjFYEaFmfNqEZxmWGJjCfD76XuzU7oY"
genai.configure(api_key=GOOGLE_API_KEY)

# Configuração do modelo Gemini otimizada para extração de dados
generation_config = {
    "temperature": 0.1,  # Baixa temperatura para respostas mais precisas
    "top_p": 0.95,
    "top_k": 40,
    "max_output_tokens": 2048,
}

model = genai.GenerativeModel(
    model_name="gemini-1.0-pro",
    generation_config=generation_config,
)

def analisar_com_gemini(texto):
    """Função para enviar texto para análise pelo Gemini"""
    prompt = """
    EXTRAIA APENAS OS SEGUINTES DADOS DO TEXTO:

    1. Quantos cartão: (preencher com "1" se não encontrar informação)
    2. CNPJ: Formato XX.XXX.XXX/XXXX-XX
    3. Nome Social: Razão Social ou Nome do Cedente
    4. Endereço: Nome da rua (sem número/complemento)
    5. Numero: Número do endereço
    6. Complemento: Ponto de referência ou "SEM PONTO" se mencionado
    7. Cidade: Nome completo
    8. UF: Sigla com 2 letras
    9. Telefone: Todos os números de contato formatados como (XX) X XXXX-XXXX
    10. Email: Primeiro e-mail encontrado
    11. Vendedor: Nome do Gestor Master ou Vendedor

    RETORNE APENAS UM JSON COM ESSES 11 CAMPOS. USE "NÃO INFORMADO" SE NÃO ENCONTRAR.
    """
    
    try:
        response = model.generate_content([prompt, texto])
        if response.text:
            return response.text
    except Exception as e:
        print(f"Erro ao consultar Gemini: {e}")
    return None

def normalizar_telefone(telefone):
    """Normaliza qualquer formato de telefone para (XX) X XXXX-XXXX"""
    if not telefone:
        return "NÃO INFORMADO"
    
    # Remove tudo que não é dígito
    numeros = re.sub(r'[^\d]', '', telefone)
    
    # Remove código do país se existir
    if numeros.startswith('55'):
        numeros = numeros[2:]
    
    # Formata o número
    if len(numeros) == 11:  # Com DDD e 9º dígito
        return f"({numeros[:2]}) {numeros[2]} {numeros[3:7]}-{numeros[7:]}"
    elif len(numeros) == 10:  # Com DDD sem 9º dígito
        return f"({numeros[:2]}) {numeros[2:6]}-{numeros[6:]}"
    return telefone  # Retorna original se não conseguir formatar

def processar_dados(texto):
    """Processa o texto e retorna os dados formatados"""
    # Primeiro tenta com a IA
    dados_json = analisar_com_gemini(texto)
    
    if dados_json:
        try:
            # Extrai o JSON da resposta
            inicio = dados_json.find('{')
            fim = dados_json.rfind('}')
            if inicio != -1 and fim != -1:
                dados = json.loads(dados_json[inicio:fim+1])
                
                # Garante todos os campos necessários
                campos = {
                    "Quantos cartão": "1",
                    "CNPJ": "NÃO INFORMADO",
                    "Nome Social": "NÃO INFORMADO",
                    "Endereço": "NÃO INFORMADO",
                    "Numero": "NÃO INFORMADO",
                    "Complemento": "NÃO INFORMADO",
                    "Cidade": "NÃO INFORMADO",
                    "UF": "NÃO INFORMADO",
                    "Telefone": "NÃO INFORMADO",
                    "Email": "NÃO INFORMADO",
                    "Vendedor": "NÃO INFORMADO"
                }
                
                # Atualiza com os dados encontrados
                for campo in campos:
                    if campo in dados:
                        campos[campo] = dados[campo]
                
                # Normaliza telefones
                if "Telefone" in campos:
                    tels = re.findall(r'\(?\d{2}\)?[\s-]?\d[\s-]?\d{4}[\s-]?\d{4}', campos["Telefone"])
                    if tels:
                        campos["Telefone"] = ", ".join([normalizar_telefone(tel) for tel in tels])
                
                return campos
        except Exception as e:
            print(f"Erro ao processar JSON: {e}")
    
    # Fallback para extração tradicional se a IA falhar
    return extrair_dados_tradicional(texto)

def extrair_dados_tradicional(texto):
    """Método tradicional de extração por regex"""
    dados = {
        "Quantos cartão": "1",
        "CNPJ": re.search(r"CNPJ[:\s]*([\d./-]+)", texto, re.IGNORECASE).group(1) if re.search(r"CNPJ", texto, re.IGNORECASE) else "NÃO INFORMADO",
        "Nome Social": re.search(r"(RAZÃO SOCIAL|CEDENTE|GESTOR MASTER)[:\s]*([^\n]+)", texto, re.IGNORECASE).group(2).strip() if re.search(r"(RAZÃO SOCIAL|CEDENTE|GESTOR MASTER)", texto, re.IGNORECASE) else "NÃO INFORMADO",
        "Endereço": re.search(r"RUA[:\s]*([^\n;]+)", texto, re.IGNORECASE).group(1).strip() if re.search(r"RUA", texto, re.IGNORECASE) else "NÃO INFORMADO",
        "Numero": re.search(r"NÚMERO[:\s]*([^\n;]+)", texto, re.IGNORECASE).group(1).strip() if re.search(r"NÚMERO", texto, re.IGNORECASE) else "NÃO INFORMADO",
        "Complemento": re.search(r"PONTO DE REF[:\s]*([^\n;]+)", texto, re.IGNORECASE).group(1).strip() if re.search(r"PONTO DE REF", texto, re.IGNORECASE) else "NÃO INFORMADO",
        "Cidade": re.search(r"CIDADE[:\s]*([^\n-]+)", texto, re.IGNORECASE).group(1).strip() if re.search(r"CIDADE", texto, re.IGNORECASE) else "NÃO INFORMADO",
        "UF": re.search(r"ESTADO[:\s]*([A-Z]{2})", texto, re.IGNORECASE).group(1) if re.search(r"ESTADO", texto, re.IGNORECASE) else "NÃO INFORMADO",
        "Telefone": ", ".join([normalizar_telefone(tel) for tel in re.findall(r"(?:CONTATO|LINHA)[:\s]*([\(\)\d\s\-+]+)", texto)]),
        "Email": re.search(r"E-?MAIL[:\s]*([^\s]+@[^\s]+)", texto, re.IGNORECASE).group(1) if re.search(r"E-?MAIL", texto, re.IGNORECASE) else "NÃO INFORMADO",
        "Vendedor": re.search(r"(GESTOR MASTER|VENDEDOR)[:\s]*([^\n]+)", texto, re.IGNORECASE).group(2).strip() if re.search(r"(GESTOR MASTER|VENDEDOR)", texto, re.IGNORECASE) else "NÃO INFORMADO"
    }
    
    # Limpeza dos dados
    dados["Complemento"] = dados["Complemento"].replace(":", "").strip()
    if dados["Complemento"] == "":
        dados["Complemento"] = "SEM PONTO"
    
    return dados

def salvar_planilha(dados, arquivo="dados_extraidos.xlsx"):
    """Salva os dados na planilha no formato solicitado"""
    try:
        # Ordem das colunas exata como solicitado
        colunas = [
            "Quantos cartão",
            "CNPJ",
            "Nome Social",
            "Endereço",
            "Numero",
            "Complemento", 
            "Cidade",
            "UF",
            "Telefone",
            "Email",
            "Vendedor"
        ]
        
        df = pd.DataFrame([dados])[colunas]
        
        if os.path.exists(arquivo):
            df_existente = pd.read_excel(arquivo)
            df_final = pd.concat([df_existente, df], ignore_index=True)
        else:
            df_final = df
        
        df_final.to_excel(arquivo, index=False)
        print("Dados salvos com sucesso na planilha!")
        
    except Exception as e:
        print(f"Erro ao salvar planilha: {e}")

# Exemplo de uso com o texto fornecido
texto_exemplo = """
LIVIA 

VENDA CANTADA 

ENDEREÇO: 

RUA: FLORESTA
BAIRRO: SAO CAETANO
NÚMERO: 331 (EDIF CRISTAL APT 203)
CIDADE: ITABUNA
CEP: 45.607-090
PONTO DE REF: SEM PONTO 


RAZÃO SOCIAL:  J S B COMERCIO E REPRESENTACAO COMERCIAL LTDA
CNPJ: 56.868.552/0001-70

SITUAÇÃO CADASTRAL: ATIVO
IE: ISENTO 

GESTOR MASTER: EDUARDO JOSE SOARES BRANDAO
CPF: 553.130.865-53
RG: 4971520
E-MAL: EDUARDOJSB10@HOTMAIL.COM

GESTOR DE CONTA: EDUARDO JOSE SOARES BRANDAO
CONTATO: (73) 9 88248659
E-MAIL: EDUARDOJSB10@HOTMAIL.COM


CEDENTE :  EDUARDO JOSE SOARES BRANDAO
EMAIL: EDUARDOJSB10@HOTMAIL.COM
CPF: 553.130.865-53
RG: 4971520

CIDADE:  ITABUNA
ESTADO: BA

PORTABILIDADE PF/PJ

PLANO: 1L 100GB + 1l 1GB 
VALOR: 129,98 
LINHA: (73) 9 81584374 (1GB) 
VENCIMENTO: 06
OPERADORA: TIM

LINHA: (73) 9 88248659 (100GB)
VENCIMENTO: 06
OPERADORA: TIM
"""

dados = processar_dados(texto_exemplo)
print("Dados extraídos:")
for chave, valor in dados.items():
    print(f"{chave}: {valor}")

salvar_planilha(dados)