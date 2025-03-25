import re
import pandas as pd
import requests
import os
import json
from datetime import datetime
import google.generativeai as genai

# Configuração do Google Generative AI
GOOGLE_API_KEY = "AIzaSyAbCjFYEaFmfNqEZxmWGJjCfD76XuzU7oY"
genai.configure(api_key=GOOGLE_API_KEY)

# Configuração otimizada do modelo Gemini
generation_config = {
    "temperature": 0.1,
    "top_p": 0.9,
    "top_k": 20,
    "max_output_tokens": 2048,
}

safety_settings = [
    {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_ONLY_HIGH"},
    {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_ONLY_HIGH"},
    {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_ONLY_HIGH"},
    {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_ONLY_HIGH"},
]

model = genai.GenerativeModel(
    model_name="gemini-1.0-pro",
    generation_config=generation_config,
    safety_settings=safety_settings,
)

def analisar_com_gemini(texto, prompt):
    """Função para enviar texto para análise pelo Gemini"""
    try:
        response = model.generate_content([prompt, texto])
        return response.text
    except Exception as e:
        print(f"Erro ao consultar Gemini: {e}")
        return None

def normalizar_telefone(telefone):
    """Normaliza números de telefone para formato (XX) XXXX-XXXX"""
    if not telefone:
        return "NÃO INFORMADO"
    
    # Remove tudo que não é dígito
    numeros = re.sub(r'[^\d]', '', telefone)
    
    # Remove código do país se existir (55)
    if numeros.startswith('55'):
        numeros = numeros[2:]
    
    # Formata o número
    if len(numeros) == 11:  # Com DDD e 9º dígito
        return f"({numeros[:2]}) {numeros[2:7]}-{numeros[7:]}"
    elif len(numeros) == 10:  # Com DDD sem 9º dígito
        return f"({numeros[:2]}) {numeros[2:6]}-{numeros[6:]}"
    elif len(numeros) == 9:  # Sem DDD com 9º dígito
        return f"(XX) {numeros[:5]}-{numeros[5:]}"
    elif len(numeros) == 8:  # Sem DDD e sem 9º dígito
        return f"(XX) {numeros[:4]}-{numeros[4:]}"
    else:
        return telefone  # Retorna original se não conseguir formatar

def normalizar_dados(dados):
    """Normaliza os dados extraídos para formato consistente"""
    # Normalização de CNPJ
    if 'CNPJ' in dados:
        dados['CNPJ'] = re.sub(r'[^0-9]', '', dados['CNPJ'])
        if len(dados['CNPJ']) == 14:
            dados['CNPJ'] = f"{dados['CNPJ'][:2]}.{dados['CNPJ'][2:5]}.{dados['CNPJ'][5:8]}/{dados['CNPJ'][8:12]}-{dados['CNPJ'][12:]}"
    
    # Normalização de telefone
    if 'Telefone' in dados:
        telefones = re.findall(r'[\d\(\)\s\-+]+', dados['Telefone'])
        if telefones:
            telefones_formatados = [normalizar_telefone(tel) for tel in telefones]
            dados['Telefone'] = ', '.join(telefones_formatados)
    
    # Separa número do endereço
    if 'Endereço' in dados:
        endereco = dados['Endereço']
        # Procura por padrões como "N° 123" ou ", 123"
        numero_match = re.search(r'(N°|n°|Nº|nº|,|;|)\s*(\d+)\b', endereco)
        if numero_match and 'Numero' not in dados:
            dados['Numero'] = numero_match.group(2)
            # Remove o número do endereço
            dados['Endereço'] = re.sub(r'(N°|n°|Nº|nº|,|;|)\s*\d+\b', '', endereco).strip()
    
    # Garantir que todos os campos necessários existam
    campos_necessarios = {
        "Quantos cartão": "NÃO INFORMADO",
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
    
    for campo, valor_padrao in campos_necessarios.items():
        if campo not in dados:
            dados[campo] = valor_padrao
    
    return dados

def extrair_dados_gemini(texto):
    """Usa Gemini para extrair dados de forma mais inteligente"""
    prompt = """
    ANALISE O TEXTO E EXTRAIA APENAS AS SEGUINTES INFORMAÇÕES:

    1. Quantos cartão: Número de cartões (se mencionado)
    2. CNPJ: No formato XX.XXX.XXX/XXXX-XX
    3. Nome Social: Razão social ou nome completo
    4. Endereço: Nome da rua/avenida (sem número)
    5. Numero: Número do endereço (extrair do endereço se necessário)
    6. Complemento: Ponto de referência ou informações adicionais
    7. Cidade: Nome da cidade
    8. UF: Sigla do estado (2 letras)
    9. Telefone: Qualquer formato de número de telefone (será normalizado)
    10. Email: Endereço de e-mail
    11. Vendedor: Nome do vendedor (se mencionado)

    RETORNE APENAS UM OBJETO JSON COM ESSES CAMPOS. USE "NÃO INFORMADO" PARA DADOS AUSENTES.
    """
    
    resposta_gemini = analisar_com_gemini(texto, prompt)
    if resposta_gemini:
        try:
            inicio_json = resposta_gemini.find('{')
            fim_json = resposta_gemini.rfind('}')
            if inicio_json != -1 and fim_json != -1:
                dados = json.loads(resposta_gemini[inicio_json:fim_json+1])
                return normalizar_dados(dados)
        except json.JSONDecodeError as e:
            print(f"Erro ao decodificar JSON do Gemini: {e}")
            print("Resposta recebida:", resposta_gemini)
    
    return extrair_dados_tradicional(texto)

def extrair_dados_tradicional(texto):
    """Método tradicional de extração por regex"""
    cnpj = re.search(r"CNPJ[:\s]*([\d./-]+)", texto, re.IGNORECASE)
    nome_social = re.search(r"(RAZ[ÃA]O SOCIAL|NOME SOCIAL|CEDENTE|GESTOR)[:\s]+([^\n]+)", texto, re.IGNORECASE)
    
    # Extrai endereço completo primeiro
    endereco_completo = re.search(r"(RUA|AVENIDA|RODOVIA|ENDEREÇO|LOGRADOURO)[\s:]*([^\n;]+)", texto, re.IGNORECASE)
    
    # Separa número do endereço
    endereco = numero = None
    if endereco_completo:
        endereco_texto = endereco_completo.group(2).strip()
        numero_match = re.search(r'(N°|n°|Nº|nº|,|;|)\s*(\d+)', endereco_texto)
        if numero_match:
            numero = numero_match.group(2)
            endereco = re.sub(r'(N°|n°|Nº|nº|,|;|)\s*\d+', '', endereco_texto).strip()
        else:
            endereco = endereco_texto
    
    complemento = re.search(r"(PONTO DE REFER[ÊE]NCIA|COMPLEMENTO|APT|APTO|SALA|LOJA|TERREO)[\s:]*([^\n;]+)", texto, re.IGNORECASE)
    cidade_uf = re.search(r"CIDADE[:\s]*([^-\n]+?)\s*-\s*([A-Z]{2})", texto, re.IGNORECASE)
    telefone = re.findall(r"(?:TELEFONE|CELULAR|CONTATO|WHATSAPP)[:\s]*([\(\)\d\s\-+]+)", texto, re.IGNORECASE) or \
              re.findall(r"[\d\(\)\s\-+]{10,}", texto)
    email = re.search(r"E-?MAIL[:\s]*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})", texto, re.IGNORECASE)
    vendedor = re.search(r"VENDEDOR[:\s]*([^\n]+)", texto, re.IGNORECASE)
    cartoes = re.search(r"(CART[ÃA]O|CARTÃO)[\s:]*(\d+)", texto, re.IGNORECASE)
    
    dados = {
        "Quantos cartão": cartoes.group(2) if cartoes else "NÃO INFORMADO",
        "CNPJ": cnpj.group(1) if cnpj else "NÃO INFORMADO",
        "Nome Social": nome_social.group(2).strip() if nome_social else "NÃO INFORMADO",
        "Endereço": endereco if endereco else "NÃO INFORMADO",
        "Numero": numero if numero else "NÃO INFORMADO",
        "Complemento": complemento.group(2).strip() if complemento else "NÃO INFORMADO",
        "Cidade": cidade_uf.group(1).strip() if cidade_uf else "NÃO INFORMADO",
        "UF": cidade_uf.group(2) if cidade_uf else "NÃO INFORMADO",
        "Telefone": ", ".join([tel.strip() for tel in telefone]) if telefone else "NÃO INFORMADO",
        "Email": email.group(1) if email else "NÃO INFORMADO",
        "Vendedor": vendedor.group(1).strip() if vendedor else "NÃO INFORMADO"
    }
    
    return normalizar_dados(dados)

def salvar_excel(dados, arquivo="dados_extraidos.xlsx"):
    """Salva os dados em uma planilha Excel, criando ou atualizando o arquivo"""
    try:
        # Ordem específica das colunas conforme solicitado
        colunas_ordenadas = [
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
        
        df_novo = pd.DataFrame([dados])
        
        # Garantir todas as colunas na ordem correta
        df_novo = df_novo[colunas_ordenadas]
        
        if os.path.exists(arquivo):
            df_existente = pd.read_excel(arquivo)
            # Garantir que o arquivo existente tenha as mesmas colunas
            for coluna in colunas_ordenadas:
                if coluna not in df_existente.columns:
                    df_existente[coluna] = "NÃO INFORMADO"
            df_existente = df_existente[colunas_ordenadas]
            
            df_final = pd.concat([df_existente, df_novo], ignore_index=True)
        else:
            df_final = df_novo
        
        df_final.to_excel(arquivo, index=False)
        print(f"Dados salvos em {arquivo}")
    except Exception as e:
        print(f"Erro ao salvar arquivo Excel: {e}")

if __name__ == "__main__":
    texto_exemplo = """
    
    """
    
    dados_extraidos = extrair_dados_gemini(texto_exemplo)
    print("\nDados extraídos:")
    for chave, valor in dados_extraidos.items():
        print(f"{chave}: {valor}")
    
    salvar_excel(dados_extraidos)