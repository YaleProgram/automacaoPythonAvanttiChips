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
    "temperature": 0.1,  # Mais baixo para extração precisa de dados
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

def corrigir_texto(texto):
    """Corrige erros gramaticais usando LanguageTool"""
    url = "https://api.languagetool.org/v2/check"
    params = {"text": texto, "language": "pt-BR"}
    
    try:
        response = requests.post(url, data=params).json()
        correcoes = []
        
        for match in response.get("matches", []):
            erro = match["context"]["text"][match["offset"]:match["offset"]+match["length"]]
            sugestao = match["replacements"][0]["value"] if match["replacements"] else erro
            correcoes.append((erro, sugestao))
        
        for erro, sugestao in correcoes:
            texto = texto.replace(erro, sugestao, 1)
        
        return texto
    except Exception as e:
        print(f"Erro ao corrigir texto: {e}")
        return texto

def normalizar_dados(dados):
    """Normaliza os dados extraídos para formato consistente"""
    # Normalização de CNPJ/CPF
    if 'CNPJ' in dados:
        dados['CNPJ'] = re.sub(r'[^0-9]', '', dados['CNPJ'])
        if len(dados['CNPJ']) == 11:  # É um CPF
            dados['CPF'] = dados['CNPJ']
            dados['CNPJ'] = "NÃO ENCONTRADO"
        elif len(dados['CNPJ']) == 14:
            dados['CNPJ'] = f"{dados['CNPJ'][:2]}.{dados['CNPJ'][2:5]}.{dados['CNPJ'][5:8]}/{dados['CNPJ'][8:12]}-{dados['CNPJ'][12:]}"
    
    # Normalização de CPF
    if 'CPF' in dados:
        dados['CPF'] = re.sub(r'[^0-9]', '', dados['CPF'])
        if len(dados['CPF']) == 11:
            dados['CPF'] = f"{dados['CPF'][:3]}.{dados['CPF'][3:6]}.{dados['CPF'][6:9]}-{dados['CPF'][9:]}"
    
    # Normalização de telefone
    for tel_field in ['TELEFONE', 'NUMERO DE TELEFONE', 'CONTATO']:
        if tel_field in dados:
            telefones = re.findall(r'\d{10,11}', re.sub(r'[^\d]', '', dados[tel_field]))
            if telefones:
                dados[tel_field] = ', '.join([f'({tel[:2]}) {tel[2:7]}-{tel[7:]}' if len(tel) == 11 
                                           else f'({tel[:2]}) {tel[2:6]}-{tel[6:]}' for tel in telefones])
                # Preenche ambos os campos de telefone
                dados['TELEFONE'] = dados[tel_field]
                dados['NUMERO DE TELEFONE'] = dados[tel_field]
    
    # Normalização de CEP
    if 'CEP' in dados:
        dados['CEP'] = re.sub(r'[^\d]', '', dados['CEP'])
        if len(dados['CEP']) == 8:
            dados['CEP'] = f"{dados['CEP'][:5]}-{dados['CEP'][5:]}"
    
    # Tratamento especial para endereço
    if 'RUA' in dados and 'ENDERECO' not in dados:
        dados['ENDERECO'] = dados['RUA']
        dados['ENDEREÇO'] = dados['RUA']
    
    # Tratamento especial para razão social
    for field in ['RAZAO_SOCIAL', 'RAZÃO SOCIAL', 'CEDENTE', 'GESTOR MASTER']:
        if field in dados and dados.get('RAZAO_SOCIAL') == "NÃO ENCONTRADO":
            dados['RAZAO_SOCIAL'] = dados[field]
            dados['RAZÃO SOCIAL'] = dados[field]
    
    # Garantir que todos os campos necessários existam
    campos_necessarios = {
        "BAIRRO": "NÃO ENCONTRADO",
        "CEP": "NÃO ENCONTRADO",
        "CIDADE": "NÃO ENCONTRADO",
        "CNPJ": "NÃO ENCONTRADO",
        "COMPLEMENTO": "NÃO ENCONTRADO",
        "CPF": "NÃO ENCONTRADO",
        "DATA_EXTRACAO": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "EMAIL": "NÃO ENCONTRADO",
        "ENDERECO": "NÃO ENCONTRADO",
        "ENDEREÇO": "NÃO ENCONTRADO",
        "ESTADO": "NÃO ENCONTRADO",
        "NUMERO DE TELEFONE": "NÃO ENCONTRADO",
        "RAZAO_SOCIAL": "NÃO ENCONTRADO",
        "RAZÃO SOCIAL": "NÃO ENCONTRADO",
        "TELEFONE": "NÃO ENCONTRADO"
    }
    
    # Preencher os campos faltantes
    for campo, valor_padrao in campos_necessarios.items():
        if campo not in dados:
            dados[campo] = valor_padrao
    
    return dados

def extrair_dados_gemini(texto):
    """Usa Gemini para extrair dados de forma mais inteligente"""
    prompt = """
    ANALISE O TEXTO E EXTRAIA AS SEGUINTES INFORMAÇÕES COM MUITA ATENÇÃO:

    1. BAIRRO: Extraia o nome do bairro após "BAIRRO;"
    2. CEP: Extraia o CEP no formato XXXXX-XXX
    3. CIDADE: Extraia o nome da cidade
    4. CNPJ: Extraia no formato XX.XXX.XXX/XXXX-XX
    5. COMPLEMENTO: Extraia informações após "TERREO", "CASA" ou similar
    6. CPF: Extraia no formato XXX.XXX.XXX-XX
    7. EMAIL: Todos os emails encontrados
    8. ENDEREÇO: Combine "RUA" com número e complemento
    9. ESTADO: Sigla de 2 letras
    10. TELEFONE: Todos os números no formato (XX) XXXX-XXXX ou (XX) XXXXX-XXXX
    11. RAZÃO SOCIAL: Pode estar em "RAZÃO SOCIAL", "CEDENTE" ou "GESTOR MASTER"

    RETORNE APENAS UM OBJETO JSON COM ESSES CAMPOS. SE NÃO ENCONTRAR, USE "NÃO ENCONTRADO".
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
    """Método tradicional de extração por regex com melhorias"""
    # Padrões melhorados para capturar os dados
    bairro = re.search(r"BAIRRO[;\s:]+([^\n;]+)", texto, re.IGNORECASE)
    cep = re.search(r"CEP[:\s]*(\d{2}\.?\d{3}-?\d{3})", texto, re.IGNORECASE)
    cidade_estado = re.search(r"CIDADE[:\s]*([^-\n]+?)\s*-\s*([A-Z]{2})", texto, re.IGNORECASE)
    cnpj = re.search(r"CNPJ[:\s]*([\d./-]+)", texto, re.IGNORECASE)
    complemento = re.search(r"(?:TERREO|CASA|APT|APTO|SALA|LOJA)[\s:]*([^\n;]+)", texto, re.IGNORECASE)
    cpf = re.search(r"CPF[:\s]*([\d.-]+)", texto, re.IGNORECASE)
    email = re.search(r"E-?MAIL[:\s]*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})", texto, re.IGNORECASE)
    endereco = re.search(r"(?:RUA|ENDEREÇO|LOGRADOURO)[;\s:]+([^\n;]+)", texto, re.IGNORECASE)
    telefone = re.findall(r"(?:CONTATO|TELEFONE|CELULAR)[:\s]*([\(\)\d\s-]+)", texto, re.IGNORECASE) or \
              re.findall(r"\(?\d{2}\)?\s?\d{4,5}-?\d{4}", texto)
    razao_social = re.search(r"(?:RAZ[ÃA]O SOCIAL|CEDENTE|GESTOR MASTER)[:\s]+([^\n]+)", texto, re.IGNORECASE)
    
    dados = {
        "BAIRRO": bairro.group(1).strip() if bairro else "NÃO ENCONTRADO",
        "CEP": cep.group(1).replace(".", "").replace("-", "")[:5] + "-" + cep.group(1).replace(".", "").replace("-", "")[5:] if cep else "NÃO ENCONTRADO",
        "CIDADE": cidade_estado.group(1).strip() if cidade_estado else "NÃO ENCONTRADO",
        "CNPJ": cnpj.group(1) if cnpj else "NÃO ENCONTRADO",
        "COMPLEMENTO": complemento.group(1).strip() if complemento else "NÃO ENCONTRADO",
        "CPF": cpf.group(1) if cpf else "NÃO ENCONTRADO",
        "EMAIL": email.group(1) if email else "NÃO ENCONTRADO",
        "ENDERECO": endereco.group(1).strip() if endereco else "NÃO ENCONTRADO",
        "ENDEREÇO": endereco.group(1).strip() if endereco else "NÃO ENCONTRADO",
        "ESTADO": cidade_estado.group(2) if cidade_estado else "NÃO ENCONTRADO",
        "TELEFONE": ", ".join([tel.strip() for tel in telefone]) if telefone else "NÃO ENCONTRADO",
        "NUMERO DE TELEFONE": ", ".join([tel.strip() for tel in telefone]) if telefone else "NÃO ENCONTRADO",
        "RAZAO_SOCIAL": razao_social.group(1).strip() if razao_social else "NÃO ENCONTRADO",
        "RAZÃO SOCIAL": razao_social.group(1).strip() if razao_social else "NÃO ENCONTRADO",
        "DATA_EXTRACAO": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    
    return normalizar_dados(dados)

def salvar_excel(dados, arquivo="dados_extraidos.xlsx"):
    """Salva os dados em uma planilha Excel, criando ou atualizando o arquivo"""
    try:
        colunas_ordenadas = [
            "BAIRRO", "CEP", "CIDADE", "CNPJ", "COMPLEMENTO", "CPF", "DATA_EXTRACAO",
            "EMAIL", "ENDERECO", "ENDEREÇO", "ESTADO", "NUMERO DE TELEFONE",
            "RAZAO_SOCIAL", "RAZÃO SOCIAL", "TELEFONE"
        ]
        
        df_novo = pd.DataFrame([dados])
        
        for coluna in colunas_ordenadas:
            if coluna not in df_novo.columns:
                df_novo[coluna] = "NÃO ENCONTRADO"
        
        df_novo = df_novo[colunas_ordenadas]
        
        if os.path.exists(arquivo):
            df_existente = pd.read_excel(arquivo)
            for coluna in colunas_ordenadas:
                if coluna not in df_existente.columns:
                    df_existente[coluna] = "NÃO ENCONTRADO"
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
    BOA TARDE 

    CAPANHA DDD 

    ENDEREÇO DIFERENTE DA RECEITA

    RUA; rua 10 , N° 23 , TERREO CASA
    BAIRRO; CAMPO DO GOVERNO
    CIDADE: ITABERABA - BA
    CEP: 46.880-000
    PONTO DE REFERENCIA : REFERENCIA : SUPERMECADO DO DENGO

    RAZÃO SOCIAL: 53.554.993 ROGERIO FIGUEREDO SOUZA
    CNPJ: 53.554.993/0001-00

    SITUAÇÃO CADASTRAL: 
    IE: INSENTO 

    GESTOR MASTER: ROGERIO FIGUEREDO SOUZA
    CPF: 83925406549
    RG: 1274391610
    E-MAIL:  figueredorogerio300@gmail.com

    GESTOR DE CONTA: ROGERIO FIGUEREDO SOUZA
    CONTATO: 75991779847
    E-MAIL figueredorogerio300@gmail.com

    CEDENTE : ROGERIO FIGUEREDO SOUZA
    EMAIL: figueredorogerio300@gmail.com
    CPF: 83925406549
    RG: 1274391610

    CIDADE: ITABERABA
    ESTADO: BA

    PORTABILIDADE PF - PJ

    PLANO: 1L 16GB R$ 39,99 
    LINHA: 75991779847
    VENCIMENTO: 10
    OPERADORA ATUAL: TIM
    """
    
    dados_extraidos = extrair_dados_gemini(texto_exemplo)
    print("\nDados extraídos:")
    for chave, valor in dados_extraidos.items():
        print(f"{chave}: {valor}")
    
    salvar_excel(dados_extraidos)