import re
import pandas as pd
import google.generativeai as genai
import os
import json

# Configuração do Google Generative AI
GOOGLE_API_KEY = "AIzaSyAbCjFYEaFmfNqEZxmWGJjCfD76XuzU7oY"
genai.configure(api_key=GOOGLE_API_KEY)

# Configuração otimizada para extração precisa
generation_config = {
    "temperature": 0.1,
    "top_p": 0.95,
    "top_k": 20,
    "max_output_tokens": 1024,
}

model = genai.GenerativeModel(
    model_name="gemini-1.0-pro",
    generation_config=generation_config,
)

def extrair_telefone_principal(texto):
    """Extrai apenas o primeiro telefone de CONTATO, ignorando LINHAS de plano"""
    # Primeiro procura por CONTATO
    contato = re.search(r"CONTATO[:\s]*([\(\)\d\s\-+]+)", texto, re.IGNORECASE)
    if contato:
        return formatar_telefone(contato.group(1))
    
    # Se não encontrar, procura o primeiro telefone válido (mas ignora LINHA:)
    telefones = re.findall(r"(?<!LINHA[:\s])(\(?\d{2}\)?[\s-]?\d[\s-]?\d{4}[\s-]?\d{4})", texto)
    if telefones:
        return formatar_telefone(telefones[0])
    
    return "NÃO INFORMADO"

def formatar_telefone(telefone):
    """Formata o telefone para (XX) X XXXX-XXXX"""
    numeros = re.sub(r'[^\d]', '', telefone)
    if len(numeros) == 11:
        return f"({numeros[:2]}) {numeros[2]} {numeros[3:7]}-{numeros[7:]}"
    elif len(numeros) == 10:
        return f"({numeros[:2]}) {numeros[2:6]}-{numeros[6:]}"
    return telefone.strip()

def processar_dados(texto):
    """Processa o texto e retorna apenas os dados necessários"""
    dados = {
        "Quantos cartão": "1",  # Valor padrão
        "CNPJ": re.search(r"CNPJ[:\s]*([\d./-]+)", texto).group(1) if re.search(r"CNPJ", texto) else "NÃO INFORMADO",
        "Nome Social": re.search(r"(RAZÃO SOCIAL|CEDENTE|GESTOR MASTER)[:\s]*([^\n]+)", texto, re.IGNORECASE).group(2).strip() if re.search(r"(RAZÃO SOCIAL|CEDENTE|GESTOR MASTER)", texto, re.IGNORECASE) else "NÃO INFORMADO",
        "Endereço": re.search(r"RUA[:\s]*([^\n;]+)", texto).group(1).strip() if re.search(r"RUA", texto) else "NÃO INFORMADO",
        "Numero": re.search(r"NÚMERO[:\s]*([^\n;]+)", texto).group(1).strip() if re.search(r"NÚMERO", texto) else "NÃO INFORMADO",
        "Complemento": re.search(r"PONTO DE REF[:\s]*([^\n;]+)", texto).group(1).strip() if re.search(r"PONTO DE REF", texto) else "SEM PONTO",
        "Cidade": re.search(r"CIDADE[:\s]*([^\n-]+)", texto).group(1).strip() if re.search(r"CIDADE", texto) else "NÃO INFORMADO",
        "UF": re.search(r"ESTADO[:\s]*([A-Z]{2})", texto).group(1) if re.search(r"ESTADO", texto) else "NÃO INFORMADO",
        "Telefone": extrair_telefone_principal(texto),
        "Email": re.search(r"E-?MAIL[:\s]*([^\s]+@[^\s]+)", texto, re.IGNORECASE).group(1) if re.search(r"E-?MAIL", texto, re.IGNORECASE) else "NÃO INFORMADO",
        "Vendedor": re.search(r"(GESTOR MASTER|VENDEDOR)[:\s]*([^\n]+)", texto, re.IGNORECASE).group(2).strip() if re.search(r"(GESTOR MASTER|VENDEDOR)", texto, re.IGNORECASE) else "NÃO INFORMADO"
    }
    
    # Correção para complemento vazio
    if dados["Complemento"] == "":
        dados["Complemento"] = "SEM PONTO"
    
    return dados

def salvar_planilha(dados, arquivo="dados_extraidos.xlsx"):
    """Salva os dados na planilha mantendo a ordem das colunas"""
    colunas = [
        "Quantos cartão", "CNPJ", "Nome Social", "Endereço", "Numero",
        "Complemento", "Cidade", "UF", "Telefone", "Email", "Vendedor"
    ]
    
    try:
        df = pd.DataFrame([dados])[colunas]
        
        if os.path.exists(arquivo):
            df_existente = pd.read_excel(arquivo)
            # Garante as mesmas colunas no arquivo existente
            for col in colunas:
                if col not in df_existente.columns:
                    df_existente[col] = "NÃO INFORMADO"
            df_final = pd.concat([df_existente, df], ignore_index=True)
        else:
            df_final = df
        
        df_final.to_excel(arquivo, index=False)
        print("✅ Dados salvos com sucesso!")
        
    except Exception as e:
        print(f"❌ Erro ao salvar planilha: {e}")

# Exemplo de uso
texto = """
LIVIA 

VENDA CANTADA 

ENDEREÇO: 

RUA: FLORESTA
BAIRRO: SAO CAETANO
NÚMERO: 331 (EDIF CRISTAL APT 203)
CIDADE: ITABUNA
CEP: 45.607-090
PONTO DE REF: SEM PONTO 

RAZÃO SOCIAL: J S B COMERCIO E REPRESENTACAO COMERCIAL LTDA
CNPJ: 56.868.552/0001-70

GESTOR MASTER: EDUARDO JOSE SOARES BRANDAO
CONTATO: (73) 9 8824-8659
E-MAIL: EDUARDOJSB10@HOTMAIL.COM

PLANO: 1L 100GB + 1l 1GB 
LINHA: (73) 9 8158-4374 (1GB) 
LINHA: (73) 9 8824-8659 (100GB)
"""

dados = processar_dados(texto)
print("📋 Dados extraídos:")
for k, v in dados.items():
    print(f"{k}: {v}")

salvar_planilha(dados)