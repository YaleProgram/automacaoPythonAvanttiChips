import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime
import google.generativeai as genai

# Configurar a API do Gemini
genai.configure(api_key="AIzaSyBQVXUpkf_G72WaXeBMlqDDEZN-jLpVPiY")  # Substitua pela sua chave da API

class ExtratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Extrator de Dados")
        
        # Interface gráfica
        self.texto_input = tk.Text(root, height=10, width=80)
        self.texto_input.pack()
        
        self.processar_button = tk.Button(root, text="Processar Texto", command=self.processar_texto)
        self.processar_button.pack()
        
        self.limpar_button = tk.Button(root, text="Limpar", command=self.limpar_campos)
        self.limpar_button.pack()
        
        self.abrir_button = tk.Button(root, text="Abrir Arquivo", command=self.abrir_arquivo)
        self.abrir_button.pack()
        
        self.status_var = tk.StringVar()
        self.status_label = tk.Label(root, textvariable=self.status_var)
        self.status_label.pack()
        
        self.status_var.set("Pronto para extrair dados")

    def processar_texto(self):
        texto = self.texto_input.get("1.0", tk.END).strip()
        if not texto:
            messagebox.showwarning("Aviso", "Por favor, cole algum texto para extrair os dados.")
            return
        
        try:
            dados = self.extrair_dados(texto)
            self.salvar_planilha(dados)
            self.status_var.set("Dados salvos com sucesso em dados_extraidos.xlsx")
            messagebox.showinfo("Sucesso", "Dados extraídos e salvos na planilha!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(e)}")
            self.status_var.set("Erro ao processar dados")

    def limpar_campos(self):
        self.texto_input.delete("1.0", tk.END)
        self.status_var.set("Pronto para extrair dados")

    def abrir_arquivo(self):
        filepath = filedialog.askopenfilename(filetypes=[("Arquivos de Texto", "*.txt"), ("Todos os arquivos", "*.*")])
        if filepath:
            try:
                with open(filepath, 'r', encoding='utf-8') as file:
                    conteudo = file.read()
                    self.texto_input.delete("1.0", tk.END)
                    self.texto_input.insert("1.0", conteudo)
                self.status_var.set(f"Arquivo carregado: {os.path.basename(filepath)}")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível abrir o arquivo:\n{str(e)}")

    def extrair_dados(self, texto):
        """Envia o texto para o Gemini e extrai os dados necessários"""
        prompt = f"""
        Extraia os seguintes dados do texto fornecido:

        - Quantos cartão
        - CNPJ
        - Nome Social (Razão Social, Cedente ou Gestor Master)
        - Endereço
        - Número
        - Complemento
        - Cidade
        - UF (Estado)
        - Telefone
        - Email
        - Vendedor (Gestor Master ou Vendedor)

        Texto de entrada:
        {texto}

        Formate a saída como um JSON com as chaves acima.
        """

        try:
            response = genai.GenerativeModel("gemini-pro").generate_content(prompt)
            dados = eval(response.text)  # Converte o JSON em dicionário
        except Exception as e:
            messagebox.showerror("Erro", f"Falha na comunicação com a API do Gemini:\n{str(e)}")
            return None

        # Garantir que todas as chaves existam no dicionário
        campos_padrao = {
            "Quantos cartão": "1",
            "CNPJ": "NÃO INFORMADO",
            "Nome Social": "NÃO INFORMADO",
            "Endereço": "NÃO INFORMADO",
            "Numero": "NÃO INFORMADO",
            "Complemento": "SEM PONTO",
            "Cidade": "NÃO INFORMADO",
            "UF": "NÃO INFORMADO",
            "Telefone": "NÃO INFORMADO",
            "Email": "NÃO INFORMADO",
            "Vendedor": "NÃO INFORMADO",
            "Data Extração": datetime.now().strftime("%d/%m/%Y %H:%M")
        }

        # Atualiza os valores extraídos com os padrões caso faltem informações
        for chave in campos_padrao:
            if chave not in dados or not dados[chave].strip():
                dados[chave] = campos_padrao[chave]

        return dados

    def salvar_planilha(self, dados):
        """Salva os dados na planilha mantendo a ordem das colunas"""
        colunas = [
            "Quantos cartão", "CNPJ", "Nome Social", "Endereço", "Numero",
            "Complemento", "Cidade", "UF", "Telefone", "Email", "Vendedor", "Data Extração"
        ]

        arquivo = "dados_extraidos.xlsx"

        if os.path.exists(arquivo):
            df_existente = pd.read_excel(arquivo)
            for col in colunas:
                if col not in df_existente.columns:
                    df_existente[col] = "NÃO INFORMADO"
            df_novo = pd.DataFrame([dados])[colunas]
            df_final = pd.concat([df_existente, df_novo], ignore_index=True)
        else:
            df_final = pd.DataFrame([dados])[colunas]

        df_final.to_excel(arquivo, index=False)

# Criar e rodar a aplicação
if __name__ == "__main__":
    root = tk.Tk()
    app = ExtratorApp(root)
    root.mainloop()
