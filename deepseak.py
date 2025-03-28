import tkinter as tk
from tkinter import messagebox, filedialog
import re
import pandas as pd
import os
from datetime import datetime
import requests

# Configuração da API DeepSeek
DEEPSEEK_API_KEY = "sk-6f66c0e4ff034a8eaabebd32fb9d628c"
DEEPSEEK_URL = "https://api.deepseek.com/v1/search"

class ExtratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Chips - Extrator de Dados")
        self.root.geometry("700x500")
        self.root.resizable(False, False)
        
        # Cores
        self.cor_vivo_roxo = "#660099"  
        self.cor_vivo_roxo_claro = "#9933CC"  
        self.cor_branco = "#FFFFFF"
        self.cor_texto = "#333333"
        self.cor_fundo = "#F5F5F5"
        
        # Configurar estilo
        self.root.configure(bg=self.cor_fundo)
        self.fonte = ("Arial", 10)
        self.fonte_titulo = ("Arial", 14, "bold")
        
        # Criar widgets
        self.criar_widgets()
        
    def criar_widgets(self):
        # Frame principal
        frame = tk.Frame(self.root, bg=self.cor_branco, padx=20, pady=20)
        frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
        
        # Cabeçalho com logo Chips
        header_frame = tk.Frame(frame, bg=self.cor_vivo_roxo)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        tk.Label(
            header_frame, 
            text="CHIPS", 
            font=("Arial", 18, "bold"), 
            bg=self.cor_vivo_roxo,
            fg=self.cor_branco,
            padx=10,
            pady=5
        ).pack(side=tk.LEFT)
        
        tk.Label(
            header_frame, 
            text="Extrator de Dados", 
            font=("Arial", 12), 
            bg=self.cor_vivo_roxo,
            fg=self.cor_branco,
            padx=10,
            pady=5
        ).pack(side=tk.LEFT)
        
        # Área de texto
        tk.Label(
            frame, 
            text="Cole o texto para extração", 
            font=self.fonte_titulo, 
            bg=self.cor_branco,
            fg=self.cor_vivo_roxo
        ).pack(pady=(0, 10))
        
        self.texto_input = tk.Text(
            frame, 
            height=15, 
            width=80, 
            font=self.fonte,
            wrap=tk.WORD,
            bg=self.cor_branco,
            fg=self.cor_texto,
            insertbackground=self.cor_vivo_roxo,
            selectbackground=self.cor_vivo_roxo_claro,
            padx=10,
            pady=10
        )
        self.texto_input.pack(pady=5)
        
        # Frame de botões
        btn_frame = tk.Frame(frame, bg=self.cor_branco)
        btn_frame.pack(pady=10)
        
        # Botão Enviar
        tk.Button(
            btn_frame, 
            text="Enviar para Planilha", 
            command=self.processar_texto,
            bg=self.cor_vivo_roxo,
            fg=self.cor_branco,
            font=self.fonte,
            padx=20,
            pady=5,
            activebackground=self.cor_vivo_roxo_claro,
            activeforeground=self.cor_branco,
            bd=0,
            highlightthickness=0
        ).pack(side=tk.LEFT, padx=5)
        
        # Botão Limpar
        tk.Button(
            btn_frame, 
            text="Limpar", 
            command=self.limpar_campos,
            bg=self.cor_branco,
            fg=self.cor_vivo_roxo,
            font=self.fonte,
            padx=20,
            pady=5,
            activebackground=self.cor_fundo,
            activeforeground=self.cor_vivo_roxo,
            bd=1,
            highlightthickness=0
        ).pack(side=tk.LEFT, padx=5)
        
        # Botão Selecionar Arquivo
        tk.Button(
            btn_frame, 
            text="Abrir Arquivo", 
            command=self.abrir_arquivo,
            bg=self.cor_vivo_roxo_claro,
            fg=self.cor_branco,
            font=self.fonte,
            padx=20,
            pady=5,
            activebackground=self.cor_vivo_roxo,
            activeforeground=self.cor_branco,
            bd=0,
            highlightthickness=0
        ).pack(side=tk.LEFT, padx=5)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Pronto para extrair dados")
        tk.Label(
            frame, 
            textvariable=self.status_var,
            bg=self.cor_branco,
            fg=self.cor_vivo_roxo,
            font=("Arial", 9)
        ).pack(pady=(10, 0))
    
    def processar_texto(self):
        texto = self.texto_input.get("1.0", tk.END).strip()
        if not texto:
            messagebox.showwarning("Aviso", "Por favor, cole algum texto para extrair os dados.")
            return
        
        try:
            dados = self.extrair_dados_com_deepseek(texto)                
            self.salvar_planilha(dados)
            self.status_var.set("Dados salvos com sucesso em dados_extraidos.xlsx")
            messagebox.showinfo("Sucesso", "Dados extraídos e salvos na planilha!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(e)}")
            self.status_var.set("Erro ao processar dados")
    
    def extrair_dados_com_deepseek(self, texto):
        """Extrai os dados usando a API DeepSeek"""
        try:
            headers = {
                "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
                "Content-Type": "application/json"
            }
            
            params = {
                "query": f"""Extraia do seguinte texto as informações no formato especificado:
                - CNPJ
                - Nome Social (RAZÃO SOCIAL ou CEDENTE ou GESTOR MASTER)
                - Endereço (RUA ou ENDEREÇO)
                - Número
                - Complemento (PONTO DE REFERÊNCIA)
                - Cidade
                - UF (ESTADO)
                - Telefone (CONTATO)
                - Email
                - Vendedor (GESTOR MASTER ou VENDEDOR)
                
                Texto: {texto}""",
                "limit": 1
            }
            
            response = requests.get(DEEPSEEK_URL, headers=headers, params=params)
            response_data = response.json()
            
            # Processar a resposta da API conforme necessário
            # Aqui você precisará adaptar conforme a estrutura de resposta da DeepSeek
            # Este é um exemplo genérico - ajuste conforme a API realmente retorna
            
            dados = {
                "Quantos cartão": "1",
                "CNPJ": self.extrair_campo(response_data, "CNPJ"),
                "Nome Social": self.extrair_campo(response_data, "Nome Social"),
                "Endereço": self.extrair_campo(response_data, "Endereço"),
                "Numero": self.extrair_campo(response_data, "Número"),
                "Complemento": self.extrair_campo(response_data, "Complemento") or "SEM PONTO",
                "Cidade": self.extrair_campo(response_data, "Cidade"),
                "UF": self.extrair_campo(response_data, "UF"),
                "Telefone": self.extrair_campo(response_data, "Telefone"),
                "Email": self.extrair_campo(response_data, "Email"),
                "Vendedor": self.extrair_campo(response_data, "Vendedor"),
                "Data Extração": datetime.now().strftime("%d/%m/%Y %H:%M")
            }
            
            return dados
                
        except Exception as e:
            print(f"Erro na API DeepSeek: {e}")
            # Se falhar, usa o método tradicional
            return self.extrair_dados(texto)
    
    def extrair_campo(self, response_data, campo):
        """Extrai um campo específico da resposta da API"""
        # Implemente a lógica específica para extrair cada campo da resposta da DeepSeek
        # Este é um placeholder - ajuste conforme a estrutura real da resposta
        if not response_data:
            return "NÃO INFORMADO"
            
        # Exemplo genérico - a implementação real depende da estrutura da resposta
        for item in response_data.get('results', []):
            if campo.lower() in item.get('text', '').lower():
                return item.get('text', 'NÃO INFORMADO')
        return "NÃO INFORMADO"
    
    def limpar_campos(self):
        self.texto_input.delete("1.0", tk.END)
        self.status_var.set("Pronto para extrair dados")
    
    def abrir_arquivo(self):
        filepath = filedialog.askopenfilename(
            filetypes=[("Arquivos de Texto", "*.txt"), ("Todos os arquivos", "*.*")]
        )
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
        """Extrai os dados do texto usando regex"""
        def safe_search(pattern, text, group=1, flags=0):
            match = re.search(pattern, text, flags) if text else None
            return match.group(group).strip() if match else "NÃO INFORMADO"

        dados = {
            "Quantos cartão": "1",
            "CNPJ": safe_search(r"CNPJ[:\s]*([\d./-]+)", texto),
            "Nome Social": safe_search(r"(RAZÃO SOCIAL|CEDENTE|GESTOR MASTER)[:\s]*([^\n]+)", texto, group=2, flags=re.IGNORECASE),
            "Endereço": safe_search(r"(RUA|ENDEREÇO)[:\s]*([^\n;]+)", texto, group=2),
            "Numero": safe_search(r"NÚMERO[:\s]*([^\n;]+)", texto),
            "Complemento": safe_search(r"PONTO DE REF[:\s]*([^\n;]+)", texto) or "SEM PONTO",
            "Cidade": safe_search(r"CIDADE[:\s]*([^\n-]+)", texto),
            "UF": safe_search(r"ESTADO[:\s]*([A-Z]{2})", texto),
            "Telefone": self.extrair_telefone_principal(texto),
            "Email": safe_search(r"E-?MAIL[:\s]*([^\s]+@[^\s]+)", texto, flags=re.IGNORECASE),
            "Vendedor": safe_search(r"(GESTOR MASTER|VENDEDOR)[:\s]*([^\n]+)", texto, group=2, flags=re.IGNORECASE),
            "Data Extração": datetime.now().strftime("%d/%m/%Y %H:%M")
        }
        
        if dados["Complemento"] == "":
            dados["Complemento"] = "SEM PONTO"
            
        return dados
    
    def extrair_telefone_principal(self, texto):
        """Extrai apenas o primeiro telefone de CONTATO, ignorando LINHAS de plano"""
        contato = re.search(r"CONTATO[:\s]*([\(\)\d\s\-+]+)", texto, re.IGNORECASE)
        if contato:
            return self.formatar_telefone(contato.group(1))
        
        telefones = re.findall(r"(?<!LINHA[:\s])(\(?\d{2}\)?[\s-]?\d[\s-]?\d{4}[\s-]?\d{4})", texto)
        if telefones:
            return self.formatar_telefone(telefones[0])
        
        return "NÃO INFORMADO"
    
    def formatar_telefone(self, telefone):
        """Formata o telefone para (XX) X XXXX-XXXX"""
        numeros = re.sub(r'[^\d]', '', telefone)
        if len(numeros) == 11:
            return f"({numeros[:2]}) {numeros[2]} {numeros[3:7]}-{numeros[7:]}"
        elif len(numeros) == 10:
            return f"({numeros[:2]}) {numeros[2:6]}-{numeros[6:]}"
        return telefone.strip()
    
    def salvar_planilha(self, dados):
        """Salva os dados na planilha mantendo a ordem das colunas"""
        colunas = [
            "Quantos cartão", "CNPJ", "Nome Social", "Endereço", "Numero",
            "Complemento", "Cidade", "UF", "Telefone", "Email", "Vendedor", "Data Extração"
        ]
        
        arquivo = "dados_extraidos.xlsx"
        
        if os.path.exists(arquivo):
            df_existente = pd.read_excel(arquivo)
            # Garante as mesmas colunas no arquivo existente
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