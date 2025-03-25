import tkinter as tk
from tkinter import messagebox, filedialog
import re
import pandas as pd
import os
from datetime import datetime
from transformers import pipeline

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
        
        # Configurar modelo Hugging Face
        try:
            # Usando um modelo BERT em português para NER (Reconhecimento de Entidades Nomeadas)
            self.ner_pipeline = pipeline(
                "ner", 
                model="neuralmind/bert-base-portuguese-cased",
                tokenizer="neuralmind/bert-base-portuguese-cased",
                aggregation_strategy="simple"
            )
            self.nlp_loaded = True
            self.status_modelo = "Modelo BERT carregado com sucesso"
        except Exception as e:
            print(f"Erro ao carregar modelo: {e}")
            self.nlp_loaded = False
            self.status_modelo = "Modelo BERT não carregado - usando regex"
            messagebox.showwarning(
                "Aviso", 
                "Modelo NLP não pôde ser carregado. Usando método regex padrão."
            )
        
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
        self.status_var.set(f"Pronto para extrair dados | {self.status_modelo}")
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
            if self.nlp_loaded and len(texto.split()) < 500:  # Limite para não sobrecarregar
                dados = self.extrair_dados_com_transformers(texto)
            else:
                dados = self.extrair_dados(texto)
                
            self.salvar_planilha(dados)
            self.status_var.set(f"Dados salvos com sucesso em dados_extraidos.xlsx | {self.status_modelo}")
            messagebox.showinfo("Sucesso", "Dados extraídos e salvos na planilha!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(e)}")
            self.status_var.set(f"Erro ao processar dados | {self.status_modelo}")
    
    def extrair_dados_com_transformers(self, texto):
        """Extrai dados usando Hugging Face Transformers"""
        # Primeiro usa o modelo de NER
        entidades = self.ner_pipeline(texto)
        
        # Organiza as entidades por tipo
        entidades_por_tipo = {}
        for ent in entidades:
            tipo = ent['entity_group']
            if tipo not in entidades_por_tipo:
                entidades_por_tipo[tipo] = []
            entidades_por_tipo[tipo].append(ent['word'])
        
        # Depois complementa com regex para padrões específicos
        dados = {
            "Quantos cartão": "-",
            "CNPJ": self.encontrar_padrao(texto, r"CNPJ[:\s]*([\d./-]+)"),
            "Nome Social": self.extrair_entidade(entidades_por_tipo, "ORG") or 
                          self.encontrar_padrao(texto, r"(RAZÃO SOCIAL|CEDENTE|GESTOR MASTER)[:\s]*([^\n]+)", 2),
            "Endereço": self.extrair_endereco(texto),
            "Numero": self.extrair_numero(texto),
            "Complemento": self.encontrar_padrao(texto, r"PONTO DE REF[:\s]*([^\n;]+)") or "SEM PONTO",
            "Cidade": self.extrair_entidade(entidades_por_tipo, "LOC") or 
                     self.encontrar_padrao(texto, r"CIDADE[:\s]*([^\n-]+)"),
            "UF": self.encontrar_padrao(texto, r"ESTADO[:\s]*([A-Z]{2})"),
            "Telefone": self.extrair_telefone_principal(texto),
            "Email": self.encontrar_padrao(texto, r"E-?MAIL[:\s]*([^\s]+@[^\s]+)"),
            "Vendedor": "-",
            "Data Extração": datetime.now().strftime("%d/%m/%Y %H:%M")
        }
        
        return dados
    
    def extrair_endereco(self, texto):
        """Extrai endereço considerando diferentes formatos (R, R., R/, R: etc.)"""
        padrao = r"(?:RUA|R\.|R/|R:|R)[\s:]*([^\n;,]+)"
        match = re.search(padrao, texto, re.IGNORECASE)
        if match:
            endereco = match.group(1).strip()
            # Remove possíveis números que pertencem ao endereço (deixando apenas o nome da rua)
            endereco = re.sub(r'\d+.*$', '', endereco).strip()
            return endereco
        return "NÃO INFORMADO"
    
    def extrair_numero(self, texto):
        """Extrai número da casa considerando diferentes formatos (NUMERO, N, Nº etc.)"""
        # Primeiro tenta encontrar padrões explícitos como "Nº", "N:", etc.
        padrao_explicito = r"(?:NUMERO|NÚMERO|N\.|Nº|N:|N)[\s:]*([^\n;,]+)"
        match = re.search(padrao_explicito, texto, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        
        # Se não encontrar, procura números no final do endereço
        padrao_endereco = r"(?:RUA|R\.|R/|R:|R)[\s:]*([^\n;,]+)"
        match_endereco = re.search(padrao_endereco, texto, re.IGNORECASE)
        if match_endereco:
            endereco_completo = match_endereco.group(1)
            # Procura o último número no endereço
            numeros = re.findall(r'\d+', endereco_completo)
            if numeros:
                return numeros[-1]
        
        return "NÃO INFORMADO"
    
    def extrair_entidade(self, entidades_por_tipo, tipo_entidade):
        """Extrai a primeira entidade do tipo especificado"""
        if tipo_entidade in entidades_por_tipo:
            return " ".join(entidades_por_tipo[tipo_entidade][:3])  # Limita a 3 palavras
        return "NÃO INFORMADO"
    
    def encontrar_padrao(self, texto, padrao, grupo=1):
        """Encontra padrões usando regex"""
        match = re.search(padrao, texto, re.IGNORECASE)
        return match.group(grupo).strip() if match else "NÃO INFORMADO"
    
    def limpar_campos(self):
        self.texto_input.delete("1.0", tk.END)
        self.status_var.set(f"Pronto para extrair dados | {self.status_modelo}")
    
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
                self.status_var.set(f"Arquivo carregado: {os.path.basename(filepath)} | {self.status_modelo}")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível abrir o arquivo:\n{str(e)}")
    
    def extrair_dados(self, texto):
        """Extrai os dados do texto usando regex"""
        dados = {
            "Quantos cartão": "-",
            "CNPJ": re.search(r"CNPJ[:\s]*([\d./-]+)", texto).group(1) if re.search(r"CNPJ", texto) else "NÃO INFORMADO",
            "Nome Social": re.search(r"(RAZÃO SOCIAL|CEDENTE|GESTOR MASTER)[:\s]*([^\n]+)", texto, re.IGNORECASE).group(2).strip() if re.search(r"(RAZÃO SOCIAL|CEDENTE|GESTOR MASTER)", texto, re.IGNORECASE) else "NÃO INFORMADO",
            "Endereço": self.extrair_endereco(texto),
            "Numero": self.extrair_numero(texto),
            "Complemento": re.search(r"PONTO DE REF[:\s]*([^\n;]+)", texto).group(1).strip() if re.search(r"PONTO DE REF", texto) else "SEM PONTO",
            "Cidade": re.search(r"CIDADE[:\s]*([^\n-]+)", texto).group(1).strip() if re.search(r"CIDADE", texto) else "NÃO INFORMADO",
            "UF": re.search(r"ESTADO[:\s]*([A-Z]{2})", texto).group(1) if re.search(r"ESTADO", texto) else "NÃO INFORMADO",
            "Telefone": self.extrair_telefone_principal(texto),
            "Email": re.search(r"E-?MAIL[:\s]*([^\s]+@[^\s]+)", texto, re.IGNORECASE).group(1) if re.search(r"E-?MAIL", texto, re.IGNORECASE) else "NÃO INFORMADO",
            "Vendedor": "-",
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