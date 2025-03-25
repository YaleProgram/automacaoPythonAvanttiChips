import tkinter as tk
from tkinter import messagebox, filedialog
import re
import pandas as pd
import os
from datetime import datetime
import google.generativeai as genai

# Colocar a chave da API do Gemini:
chaveApi = "AIzaSyCOif8UqGDkjtswQfe-rIkD4ZEX0upoqXg"

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
        
        # Configurar Gemini
        self.configurar_gemini()
        
        # Criar widgets
        self.criar_widgets()
        
    def configurar_gemini(self):
        """Configura a API do Gemini"""
        try:
            genai.configure(api_key=chaveApi)
            # Usando o modelo correto para a API atual
            self.gemini_model = genai.GenerativeModel('gemini-1.0-pro')
            self.gemini_prompt = """Analise o texto e extraia as informações no formato especificado. 
            Retorne APENAS um dicionário Python válido com as chaves exatamente como especificado.
            Se alguma informação não for encontrada, use "NÃO INFORMADO".
            
            IMPORTANTE: 
            - Não inclua comentários ou explicações no retorno
            - O retorno deve ser APENAS o dicionário Python
            - As chaves devem estar exatamente como especificado
            
            Formato do dicionário:
            {
                "CNPJ": "valor",
                "Nome Social": "valor",
                "Endereço": "valor",
                "Numero": "valor",
                "Complemento": "valor",
                "Cidade": "valor",
                "UF": "valor",
                "Telefone": "valor",
                "Email": "valor",
                "Vendedor": "valor"
            }
            
            Texto para análise:
            """
        except Exception as e:
            print(f"Erro ao configurar Gemini: {e}")
            self.gemini_model = None
    
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
            # Tenta primeiro com o Gemini
            if self.gemini_model:
                dados = self.extrair_dados_com_gemini(texto)
            else:
                dados = self.extrair_dados(texto)
                
            self.salvar_planilha(dados)
            self.status_var.set("Dados salvos com sucesso em dados_extraidos.xlsx")
            messagebox.showinfo("Sucesso", "Dados extraídos e salvos na planilha!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(e)}")
            self.status_var.set("Erro ao processar dados")
    
    def extrair_dados_com_gemini(self, texto):
        """Extrai os dados usando o Gemini AI"""
        try:
            full_prompt = self.gemini_prompt + texto
            response = self.gemini_model.generate_content(full_prompt)
            
            if response.text:
                # Limpa a resposta removendo possíveis marcações de código
                clean_text = response.text.replace('```python', '').replace('```', '').strip()
                dados = eval(clean_text)
                
                # Adiciona campos padrão
                dados["Quantos cartão"] = "1"
                dados["Data Extração"] = datetime.now().strftime("%d/%m/%Y %H:%M")
                
                # Garante que todos os campos existam
                campos_necessarios = [
                    "Quantos cartão", "CNPJ", "Nome Social", "Endereço", "Numero",
                    "Complemento", "Cidade", "UF", "Telefone", "Email", "Vendedor", "Data Extração"
                ]
                
                for campo in campos_necessarios:
                    if campo not in dados:
                        dados[campo] = "NÃO INFORMADO"
                
                return dados
            else:
                raise ValueError("Gemini não retornou dados válidos")
                
        except Exception as e:
            print(f"Erro no Gemini: {e}")
            # Se falhar, usa o método tradicional
            return self.extrair_dados(texto)
    
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