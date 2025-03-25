# 📋 Extrator de Dados para Planilha Automatizado

## 🌟 Visão Geral
Este script Python automatiza a extração de informações específicas de textos estruturados (como contratos ou formulários) e organiza em uma planilha Excel com formato padronizado.

## ✨ Funcionalidades Principais
- Extrai **11 campos específicos** de documentos textuais
- Foca no **telefone principal** (CONTATO), ignorando linhas adicionais
- Formata automaticamente **CNPJ e telefones**
- Organiza os dados em **Excel** mantendo a ordem das colunas
- Atualiza planilhas existentes ou cria novas

## 📋 Campos Extraídos
| Campo | Exemplo | Observação |
|-------|---------|------------|
| Quantos cartão | 1 | Valor padrão "1" |
| CNPJ | 56.868.552/0001-70 | Formatado corretamente |
| Nome Social | J S B COMERCIO LTDA | - |
| Endereço | RUA FLORESTA | Sem número |
| Numero | 331 | Extraído do endereço |
| Complemento | EDIF CRISTAL APT 203 | Ou "SEM PONTO" |
| Cidade | ITABUNA | - |
| UF | BA | Sigla com 2 letras |
| Telefone | (73) 9 8824-8659 | Primeiro CONTATO encontrado |
| Email | email@exemplo.com | - |
| Vendedor | EDUARDO BRANDAO | - |

## ⚙️ Pré-requisitos
Para utilizar este script, você precisa ter instalado:

```bash
pip install google-generativeai pandas openpyxl
```

## 🛠️ Como Usar
1. Substitua `"AIzaSyAbCjFYEaFmfNqEZxmWGJjCfD76XuzU7oY"` por sua [chave de API do Google Gemini](https://aistudio.google.com/)
2. Cole o texto a ser processado na variável `texto`
3. Execute o script:
```bash
python seu_script.py
```

4. Os dados serão salvos em `dados_extraidos.xlsx`

## 🔍 Lógica de Extração
1. **Telefone Principal**: Extrai APENAS o número após "CONTATO:", ignorando "LINHA:"
2. **Endereço Completo**: Separa automaticamente rua, número e complemento
3. **Valores Padrão**: Preenche com "NÃO INFORMADO" ou valores padrão quando necessário
4. **Formatação Automática**:
   - CNPJ: `XX.XXX.XXX/XXXX-XX`
   - Telefone: `(XX) X XXXX-XXXX`
   - Complemento: "SEM PONTO" quando vazio

## ⚠️ Limitações
- Funciona melhor com textos estruturados (formulários padronizados)
- Assume que o primeiro telefone após "CONTATO:" é o principal
- Requer conexão com internet para a API do Gemini

## 📌 Dicas
1. Para múltiplos documentos, crie um loop para processar vários textos
2. Verifique sempre os dados extraídos contra o original
3. Ajuste os regex caso seus documentos tenham formatos diferentes

Este script economiza horas de trabalho manual de extração e organização de dados, garantindo consistência e padronização nas suas planilhas! 🚀
