# 📋 Extrator de Dados Automatizado com IA

**Um sistema inteligente para extração e organização de informações de documentos textuais usando Google Gemini AI**

## 🌟 Visão Geral

Este projeto foi desenvolvido para automatizar a extração de dados importantes (como CNPJ, CPF, endereços, contatos etc.) de documentos textuais, transformando informações não estruturadas em dados organizados em planilhas Excel.

## ✨ Por que este projeto foi criado?

- **Automatizar** o processo manual de extração de dados de documentos
- **Reduzir erros** humanos na transcrição de informações
- **Acelerar** o processo de organização de dados
- **Integrar IA** para lidar com variações de formato nos documentos
- **Padronizar** a saída em um formato consistente e utilizável

## 🚀 Como usar

### Pré-requisitos
1. Tenha uma chave de API do Google Gemini (gratuita no [Google AI Studio](https://aistudio.google.com/))
2. Python 3.8+ instalado
3. Pacotes necessários: `pip install google-generativeai pandas requests`

### Configuração
1. Substitua `"AIzaSyAbCjFYEaFmfNqEZxmWGJjCfD76XuzU7oY"` por sua chave de API
2. (Opcional) Ajuste as configurações do Gemini no código conforme necessidade

### Execução
```bash
python main.py
```

### Entrada
Cole o texto contendo os dados a serem extraídos na variável `texto_exemplo` ou modifique o código para ler de:
- Arquivos de texto
- PDFs convertidos para texto
- Entrada do usuário

### Saída
Os dados serão salvos em `dados_extraidos.xlsx`, criando o arquivo ou atualizando um existente.

## 🔍 Funcionalidades Principais

1. **Extração Inteligente** usando Google Gemini AI
2. **Fallback para Regex** quando a IA não está disponível
3. **Correção Automática** de erros gramaticais com LanguageTool
4. **Normalização** de formatos (CNPJ, CPF, telefones, CEP)
5. **Preenchimento Completo** de todos campos solicitados
6. **Atualização** de planilhas Excel existentes

## 📊 Campos Extraídos

| Campo | Exemplo | Observação |
|-------|---------|------------|
| BAIRRO | CAMPO DO GOVERNO | - |
| CEP | 46880-000 | Sempre formatado |
| CIDADE | ITABERABA | - |
| CNPJ | 53.554.993/0001-00 | Formatado corretamente |
| COMPLEMENTO | TERREO CASA | - |
| CPF | 839.254.065-49 | Formatado corretamente |
| EMAIL | figueredorogerio300@gmail.com | - |
| ENDEREÇO | rua 10, N° 23 | - |
| ESTADO | BA | Sigla de 2 letras |
| TELEFONE | (75) 99177-9847 | Formatado corretamente |
| RAZÃO SOCIAL | ROGERIO FIGUEREDO SOUZA | - |

## 🛠️ Personalização

Você pode ajustar:
- Os campos a serem extraídos modificando `campos_necessarios`
- A sensibilidade da IA ajustando `generation_config`
- Os padrões de regex para casos específicos
- O nome do arquivo de saída

## ⚡ Vantagens

✅ **Precisão**: Combina IA com validação por regex  
✅ **Flexibilidade**: Lida com diferentes formatos de documento  
✅ **Eficiência**: Processa dados em segundos  
✅ **Atualizável**: Fácil de adicionar novos campos  
✅ **Integração**: Funciona com planilhas existentes  

## 📝 Notas

- Para documentos muito longos, pode ser necessário dividir o texto
- A API gratuita do Gemini tem limites de uso
- Sempre verifique os dados extraídos contra a fonte original

Este projeto transforma horas de trabalho manual em um processo automático de segundos, mantendo a precisão e organização dos dados! 🚀
