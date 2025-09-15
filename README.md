# SheetChat

SheetChat é um **chatbot inteligente para planilhas**, que permite interagir com arquivos Excel ou CSV e obter respostas rápidas sobre os dados contidos neles. O bot entende perguntas em linguagem natural e pode analisar múltiplas abas de uma planilha, fornecendo respostas claras e organizadas.

[![Status](https://img.shields.io/badge/status-em%20desenvolvimento-yellow)]()  
[![Python](https://img.shields.io/badge/python-3.13-blue?logo=python)]()  
[![License](https://img.shields.io/badge/license-Rafa25MF-green)]()  
---

## 🔹 Funcionalidades

- 📄 Suporte a arquivos Excel (.xlsx, .xls) e CSV  
- 🗂 Suporte a múltiplas abas, com navegação via botões  
- 💬 Chat estilo bot moderno, com respostas em caixas verdes  
- 🔎 Responde perguntas objetivas sobre dados (linhas, colunas, valores, soma, média, máximo, mínimo)  
- 💰 Responde perguntas como “Qual foi o lucro total de maio + junho + julho”  
- 🖥 Visualização organizada das primeiras linhas da planilha  
- 🧹 Botão para limpar o chat  
- ✨ Interface moderna usando CustomTkinter  

---

## 🔹 Tecnologias Utilizadas

- **Python 3.13** – Linguagem principal do projeto, responsável pela lógica do chatbot e manipulação de dados.  
- **[CustomTkinter](https://github.com/TomSchimansky/CustomTkinter)** – Biblioteca para criação de interfaces gráficas modernas, responsivas e personalizáveis.  
- **[Pandas](https://pandas.pydata.org/)** – Biblioteca poderosa para análise e manipulação de planilhas e dados tabulares.  
- **[Google Gemini API](https://ai.google.com/studio)** – Plataforma de IA generativa utilizada para fornecer respostas avançadas em linguagem natural.


---

## ⚙️ Tecnologias Utilizadas  

| Funcionalidade | Biblioteca |
|----------------|------------|
| Linguagem principal | [Python](https://www.python.org/) 🐍 |
| Manipulação de dados em planilhas | [Pandas](https://pandas.pydata.org/) |
| Inteligência Artificial (LLM) | [Google Generative AI](https://ai.google.dev/) |
| Interface gráfica moderna | [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) |
| Interface gráfica padrão | [Tkinter](https://docs.python.org/3/library/tkinter.html) |
| Expressões regulares | [re (Regex)](https://docs.python.org/3/library/re.html) |

---

## 📦 Instalação das Dependências  

Para instalar todas as bibliotecas necessárias de uma vez:  

```bash
pip install pandas google-generativeai customtkinter
