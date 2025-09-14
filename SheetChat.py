import pandas as pd
import google.generativeai as genai
import customtkinter as ctk
from tkinter import filedialog
import re

# ============================
# 1. Configurar Gemini
# ============================
genai.configure(api_key="AIzaSyC7K0Nj3Y1gTZ0P9LOZ9qPlrYNEI439rFU")
model = genai.GenerativeModel("gemini-1.5-flash")

# ============================
# 2. Funções planilha
# ============================
def selecionar_arquivo():
    caminho = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("CSV", "*.csv")]
    )
    return caminho

def carregar_arquivo_excel(caminho):
    try:
        if caminho.endswith((".xlsx", ".xls")):
            return pd.ExcelFile(caminho)
        elif caminho.endswith(".csv"):
            return pd.ExcelFile(caminho)
        else:
            raise ValueError("Formato não suportado")
    except Exception as e:
        ctk.CTkMessagebox.show_error(title="Erro", message=f"Erro ao abrir arquivo: {e}")
        return None

def carregar_aba(xls, nome_aba):
    try:
        return pd.read_excel(xls, sheet_name=nome_aba)
    except Exception as e:
        ctk.CTkMessagebox.show_error(title="Erro", message=f"Erro ao abrir aba {nome_aba}: {e}")
        return None

def formatar_planilha(df, linhas=10):
    """Formata planilha para exibir no chat de forma limpa"""
    df_mostra = df.head(linhas).copy()
    df_mostra = df_mostra.fillna("")  # vazio no lugar de NaN
    # Converte em string bem formatada
    return df_mostra.to_string(index=False)

# ============================
# 3. Interpretação de perguntas
# ============================
def celula_para_indices(celula):
    match = re.match(r"([A-Z]+)(\d+)", celula.upper())
    if not match:
        return None
    col_str, row_str = match.groups()
    col_idx = sum([(ord(char)-65)*(26**i) for i, char in enumerate(reversed(col_str))])
    row_idx = int(row_str)-1
    return row_idx, col_idx

def detectar_python(pergunta):
    palavras_chave = ["quantas linhas","quais colunas","linha","coluna","celula","célula","lucro total","soma","média","maior","menor"]
    return any(palavra in pergunta.lower() for palavra in palavras_chave)

def responder_com_python(pergunta, df):
    pl = pergunta.lower()
    if "quantas linhas" in pl:
        return f"A planilha atual tem {len(df)} linhas."
    if "quais colunas" in pl or "nome das colunas" in pl:
        return f"As colunas da planilha são: {list(df.columns)}"
    match = re.search(r"linha (\d+).*coluna (\w+)|c[eé]lula (\w+\d+)", pl)
    if match:
        try:
            if match.group(3):
                linha, col = celula_para_indices(match.group(3))
            else:
                linha = int(match.group(1))-1
                col = ord(match.group(2).upper())-65
            valor = df.iloc[linha, col]
            return f"O valor na célula especificada é: {valor}"
        except:
            return "Não consegui interpretar a célula."
    # Soma, média, máximo, mínimo
    for col in df.columns:
        if "lucro" in col.lower() or "venda" in col.lower():
            if "lucro total" in pl or "soma" in pl:
                return f"O total da coluna '{col}' é: {df[col].sum()}"
            if "média" in pl:
                return f"A média da coluna '{col}' é: {df[col].mean()}"
            if "maior" in pl:
                idx = df[col].idxmax()
                return f"O maior valor de '{col}' é {df[col].max()} na linha {idx+1}"
            if "menor" in pl:
                idx = df[col].idxmin()
                return f"O menor valor de '{col}' é {df[col].min()} na linha {idx+1}"
    return None

def perguntar_gemini(pergunta, df):
    resumo = formatar_planilha(df, linhas=100)
    prompt = f"""
Você é um assistente de dados amigável. Explique os dados de forma clara e resumida, sem termos técnicos complicados. 
Aqui está um resumo da planilha:

{resumo}

Pergunta: {pergunta}
"""
    resposta = model.generate_content(prompt)
    return resposta.text

# ============================
# 4. ChatBot com CustomTkinter
# ============================
class ChatBotPlanilha(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("ChatBot de Planilha")
        self.geometry("900x650")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("green")

        # Estruturas
        self.df_dict = {}
        self.aba_atual = None

        # Layout
        self.chat_frame = ctk.CTkScrollableFrame(self, width=880, height=450)
        self.chat_frame.pack(padx=10, pady=10, fill="both", expand=True)

        self.entry_frame = ctk.CTkFrame(self)
        self.entry_frame.pack(padx=10, pady=5, fill="x")

        self.entrada = ctk.CTkEntry(self.entry_frame, placeholder_text="Digite sua pergunta...", font=("Arial", 14))
        self.entrada.pack(side="left", padx=5, pady=5, fill="x", expand=True)
        self.entrada.bind("<Return>", self.processar_pergunta)

        self.enviar_btn = ctk.CTkButton(self.entry_frame, text="Enviar", command=self.processar_pergunta)
        self.enviar_btn.pack(side="left", padx=5, pady=5)

        self.btn_frame = ctk.CTkFrame(self)
        self.btn_frame.pack(padx=10, pady=5, fill="x")

        self.add_file_btn = ctk.CTkButton(self.btn_frame, text="Adicionar Planilha", command=self.adicionar_planilha)
        self.add_file_btn.pack(side="left", padx=5)

        self.limpar_btn = ctk.CTkButton(self.btn_frame, text="Limpar Chat", command=self.limpar_chat)
        self.limpar_btn.pack(side="left", padx=5)

        self.aba_buttons_frame = ctk.CTkFrame(self.btn_frame)
        self.aba_buttons_frame.pack(side="left", padx=10)

    def escrever_chat(self, texto, user=False):
        cor = "#022cd7" if user else "#0629b3"
        label = ctk.CTkLabel(self.chat_frame, text=texto, wraplength=800, fg_color=cor, corner_radius=10, anchor="w", justify="left", font=("Arial", 14))
        label.pack(padx=5, pady=5, anchor="w" if not user else "e")

    def limpar_chat(self):
        for widget in self.chat_frame.winfo_children():
            widget.destroy()

    def adicionar_planilha(self):
        caminho = selecionar_arquivo()
        if not caminho:
            return
        xls = carregar_arquivo_excel(caminho)
        if not xls:
            return
        for aba in xls.sheet_names:
            df = carregar_aba(xls, aba)
            self.df_dict[aba] = df
            btn = ctk.CTkButton(self.aba_buttons_frame, text=aba, command=lambda a=aba: self.mudar_aba(a))
            btn.pack(side="left", padx=2)
        self.aba_atual = xls.sheet_names[0]

    def mudar_aba(self, aba):
        if aba in self.df_dict:
            self.aba_atual = aba
            self.escrever_chat(f"📊 Agora analisando a aba: {aba}")
            

    def processar_pergunta(self, event=None):
        pergunta = self.entrada.get()
        if not pergunta:
            return
        self.entrada.delete(0, "end")
        self.escrever_chat(pergunta, user=True)

        if self.aba_atual is None:
            self.escrever_chat("❌ Nenhuma aba selecionada ainda.")
            return

        df = self.df_dict[self.aba_atual]
        if detectar_python(pergunta):
            resposta = responder_com_python(pergunta, df)
            if resposta is None:
                resposta = perguntar_gemini(pergunta, df)
        else:
            resposta = perguntar_gemini(pergunta, df)

        self.escrever_chat(resposta, user=False)
        

# ============================
# 5. Executar
# ============================
if __name__ == "__main__":
    app = ChatBotPlanilha()
    app.mainloop()
