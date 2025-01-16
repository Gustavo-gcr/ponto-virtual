import tkinter as tk
from tkinter import messagebox
from datetime import datetime, timedelta
import pandas as pd
import os

# Nome do arquivo Excel
def get_excel_filename():
    mes_atual = datetime.now().strftime("%B").capitalize()
    ano_atual = datetime.now().strftime("%Y")
    return f"{mes_atual}_{ano_atual}_horarios_gustavo.xlsx"

EXCEL_FILE = get_excel_filename()

# Criar o Excel inicial se não existir
def criar_planilha_se_nao_existir():
    if not os.path.exists(EXCEL_FILE):
        colunas = ["Dia", "Chegada", "Almoco", "Saida", "Horas Totais", "OBS"]
        
        # Determinar o número de dias no mês atual
        hoje = datetime.now()
        primeiro_dia = hoje.replace(day=1)
        ultimo_dia = (primeiro_dia + timedelta(days=32)).replace(day=1) - timedelta(days=1)
        dias_no_mes = ultimo_dia.day
        
        # Preencher os dias do mês
        dias = list(range(1, dias_no_mes + 1))
        df = pd.DataFrame({
            "Dia": dias,
            "Chegada": ["" for _ in dias],
            "Almoco": ["" for _ in dias],
            "Saida": ["" for _ in dias],
            "Horas Totais": ["" for _ in dias],
            "OBS": ["" for _ in dias]
        })
        
        # Salvar o DataFrame no arquivo Excel
        df.to_excel(EXCEL_FILE, index=False)

def registrar_horario(tipo):
    try:
        # Obter o dia atual
        dia_atual = datetime.now().day
        horario_atual = datetime.now().strftime("%H:%M")

        # Carregar o Excel existente
        df = pd.read_excel(EXCEL_FILE)

        # Converter a coluna "Dia" para o tipo inteiro (caso necessário)
        df["Dia"] = df["Dia"].astype(int)

        # Garantir que as colunas sejam do tipo string
        for coluna in ["Chegada", "Almoco", "Saida", "Horas Totais", "OBS"]:
            if coluna in df.columns:
                df[coluna] = df[coluna].fillna("")  # Substituir NaN por vazio

        # Verificar se o dia atual existe no DataFrame
        if dia_atual not in df["Dia"].values:
            raise ValueError(f"O dia {dia_atual} não foi encontrado na planilha.")

        # Atualizar o horário correspondente
        if tipo == "Chegada":
            df.loc[df["Dia"] == dia_atual, "Chegada"] = horario_atual
        elif tipo == "Almoco":
            df.loc[df["Dia"] == dia_atual, "Almoco"] = horario_atual
        elif tipo == "Saida":
            df.loc[df["Dia"] == dia_atual, "Saida"] = horario_atual

        # Calcular horas totais se "Chegada" e "Saída" estiverem preenchidas
        chegada = df.loc[df["Dia"] == dia_atual, "Chegada"].values[0]
        saida = df.loc[df["Dia"] == dia_atual, "Saida"].values[0]

        if chegada and saida:
            try:
                h_chegada = datetime.strptime(chegada, "%H:%M")
                h_saida = datetime.strptime(saida, "%H:%M")
                horas_trabalhadas = h_saida - h_chegada
                horas_totais = f"{horas_trabalhadas.seconds // 3600}:{(horas_trabalhadas.seconds // 60) % 60}"
                df.loc[df["Dia"] == dia_atual, "Horas Totais"] = horas_totais
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao calcular horas totais: {e}")
        else:
            df.loc[df["Dia"] == dia_atual, "Horas Totais"] = ""

        # Salvar de volta no Excel
        df.to_excel(EXCEL_FILE, index=False)

        messagebox.showinfo("Sucesso", f"{tipo} registrada com sucesso para o dia {dia_atual}!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao registrar horário: {e}")

def criar_interface():
    criar_planilha_se_nao_existir()

    janela = tk.Tk()
    janela.title("Ponto Virtual - Registro de Horários")

    # Botões para registrar horários
    botao_chegada = tk.Button(janela, text="Registrar Chegada", command=lambda: registrar_horario("Chegada"), width=20)
    botao_chegada.pack(pady=10)

    botao_almoco = tk.Button(janela, text="Registrar Almoço", command=lambda: registrar_horario("Almoco"), width=20)
    botao_almoco.pack(pady=10)

    botao_saida = tk.Button(janela, text="Registrar Saída", command=lambda: registrar_horario("Saida"), width=20)
    botao_saida.pack(pady=10)

    janela.mainloop()

if __name__ == "__main__":
    criar_interface()
