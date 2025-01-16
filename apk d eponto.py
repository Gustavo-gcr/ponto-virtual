import tkinter as tk
from tkinter import messagebox
from datetime import datetime, timedelta
import pandas as pd
import os
from fpdf import FPDF

# Nome do arquivo Excel
EXCEL_FILE = "registros.xlsx"

# Criar o Excel inicial se não existir
if not os.path.exists(EXCEL_FILE):
    colunas = ["Dia", "Chegada", "Almoco", "Saida", "Horas Totais", "OBS"]
    dias = list(range(1, 32))
    df = pd.DataFrame({"Dia": dias, "Chegada": ["" for _ in dias], "Almoco": ["" for _ in dias],
                       "Saida": ["" for _ in dias], "Horas Totais": ["" for _ in dias], "OBS": ["" for _ in dias]})
    df.to_excel(EXCEL_FILE, index=False)

def registrar_horario(tipo):
    try:
        dia_selecionado = int(entry_dia.get())
        horario_atual = datetime.now().strftime("%H:%M")

        # Verificar se o dia é válido
        if dia_selecionado < 1 or dia_selecionado > 31:
            messagebox.showerror("Erro", "Dia inválido. Por favor, insira um valor entre 1 e 31.")
            return

        # Carregar o Excel existente
        df = pd.read_excel(EXCEL_FILE)

        # Atualizar o horário correspondente
        if tipo == "Chegada":
            df.loc[df["Dia"] == dia_selecionado, "Chegada"] = horario_atual
        elif tipo == "Almoco":
            df.loc[df["Dia"] == dia_selecionado, "Almoco"] = horario_atual
        elif tipo == "Saida":
            df.loc[df["Dia"] == dia_selecionado, "Saida"] = horario_atual

        # Calcular horas totais se todos os horários forem válidos
        chegada = df.loc[df["Dia"] == dia_selecionado, "Chegada"].values[0]
        almoco = df.loc[df["Dia"] == dia_selecionado, "Almoco"].values[0]
        saida = df.loc[df["Dia"] == dia_selecionado, "Saida"].values[0]

        if pd.notna(chegada) and pd.notna(almoco) and pd.notna(saida):
            try:
                h_chegada = datetime.strptime(chegada, "%H:%M")
                h_almoco = datetime.strptime(almoco, "%H:%M")
                h_saida = datetime.strptime(saida, "%H:%M")
                horas_trabalhadas = (h_almoco - h_chegada) + (h_saida - h_almoco)
                horas_totais = str(horas_trabalhadas)
                df.loc[df["Dia"] == dia_selecionado, "Horas Totais"] = horas_totais
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao calcular horas totais: {e}")

        # Salvar de volta no Excel
        df.to_excel(EXCEL_FILE, index=False)

        messagebox.showinfo("Sucesso", f"{tipo} registrada com sucesso para o dia {dia_selecionado}!")
    except ValueError:
        messagebox.showerror("Erro", "Por favor, insira um dia válido.")

def criar_interface():
    global entry_dia

    janela = tk.Tk()
    janela.title("Ponto Virtual - Registro de Horários")

    # Tabela visual com dias
    tabela_frame = tk.Frame(janela)
    tabela_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

    # Cabeçalho
    headers = ["Dia", "Chegada", "Almoco", "Saida", "Horas Totais", "OBS"]
    for col, header in enumerate(headers):
        tk.Label(tabela_frame, text=header, borderwidth=1, relief="solid", width=15).grid(row=0, column=col)

    # Dias do mês
    for dia in range(1, 32):
        tk.Label(tabela_frame, text=str(dia), borderwidth=1, relief="solid", width=15).grid(row=dia, column=0)

    # Inputs para registro
    tk.Label(janela, text="Dia:").grid(row=1, column=0, padx=5, pady=5)
    entry_dia = tk.Entry(janela)
    entry_dia.grid(row=1, column=1, padx=5, pady=5)

    # Botões para registrar horários
    botao_chegada = tk.Button(janela, text="Registrar Chegada", command=lambda: registrar_horario("Chegada"), width=20)
    botao_chegada.grid(row=2, column=0, columnspan=2, pady=5)

    botao_almoco = tk.Button(janela, text="Registrar Almoço", command=lambda: registrar_horario("Almoco"), width=20)
    botao_almoco.grid(row=3, column=0, columnspan=2, pady=5)

    botao_saida = tk.Button(janela, text="Registrar Saída", command=lambda: registrar_horario("Saida"), width=20)
    botao_saida.grid(row=4, column=0, columnspan=2, pady=5)

    # Botão para gerar relatório
  #  botao_relatorio = tk.Button(janela, text="Gerar Relatório", command=gerar_relatorio, width=20)
   # botao_relatorio.grid(row=5, column=0, columnspan=2, pady=10)

    janela.mainloop()

if __name__ == "__main__":
    criar_interface()
