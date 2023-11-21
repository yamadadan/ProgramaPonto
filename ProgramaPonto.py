
# import tkinter as tk
from tkinter import Tk,font, GROOVE, Button
# from ttkthemes import ThemedTk
from funcoes import iniciar_arrastar, arrastar_janela, marcar_ponto
import cProfile
import pstats
# from openpyxl.styles import NamedStyle

# Criar a janela principal
janela = Tk()
# janela = ThemedTk(theme="arc")  # Você pode escolher um tema diferente se preferir
janela.title("Registro de Ponto")

# Criar uma instância da classe Font e configurar as propriedades desejadas
estilo_fonte = font.Font(family="Helvetica", size=12, weight="bold")
# Criar um botão na janela
botao_registrar_ponto = Button(janela, text="Registrar Ponto", command=lambda: marcar_ponto(janela), bg="#FF4021", fg="white", relief=GROOVE, width=15, height=2, font=estilo_fonte)
botao_registrar_ponto.pack(pady=20)

# Permitir que a janela seja movida arrastando a barra superior
janela.bind("<ButtonPress-1>", iniciar_arrastar)
janela.bind("<B1-Motion>", lambda event: arrastar_janela(event, janela))

largura_janela = 200
altura_janela = 100

largura_tela = janela.winfo_screenwidth()
altura_tela = janela.winfo_screenheight()

posicao_x = (largura_tela - largura_janela) // 2
posicao_y = (altura_tela - altura_janela) // 2

janela.geometry(f"{largura_janela}x{altura_janela}+{posicao_x}+{posicao_y}")

# Criando um objeto cProfile
profiler = cProfile.Profile()

# Iniciando o profiling
profiler.enable()


# Iniciar o loop principal da interface gráfica
janela.mainloop()

# Parando o profiling
profiler.disable()

# Criando as estatísticas
stats = pstats.Stats(profiler)

# Imprimindo as estatísticas ordenadas pelo tempo crescente
stats.sort_stats('cumulative').print_stats()
# Imprimindo as estatísticas
# profiler.print_stats(sort='time')