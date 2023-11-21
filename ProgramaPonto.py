
from tkinter import Tk,font, GROOVE, Button, Menu
from funcoes import iniciar_arrastar, arrastar_janela, marcar_ponto, selecionar_diretorio, carregar_diretorio_salvo, modificar_diretorio

# Criar a janela principal
janela = Tk()
# janela = ThemedTk(theme="arc")  # Você pode escolher um tema diferente se preferir
janela.title("Registro de Ponto")

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

# Barra de opções
barra_opcoes = Menu(janela)

# Menu 'Arquivo'
menu_arquivo = Menu(barra_opcoes, tearoff=0)
menu_arquivo.add_command(label="Modificar Diretorio Salvo", command=lambda: modificar_diretorio())
menu_arquivo.add_separator()
menu_arquivo.add_command(label="Sair", command=janela.destroy)

# Adicionar o menu 'Arquivo' à barra de opções
barra_opcoes.add_cascade(label="Arquivo", menu=menu_arquivo)

# Configurar a barra de opções na janela
janela.config(menu=barra_opcoes)

# Tenta carregar o diretório salvo
diretorio_salvo = carregar_diretorio_salvo()
if diretorio_salvo:
  nome_arquivo = diretorio_salvo+'/ponto.xlsx'
  pass
else:
  # Pede ao usuário que selecione um diretório
  diretorio_salvo = selecionar_diretorio()
  nome_arquivo = diretorio_salvo+'/ponto.xlsx'
# Criar uma instância da classe Font e configurar as propriedades desejadas
estilo_fonte = font.Font(family="Helvetica", size=12, weight="bold")
# Criar um botão na janela
botao_registrar_ponto = Button(janela, text="Registrar Ponto", command=lambda: marcar_ponto(janela, nome_arquivo), bg="#FF4021", fg="white", relief=GROOVE, width=15, height=2, font=estilo_fonte)
botao_registrar_ponto.pack(pady=20)

# Iniciar o loop principal da interface gráfica
janela.mainloop()