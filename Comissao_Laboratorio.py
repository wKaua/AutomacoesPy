import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, Label
import pandas as pd
import os
from datetime import datetime
import win32com.client as win32
import threading
from datetime import datetime

data_atual = datetime.now().strftime("%d-%m-%Y")

# Variáveis globais para armazenar os caminhos dos arquivos
caminho_arquivo1 = ""
caminho_arquivo2 = ""
caminho_arquivo3 = ""

class PlaceholderEntry(tk.Entry):
    def __init__(self, master=None, placeholder="", **kwargs):
        super().__init__(master, **kwargs)
        self.placeholder = placeholder
        self.insert(0, placeholder)
        self.config(fg='gray', justify='left')
        self.bind("<FocusIn>", self.remove_placeholder)
        self.bind("<FocusOut>", self.add_placeholder)

    def remove_placeholder(self, event):
        if self.get() == self.placeholder:
            self.delete(0, tk.END)
            self.config(fg='black')

    def add_placeholder(self, event):
        if not self.get():
            self.insert(0, self.placeholder)
            self.config(fg='gray')

def saudacao_por_hora():
    hora_atual = datetime.now().hour
    if 6 <= hora_atual < 12:
        return "bom dia!"
    elif 12 <= hora_atual <18:
        return "boa tarde!"
    else:
        return "boa noite"

def texto_email(arquivos_selecionados):
    quantidade_arquivos_selecionados = len(arquivos_selecionados)
    
    if quantidade_arquivos_selecionados == 1:
        trecho_texto = "Segue em anexo o relatório detalhado da comissão do laboratório."
    else:
        trecho_texto = "Seguem em anexo os relatórios detalhados da comissão dos laboratórios."
    
    # Usando tags <p> para parágrafos
    return f"""
    <p>Vânia, {saudacao_por_hora()}</p>
    <p>{trecho_texto}</p>
    <p>At.te,</p>
    """


def selecionar_arquivo(entry):
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione um arquivo",
        filetypes=[("Arquivos Excel", "*.xlsx;*.xls")]
    )
    if caminho_arquivo:
        nome_arquivo = os.path.basename(caminho_arquivo)
        if entry == entry_arquivo1:
            global caminho_arquivo1
            caminho_arquivo1 = caminho_arquivo
        elif entry == entry_arquivo2:
            global caminho_arquivo2
            caminho_arquivo2 = caminho_arquivo
        elif entry == entry_arquivo3:
            global caminho_arquivo3
            caminho_arquivo3 = caminho_arquivo

        entry.config(state=tk.NORMAL)
        entry.delete(0, tk.END)
        entry.insert(0, nome_arquivo)
        entry.config(fg='black')

def on_entry_click(event):
    selecionar_arquivo(event.widget)

def fechar():
    root.destroy()

def atualizar_entrada_email():
    if var_usar_email.get():
        entry_email.config(fg='black')
        entry_email.delete(0, tk.END)
        entry_email.insert(0, "vania.santos@laboratoriosodre.com.br")

def atualizar_entrada_assunto():
    if var_usar_assunto.get():
        entry_assunto.config(fg='black')
        entry_assunto.delete(0, tk.END)
        entry_assunto.insert(0, "Comissão Laboratórios")
        
def enviar_email(arquivos, destinatario, assunto):
    if not destinatario or not assunto:
        messagebox.showwarning("Aviso", "Certifique-se de preencher o e-mail e o assunto do e-mail.")
        return
    
    try:
        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)
        email.To = destinatario
        email.Subject = assunto
        email.HTMLBody = texto_email(arquivos)  
        for arquivo in arquivos:
            email.Attachments.Add(arquivo)
        email.Send()
        print("Email Enviado")
        
        if len(arquivos) == 1:
            messagebox.showinfo("Sucesso", "O arquivo foi enviado com sucesso!")
        else:
            messagebox.showinfo("Sucesso", "Os arquivos foram enviados com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", f'Ocorreu um erro ao enviar o e-mail: {e}')
        print(f'Erro ao enviar e-mail: {e}')


def processar_arquivos(arquivos):
    resultados = []
    arquivos_para_envio = []
    try:
        for caminho, valor_amostra, percentual in arquivos:
            if caminho:
                nome_arquivo = os.path.basename(caminho)
                nome_arquivo_novo = nome_arquivo.replace(".xlsx", f" {data_atual}.xlsx")
                caminho_novo_arquivo = os.path.join(os.path.dirname(caminho), nome_arquivo_novo)
                
                df = pd.read_excel(caminho)
                
                # Agrupando por 'FANTASIA_LAB'
                grupos = df.groupby('FANTASIA_LAB')
                
                resumo_por_unidade = []
                for unidade, dados in grupos:
                    quantidade_etiquetas = len(dados)
                    excluidas = (dados['Excluido'] == 1).sum()
                    canceladas_excluida = (dados['Cancelado']).sum()
                    Filtro_Validas_Pag_Aberto = (
                        (dados['StatusPagamento']=='Em Aberto') &
                        (dados['Cancelado']==0) &
                        (dados['Excluido']==0) &
                        (dados['DataCancelado'].isnull())
                    )
                    Validas_Pag_Aberto = Filtro_Validas_Pag_Aberto.sum()
                    total_validas = quantidade_etiquetas - (excluidas + canceladas_excluida)
                    
                    if percentual is not None:
                        total_receber = (total_validas * valor_amostra) * percentual
                    else:
                        total_receber = valor_amostra * total_validas
                    
                    format_total_receber = f"{total_receber:,}".replace(',', 'X').replace('.', ',').replace('X', '.')
                    format_total_receber = f"{format_total_receber},00"

                    resumo_unidade = {
                        "Unidade": unidade,
                        "Total Etiquetas": quantidade_etiquetas,
                        "Etiquetas Excluídas": excluidas,
                        "Etiquetas Canceladas/Excluídas": canceladas_excluida,
                        "Etiquetas Válidas com Pagamentos em Aberto": Validas_Pag_Aberto,
                        "Total Etiquetas Válidas": total_validas,
                        "Valor por Amostras": valor_amostra,
                        "Total a Receber": format_total_receber
                    }
                    resumo_por_unidade.append(resumo_unidade)
                
                # Resumo geral
                quantidade_etiquetas_total = df.shape[0]
                excluidas_total = (df['Excluido'] == 1).sum()
                canceladas_excluida_total = (df['Cancelado']).sum()
                Filtro_Validas_Pag_Aberto_total = (
                    (df['StatusPagamento']=='Em Aberto') &
                    (df['Cancelado']==0) &
                    (df['Excluido']==0) &
                    (df['DataCancelado'].isnull())
                )
                Validas_Pag_Aberto_total = Filtro_Validas_Pag_Aberto_total.sum()
                total_validas_total = quantidade_etiquetas_total - (excluidas_total + canceladas_excluida_total)
                
                if percentual is not None:
                    total_receber_total = (total_validas_total * valor_amostra) * percentual
                else:
                    total_receber_total = valor_amostra * total_validas_total
                
                format_total_receber_total = f"{total_receber_total:,}".replace(',', 'X').replace('.', ',').replace('X', '.')
                format_total_receber_total = f"{format_total_receber_total},00"

                resumo_geral = {
                    "Unidade": "Total",
                    "Total Etiquetas": quantidade_etiquetas_total,
                    "Etiquetas Excluídas": excluidas_total,
                    "Etiquetas Canceladas/Excluídas": canceladas_excluida_total,
                    "Etiquetas Válidas com Pagamentos em Aberto": Validas_Pag_Aberto_total,
                    "Total Etiquetas Válidas": total_validas_total,
                    "Valor por Amostras": valor_amostra,
                    "Total a Receber": format_total_receber_total
                }
                
                # Gerar o resumo em uma nova aba
                df_resumo_unidades = pd.DataFrame(resumo_por_unidade)
                df_resumo_geral = pd.DataFrame([resumo_geral])
                
                with pd.ExcelWriter(caminho_novo_arquivo, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name='Original')
                    df_resumo_unidades.to_excel(writer, sheet_name="Resumo por Unidade", index=False)
                    df_resumo_geral.to_excel(writer, sheet_name="Resumo Geral", index=False)
                arquivos_para_envio.append(caminho_novo_arquivo)

        return resultados, arquivos_para_envio
    except PermissionError:
        messagebox.showerror("Erro", "Permissão negada. Certifique-se de que os arquivos não estão abertos em outro programa e tente novamente.")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
        print(f'Ocorreu um erro: {e}')
    return resultados, arquivos_para_envio

def processar_e_continuar():
    global caminho_arquivo1, caminho_arquivo2, caminho_arquivo3
    try:
        arquivos = [
            (caminho_arquivo1, 25, None), 
            (caminho_arquivo2, 25, None), 
            (caminho_arquivo3, 60, 0.04)
        ]
        texto_email(arquivos)
        print("Iniciando processamento de arquivos...")
        resultados, arquivos_para_envio = processar_arquivos(arquivos)
        print(f"Resultados: {resultados}")
        print(f"Arquivos para envio: {arquivos_para_envio}")
        if arquivos_para_envio:
            destinatario = entry_email.get().strip()
            assunto = entry_assunto.get().strip()
            enviar_email(arquivos_para_envio, destinatario, assunto)
    except Exception as e:
        print(f"Ocorreu um erro durante o processamento: {e}")
    finally:
        loading_window.destroy()
        fechar()

def continuar():
    
    global caminho_arquivo1, caminho_arquivo2, caminho_arquivo3
    
    # Filtrar arquivos não selecionados
    arquivos = [
        (caminho_arquivo1, 25, None), 
        (caminho_arquivo2, 25, None), 
        (caminho_arquivo3, 60, 0.04)
    ]
    
    # Filtra apenas os arquivos que possuem caminho
    arquivos_selecionados = [arquivo for arquivo in arquivos if arquivo[0]]

    if arquivos_selecionados:
        # Cria a janela de carregamento
        global loading_window
        loading_window = Toplevel(root)
        loading_window.title("Processando")
        Label(loading_window, text="Processando, por favor aguarde...").pack(padx=20, pady=20)
        loading_window.geometry("300x100")
        loading_window.transient(root)
        loading_window.grab_set()
        root.update()

        # Inicia a thread
        threading.Thread(target=lambda: processar_e_continuar(arquivos_selecionados)).start()
    else:
        messagebox.showwarning("Aviso", "Por favor, selecione ao menos um arquivo.")

def processar_e_continuar(arquivos):
    try:
        texto_email(arquivos)
        print("Iniciando processamento de arquivos...")
        resultados, arquivos_para_envio = processar_arquivos(arquivos)
        print(f"Resultados: {resultados}")
        print(f"Arquivos para envio: {arquivos_para_envio}")
        if arquivos_para_envio:
            destinatario = entry_email.get().strip()
            assunto = entry_assunto.get().strip()
            enviar_email(arquivos_para_envio, destinatario, assunto)
    except Exception as e:
        print(f"Ocorreu um erro durante o processamento: {e}")
    finally:
        loading_window.destroy()
        fechar()

root = tk.Tk()
root.title("Selecione os Arquivos Excel e Envie o E-Mail")

frame_arquivos = tk.LabelFrame(root, text="Seleção de Arquivos")
frame_arquivos.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

frame_arquivos.columnconfigure(1, weight=1)
frame_arquivos.rowconfigure(0, weight=1)

label_arquivo1 = tk.Label(frame_arquivos, text="Selecione o Arquivo Acesso Saúde:", anchor='w')
label_arquivo1.grid(row=0, column=0, padx=10, pady=5, sticky="w")

entry_arquivo1 = PlaceholderEntry(frame_arquivos, placeholder="Clique aqui para selecionar o arquivo Acesso Saúde", width=55)
entry_arquivo1.grid(row=0, column=1, padx=10, pady=5)
entry_arquivo1.bind("<Button-1>", on_entry_click)

label_arquivo2 = tk.Label(frame_arquivos, text="Selecione o Arquivo Cipax:", anchor='w')
label_arquivo2.grid(row=1, column=0, padx=10, pady=5, sticky="w")

entry_arquivo2 = PlaceholderEntry(frame_arquivos, placeholder="Clique aqui para selecionar o arquivo Cipax", width=55)
entry_arquivo2.grid(row=1, column=1, padx=10, pady=5)
entry_arquivo2.bind("<Button-1>", on_entry_click)

label_arquivo3 = tk.Label(frame_arquivos, text="Selecione o Arquivo São João - Rondonópolis:", anchor='w')
label_arquivo3.grid(row=2, column=0, padx=10, pady=5, sticky="w")

entry_arquivo3 = PlaceholderEntry(frame_arquivos, placeholder="Clique aqui para selecionar o arquivo São João - Rondonópolis", width=55)
entry_arquivo3.grid(row=2, column=1, padx=10, pady=5)
entry_arquivo3.bind("<Button-1>", on_entry_click)

frame_email = tk.LabelFrame(root, text="Envio de E-Mail")
frame_email.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

label_email = tk.Label(frame_email, text="E-mail do Destinatário:")
label_email.grid(row=0, column=0, padx=10, pady=5, sticky="w")

entry_email = PlaceholderEntry(frame_email, placeholder="vania.santos@laboratoriosodre.com.br", width=50)
entry_email.grid(row=0, column=1, padx=10, pady=5)

label_assunto = tk.Label(frame_email, text="Assunto do E-Mail:")
label_assunto.grid(row=1, column=0, padx=10, pady=5, sticky="w")

entry_assunto = PlaceholderEntry(frame_email, placeholder="Comissão Laboratórios", width=50)
entry_assunto.grid(row=1, column=1, padx=10, pady=5)

var_usar_email = tk.BooleanVar(value=False)
check_usar_email = tk.Checkbutton(frame_email, text="Usar e-mail pré-preenchido", variable=var_usar_email, command=atualizar_entrada_email)
check_usar_email.grid(row=0, column=2, columnspan=2, padx=10, pady=5, sticky="w")

var_usar_assunto = tk.BooleanVar(value=False)
check_usar_assunto = tk.Checkbutton(frame_email, text="Usar assunto pré-preenchido", variable=var_usar_assunto, command=atualizar_entrada_assunto)
check_usar_assunto.grid(row=1, column=2, columnspan=2, padx=10, pady=5, sticky="w")

frame_botoes = tk.Frame(root)
frame_botoes.grid(row=2, column=0, padx=10, pady=10, sticky="ew")

btn_fechar = tk.Button(root, text="Fechar", command=fechar)
btn_fechar.grid(row=2, column=0, padx=10, pady=10, sticky="w")

btn_continuar = tk.Button(root, text="Continuar", command=continuar)
btn_continuar.grid(row=2, column=0, padx=10, pady=10, sticky="e")

root.mainloop()
