"""
Automação SEFIP - Robô de Processamento
Versão: 1.6
Data: 24/01/2026
Alterações:
- v1.6: Implementado sistema de "Tentar Novamente" ou "Pular" quando não localizar imagem
- v1.5: Implementado novo sistema de log CSV com formato inicio,competencia,valor,fim,status
- v1.4.1: Corrigido caminho das imagens para funcionar em diferentes máquinas
- v1.4: Implementado processamento em lote de múltiplos meses via CSV com valores variáveis
- v1.3: Adicionada etapa_0_limparbase para limpar base antes da importação
- v1.2: Alterações nas etapas 5, 6 e 7 (trabalhadores2.png, jfescrita2.png, press enter)
- v1.1: Inclusão da etapa 6 (adicionar_valor) e renumeração da etapa 7
- v1.0: Versão inicial com 7 etapas de automação
"""

import pyautogui
import time
import logging
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import threading
import os
import csv
import sys

log_dir = "logs"
if not os.path.exists(log_dir):
    os.makedirs(log_dir)

log_file = os.path.join(log_dir, f"execucao_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
logging.basicConfig(level=logging.INFO, filename=log_file, format='%(asctime)s - %(levelname)s - %(message)s')

log_csv_file = os.path.join(log_dir, f"processamento_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")

pyautogui.PAUSE = 1.5
pyautogui.FAILSAFE = True


def inicializar_log_csv():
    try:
        with open(log_csv_file, 'w', newline='', encoding='utf-8') as arquivo:
            escritor = csv.writer(arquivo)
            escritor.writerow(['inicio', 'competencia', 'valor', 'fim', 'status'])
    except Exception as e:
        logging.error(f"Erro ao criar log CSV: {str(e)}")


def registrar_processamento_csv(inicio, competencia, valor, fim, status):
    try:
        with open(log_csv_file, 'a', newline='', encoding='utf-8') as arquivo:
            escritor = csv.writer(arquivo)
            escritor.writerow([inicio, competencia, valor, fim, status])
    except Exception as e:
        logging.error(f"Erro ao registrar no log CSV: {str(e)}")


class ImagemNaoEncontradaException(Exception):
    pass


class RoboSEFIP:
    def __init__(self, ano, mes, valor, log_widget):
        self.ano = ano
        self.mes = mes
        self.valor = valor
        self.log_widget = log_widget
        self.inicio_processamento = None

        if getattr(sys, 'frozen', False):
            self.img_path = os.path.join(os.path.dirname(sys.executable), "imagens")
        else:
            self.img_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "imagens")

        if not os.path.exists(self.img_path):
            raise Exception(f"Pasta de imagens não encontrada: {self.img_path}")

        self.log_print(f"Pasta de imagens: {self.img_path}")

        self.caminho_arquivo = f"C:\\Robo_SEFIP\\{ano}\\{ano}{mes}\\SEFIP.RE"

    def log_print(self, mensagem):
        logging.info(mensagem)
        self.log_widget.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {mensagem}\n")
        self.log_widget.see(tk.END)

    def clicar_imagem(self, nome_imagem, conf=0.9, timeout=10, clicar=True, duplo=False, permitir_pular=True):
        caminho = os.path.join(self.img_path, nome_imagem)

        if not os.path.exists(caminho):
            self.log_print(f"ERRO: Arquivo não encontrado: {caminho}")
            raise Exception(f"Arquivo de imagem não encontrado: {caminho}")

        while True:
            start_time = time.time()

            while time.time() - start_time < timeout:
                try:
                    ponto = pyautogui.locateCenterOnScreen(caminho, confidence=conf)
                    if ponto:
                        if clicar:
                            if duplo:
                                pyautogui.doubleClick(ponto)
                            else:
                                pyautogui.click(ponto)
                        self.log_print(f"Sucesso: {nome_imagem}")
                        return ponto
                except Exception as e:
                    if "could not be found" in str(e).lower() or "não encontrado" in str(e).lower():
                        self.log_print(f"Erro ao localizar {nome_imagem}: {str(e)}")
                    pass
                time.sleep(0.7)

            self.log_print(f"⚠ Não encontrou {nome_imagem} na tela após {timeout}s")

            if not permitir_pular:
                raise Exception(f"Imagem não encontrada na tela: {nome_imagem}")

            resposta = self.perguntar_acao_imagem(nome_imagem)

            if resposta == "tentar":
                self.log_print(f"↻ Tentando novamente localizar: {nome_imagem}")
                continue
            elif resposta == "pular":
                self.log_print(f"⏭ Pulando imagem: {nome_imagem}")
                raise ImagemNaoEncontradaException(f"Usuário optou por pular: {nome_imagem}")
            else:
                self.log_print(f"✖ Processamento cancelado pelo usuário")
                raise Exception("Processamento cancelado pelo usuário")

    def perguntar_acao_imagem(self, nome_imagem):
        dialog = tk.Toplevel()
        dialog.title("Imagem Não Encontrada")
        dialog.geometry("450x200")
        dialog.transient()
        dialog.grab_set()

        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (450 // 2)
        y = (dialog.winfo_screenheight() // 2) - (200 // 2)
        dialog.geometry(f"450x200+{x}+{y}")

        resultado = {"acao": "cancelar"}

        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            frame,
            text=f"⚠ Imagem não encontrada: {nome_imagem}",
            font=('Arial', 11, 'bold'),
            foreground='red'
        ).pack(pady=10)

        ttk.Label(
            frame,
            text="O que deseja fazer?",
            font=('Arial', 10)
        ).pack(pady=5)

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=20)

        def escolher_tentar():
            resultado["acao"] = "tentar"
            dialog.destroy()

        def escolher_pular():
            resultado["acao"] = "pular"
            dialog.destroy()

        def escolher_cancelar():
            resultado["acao"] = "cancelar"
            dialog.destroy()

        ttk.Button(
            btn_frame,
            text="↻ Tentar Novamente",
            command=escolher_tentar,
            width=20
        ).grid(row=0, column=0, padx=5)

        ttk.Button(
            btn_frame,
            text="⏭ Pular Este Passo",
            command=escolher_pular,
            width=20
        ).grid(row=0, column=1, padx=5)

        ttk.Button(
            btn_frame,
            text="✖ Cancelar Tudo",
            command=escolher_cancelar,
            width=20
        ).grid(row=1, column=0, columnspan=2, pady=10)

        dialog.wait_window()

        return resultado["acao"]

    def etapa_0_limparbase(self):
        self.log_print("Iniciando Etapa 0: Limpar Base")
        self.clicar_imagem("ferramentas.png")
        self.clicar_imagem("limpar.png")
        pyautogui.press('enter')
        pyautogui.press('enter')

    def etapa_1_importar(self):
        self.log_print("Iniciando Etapa 1: Importação")
        self.clicar_imagem("arquivo.png")
        self.clicar_imagem("importar.png")
        time.sleep(1.5)
        pyautogui.write(self.caminho_arquivo)
        pyautogui.press('enter')
        #self.clicar_imagem("sim.png", timeout=25)
        self.clicar_imagem("ok.png", timeout=25)

    def etapa_2_remover_daniel(self):
        self.log_print("Iniciando Etapa 2: Remover Daniel")
        self.clicar_imagem("cadastro.png")
        self.clicar_imagem("jfescrita.png", duplo=True)
        self.clicar_imagem("daniel.png")
        self.clicar_imagem("excluir.png")
        self.clicar_imagem("sim.png")

    def etapa_3_cadastrar_daniel(self):
        self.log_print("Iniciando Etapa 3: Cadastrar Daniel")
        self.clicar_imagem("jfescrita.png")
        self.clicar_imagem("novotrabalhador.png")

        pyautogui.write("26979875262")
        pyautogui.press('tab')
        pyautogui.write("DANIEL ANGELO BRAGA")

        pyautogui.press('tab')
        pyautogui.write('11')
        pyautogui.press('enter')
        pyautogui.press('tab')
        pyautogui.write('01231')
        pyautogui.press('enter')
        pyautogui.press(['tab', 'tab'])
        pyautogui.press('down')
        pyautogui.press('tab')
        pyautogui.write('01042005')
        self.clicar_imagem("salvar.png")
        self.clicar_imagem("sim.png")

    def etapa_4_adicionardanielmodalidade1(self):
        self.log_print("Iniciando Etapa 4: Adicionar Daniel Modalidade 1")
        self.clicar_imagem("movimento.png")
        self.clicar_imagem("jfescrita.png")
        pyautogui.hotkey('ctrl', 'm')

        pos_destino = self.clicar_imagem("destino.png", clicar=False)
        pyautogui.click(pos_destino.x, pos_destino.y + 20)
        pyautogui.press('down')
        pyautogui.press('down')
        pyautogui.press('enter')

        self.clicar_imagem("setadupla.png")
        self.clicar_imagem("salvar.png")
        self.clicar_imagem("ok.png")

    def etapa_5_adicionardemaismodalidade9(self):
        self.log_print("Iniciando Etapa 5: Adicionar Demais Modalidade 9")
        pyautogui.hotkey('ctrl', 'm')
        self.clicar_imagem("trabalhadores2.png")
        pyautogui.press('down')
        pyautogui.press('enter')

        pos_destino = self.clicar_imagem("destino.png", clicar=False)
        pyautogui.click(pos_destino.x, pos_destino.y + 20)
        pyautogui.press('down')
        pyautogui.press('down')
        pyautogui.press('enter')

        self.clicar_imagem("setadupla.png")
        self.clicar_imagem("salvar.png")
        self.clicar_imagem("ok.png")

    def etapa_6_adicionar_valor(self):
        self.log_print("Iniciando Etapa 6: Adicionar Valor")
        self.clicar_imagem("jfescrita2.png", duplo=True)
        self.clicar_imagem("categoria1.png", duplo=True)
        self.clicar_imagem("daniel.png")
        self.clicar_imagem("marcarmovimento.png")
        self.clicar_imagem("dadosmovimento.png")
        pyautogui.write(str(self.valor))
        self.clicar_imagem("salvar.png")
        self.clicar_imagem("sim.png")

    def etapa_7_salvar_retificado(self):
        self.log_print("Iniciando Etapa 7: Salvar Retificado")
        self.clicar_imagem("115.png")
        self.clicar_imagem("executar.png")
        self.log_print("Aguardando confirmação (14 segundos)...")
        self.clicar_imagem("ok.png", timeout=20)
        self.clicar_imagem("salvar.png")
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.press('enter')


def carregar_csv(caminho_csv):
    dados = []
    try:
        with open(caminho_csv, 'r', encoding='utf-8') as arquivo:
            leitor = csv.DictReader(arquivo)
            for linha in leitor:
                dados.append({
                    'ano': linha['ano'].strip(),
                    'mes': linha['mes'].strip(),
                    'valor': linha['valor'].strip()
                })
        return dados
    except Exception as e:
        raise Exception(f"Erro ao carregar CSV: {str(e)}")


def iniciar_automacao_lote(caminho_csv, log_widget, progress_bar, status_label):
    def run():
        try:
            inicializar_log_csv()

            dados = carregar_csv(caminho_csv)
            total = len(dados)

            if total == 0:
                log_widget.insert(tk.END, "ERRO: Nenhum dado encontrado no CSV\n")
                return

            log_widget.insert(tk.END, f"=== Processamento em Lote Iniciado ===\n")
            log_widget.insert(tk.END, f"Total de meses a processar: {total}\n")
            log_widget.insert(tk.END, f"Log CSV: {log_csv_file}\n")
            log_widget.insert(tk.END, f"{'=' * 40}\n")

            for idx, item in enumerate(dados, 1):
                ano = item['ano']
                mes = item['mes']
                valor = item['valor']
                competencia = f"{mes}/{ano}"

                inicio = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

                log_widget.insert(tk.END, f"\n{'=' * 40}\n")
                log_widget.insert(tk.END, f"PROCESSANDO {idx}/{total}: {competencia} - Valor: R$ {valor}\n")
                log_widget.insert(tk.END, f"Início: {inicio}\n")
                log_widget.insert(tk.END, f"{'=' * 40}\n")
                log_widget.see(tk.END)

                progresso = (idx / total) * 100
                progress_bar['value'] = progresso
                status_label.config(text=f"Processando {idx}/{total}: {competencia}")

                status = "erro"

                try:
                    robo = RoboSEFIP(ano, mes, valor, log_widget)
                    robo.etapa_0_limparbase()
                    robo.etapa_1_importar()
                    robo.etapa_2_remover_daniel()
                    robo.etapa_3_cadastrar_daniel()
                    robo.etapa_4_adicionardanielmodalidade1()
                    robo.etapa_5_adicionardemaismodalidade9()
                    robo.etapa_6_adicionar_valor()
                    robo.etapa_7_salvar_retificado()

                    status = "sucesso"

                    fim = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

                    log_widget.insert(tk.END, f"✓ Mês {competencia} PROCESSADO COM SUCESSO!\n")
                    log_widget.insert(tk.END, f"Fim: {fim}\n")
                    logging.info(f"Mês {competencia} processado com sucesso")

                    registrar_processamento_csv(inicio, competencia, valor, fim, status)

                except ImagemNaoEncontradaException as e:
                    fim = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    status = "parcial - pulou etapa"

                    log_widget.insert(tk.END, f"⚠ Mês {competencia} PROCESSADO PARCIALMENTE (pulou etapa)\n")
                    log_widget.insert(tk.END, f"Fim: {fim}\n")
                    logging.warning(f"Mês {competencia} processado parcialmente: {str(e)}")

                    registrar_processamento_csv(inicio, competencia, valor, fim, status)

                    continuar = messagebox.askyesno(
                        "Etapa Pulada",
                        f"Uma etapa foi pulada em {competencia}.\n\nDeseja continuar para o próximo mês?"
                    )
                    if not continuar:
                        log_widget.insert(tk.END, "\nProcessamento interrompido pelo usuário.\n")
                        break

                except Exception as e:
                    fim = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

                    if "cancelado pelo usuário" in str(e).lower():
                        status = "cancelado"
                        log_widget.insert(tk.END, f"✖ Processamento CANCELADO pelo usuário\n")
                        log_widget.insert(tk.END, f"Fim: {fim}\n")
                        registrar_processamento_csv(inicio, competencia, valor, fim, status)
                        break

                    status = f"erro: {str(e)}"

                    log_widget.insert(tk.END, f"✗ ERRO no mês {competencia}: {str(e)}\n")
                    log_widget.insert(tk.END, f"Fim: {fim}\n")
                    logging.error(f"Erro no mês {competencia}: {str(e)}")

                    registrar_processamento_csv(inicio, competencia, valor, fim, status)

                    continuar = messagebox.askyesno(
                        "Erro no Processamento",
                        f"Erro ao processar {competencia}:\n{str(e)}\n\nDeseja continuar com os próximos meses?"
                    )
                    if not continuar:
                        log_widget.insert(tk.END, "\nProcessamento interrompido pelo usuário.\n")
                        break

            log_widget.insert(tk.END, f"\n{'=' * 40}\n")
            log_widget.insert(tk.END, "=== PROCESSAMENTO EM LOTE FINALIZADO ===\n")
            log_widget.insert(tk.END, f"Log detalhado salvo em: {log_csv_file}\n")
            log_widget.insert(tk.END, f"{'=' * 40}\n")
            status_label.config(text="Processamento Finalizado!")
            messagebox.showinfo("Concluído",
                                f"Processamento finalizado!\n{total} mês(es) processado(s).\n\nLog CSV: {log_csv_file}")

        except Exception as e:
            log_widget.insert(tk.END, f"\nERRO CRÍTICO: {str(e)}\n")
            logging.error(f"Erro crítico no processamento em lote: {str(e)}")
            messagebox.showerror("Erro Crítico", f"Erro crítico:\n{str(e)}")

        finally:
            progress_bar['value'] = 0
            status_label.config(text="Aguardando...")

    threading.Thread(target=run, daemon=True).start()


def criar_csv_modelo():
    caminho = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSV files", "*.csv")],
        initialfile="modelo_meses.csv"
    )

    if caminho:
        try:
            with open(caminho, 'w', newline='', encoding='utf-8') as arquivo:
                escritor = csv.writer(arquivo)
                escritor.writerow(['ano', 'mes', 'valor'])
                escritor.writerow(['2006', '01', '300'])
                escritor.writerow(['2006', '02', '350'])
                escritor.writerow(['2006', '03', '400'])

            messagebox.showinfo("Sucesso", f"Arquivo modelo criado:\n{caminho}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao criar arquivo:\n{str(e)}")


root = tk.Tk()
root.title("Automação SEFIP - Processamento em Lote")
root.geometry("700x550")

frame = ttk.Frame(root, padding="10")
frame.pack(fill=tk.BOTH, expand=True)

titulo = ttk.Label(frame, text="Automação SEFIP", font=('Arial', 14, 'bold'))
titulo.pack(pady=5)

frame_arquivo = ttk.LabelFrame(frame, text=""
                                           "CSV", padding=10)
frame_arquivo.pack(fill=tk.X, pady=10)

caminho_csv_var = tk.StringVar()


def selecionar_csv():
    caminho = filedialog.askopenfilename(
        title="Selecione o arquivo CSV",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    if caminho:
        caminho_csv_var.set(caminho)


ttk.Button(frame_arquivo, text="📁 CSV", command=selecionar_csv).pack(side=tk.LEFT, padx=5)
ttk.Button(frame_arquivo, text="📄 Criar Modelo CSV", command=criar_csv_modelo).pack(side=tk.LEFT, padx=5)

entry_csv = ttk.Entry(frame_arquivo, textvariable=caminho_csv_var, width=50, state='readonly')
entry_csv.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

frame_status = ttk.LabelFrame(frame, text="Status", padding=10)
frame_status.pack(fill=tk.X, pady=5)

status_label = ttk.Label(frame_status, text="Aguardando...", font=('Arial', 10))
status_label.pack()

progress_bar = ttk.Progressbar(frame_status, length=660, mode='determinate')
progress_bar.pack(pady=5)

btn_iniciar = ttk.Button(
    frame,
    text="▶ Iniciar Processamento em Lote",
    command=lambda: iniciar_automacao_lote(caminho_csv_var.get(), log_display, progress_bar, status_label)
)
btn_iniciar.pack(pady=10)

frame_log = ttk.LabelFrame(frame, text="Logs de Execução", padding=10)
frame_log.pack(fill=tk.BOTH, expand=True, pady=5)

log_display = scrolledtext.ScrolledText(frame_log, height=15, width=80)
log_display.pack(fill=tk.BOTH, expand=True)

instrucoes = """
INSTRUÇÕES:
1. Clique em "Criar Modelo CSV" para gerar um arquivo exemplo
2. Edite o CSV com seus dados (ano, mes, valor)
3. Clique em "Selecionar CSV" para escolher o arquivo
4. Clique em "Iniciar Processamento em Lote"

Formato do CSV:
ano,mes,valor
2006,01,300
2006,02,350
2006,03,400
"""

ttk.Label(frame, text=instrucoes, justify=tk.LEFT, font=('Courier', 8)).pack(pady=5)

root.mainloop()