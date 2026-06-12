import re
import customtkinter as ctk
from tkinter import filedialog
import xlwings as xw
import os
import threading
import queue
import sys
import traceback
import time

# Constantes
COR_AZUL = 0x0000FF
ALTURA_SETA_PX = 80
LARGURA_SETA_PX = 120
ESPESSURA_LINHA = 1.5
TAMANHO_FONTE = 10
LINHA_MAX_FALLBACK = 2000

# Configurações por aba do Excel.
# Cada chave é o nome exato da aba. Para adicionar suporte a uma nova aba,
# basta incluir uma nova entrada aqui — sem alterar o código de processamento.
#
# Campos obrigatórios:
#   modelo_RF     — controla o posicionamento das setas ("RJ/NI" ou "Demais_RFs")
#   coluna_codigo — coluna onde os códigos de item são buscados (ex: "P", "L")
#   coluna_dados  — coluna onde os valores de medida estão (ex: "O", "K")
#   offset_comp   — quantas linhas abaixo do código está o comprimento
#   offset_alt    — quantas linhas abaixo do código está a altura
#   offset_larg   — quantas linhas abaixo do código está a largura (None = não existe)
#   codigos       — lista de códigos que disparam a geração de cota nessa aba
CONFIGURACOES_ABAS = {
    "RELATORIO": {
        "modelo_RF": "RJ/NI",
        "coluna_codigo": "P",
        "coluna_dados": "O",
        "offset_comp": 2,   # comprimento em O(n+2)
        "offset_alt": 3,    # altura em O(n+3)
        "offset_larg": 5,   # largura em O(n+5)
        "codigos": ["17.1", "17.3", "17.4", "17.6", "17.7", "17.8", "17.10", "17.11", "29.2", "29.7"],
    },
    "RELATÓRIO GERAL": {
        "modelo_RF": "Demais_RFs",
        "coluna_codigo": "L",
        "coluna_dados": "K",
        "offset_comp": 3,    # valor inicial; sobrescrito pela detecção automática
        "offset_alt": 4,     # valor inicial; sobrescrito pela detecção automática
        "offset_larg": None, # largura não existe nesse modelo
        # Pares (offset_comp, offset_alt) testados em ordem até um produzir valor numérico.
        # Layout 3/4: comprimento em K(n+3), altura em K(n+4)
        # Layout 1/2: comprimento em K(n+1), altura em K(n+2)
        "offset_candidatos": [(3, 4), (1, 2)],
        "codigos": [
            "17.1 CORRIMÃO", "17.1 ESCUDO", "17.1 PIQUETE", "17.1 BICICLETÁRIO",
            "17.3 PAREDE", "17.4 PAREDE", "17.4 MARQUISE", "17.4 FORRO", "17.6 PAREDE",
            "17.6 RODAPÉ", "17.6 PILAR", "17.6 MURETA", "17.6 MURO", "17.6 MARQUISE",
            "17.6 FORRO", "17.7 PORTA", "17.8 VAGAS", "17.8 TÁTIL", "17.9 LETREIRO",
            "17.9 TOTEM", "17.9 LIXEIRA", "17.11 PAREDE", "17.11 RODAPÉ", "17.11 PILAR",
            "17.11 MURETA", "17.11  MURO", "17.11 MARQUISE", "17.11  FORRO", "29.7"
        ],
    },
}

# Configuração usada quando a aba encontrada não tem entrada em CONFIGURACOES_ABAS.
# Espelha o layout da aba RELATORIO com coluna de dados ajustada para S.
CONFIGURACAO_PADRAO = {
    "modelo_RF": "RJ/NI",
    "coluna_codigo": "P",
    "coluna_dados": "S",
    "offset_comp": 2,
    "offset_alt": 3,
    "offset_larg": None,
    "codigos": ["17.1", "17.3", "17.4", "17.6", "17.7", "17.8", "17.10", "17.11", "29.2", "29.7"],
}

# Configuração inicial do tema do CustomTkinter
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# Variáveis globais para threading e logging
log_queue = queue.Queue()
debug_text_widget = None
stop_event = threading.Event()  # sinaliza para a thread interromper o processamento

class DebugCapture:
    """Classe para capturar prints e redirecionar para a interface"""
    def __init__(self, queue_obj):
        self.queue = queue_obj
        self.original_stdout = sys.__stdout__

    def write(self, text):
        if text.strip():
            try:
                self.queue.put(text.strip())
            except Exception as e:
                if self.original_stdout:
                    self.original_stdout.write(f"\nErro ao adicionar à queue: {e}\n")

        if self.original_stdout:
            try:
                self.original_stdout.write(text)
            except Exception:
                pass

    def flush(self):
        if self.original_stdout and hasattr(self.original_stdout, 'flush'):
            try:
                self.original_stdout.flush()
            except Exception:
                pass

def log_debug(message):
    """Função para adicionar mensagens de debug"""
    if debug_text_widget:
        try:
            log_queue.put(message)
        except Exception:
            print(message)
    else:
        print(message)

def update_debug_display():
    """Atualiza a caixa de texto com as mensagens de debug"""
    try:
        while True:
            message = log_queue.get_nowait()
            if debug_text_widget:
                debug_text_widget.insert("end", message + "\n")
                debug_text_widget.see("end")
    except queue.Empty:
        pass

    if debug_text_widget:
        app.after(100, update_debug_display)

def selecionar_arquivo():
    caminho_arquivo = filedialog.askopenfilename(
        filetypes=[
            ("Arquivos Excel", "*.xlsx"),
            ("Arquivos Excel com Macro", "*.xlsm"),
            ("Todos os arquivos", "*.*")
        ]
    )

    if caminho_arquivo and not caminho_arquivo.endswith(('.xlsx', '.xlsm')):
        status_label.configure(text="Erro: O arquivo selecionado não é um Excel válido.", text_color="red")
        return

    if caminho_arquivo:
        entrada_caminho.delete(0, 'end')
        entrada_caminho.insert(0, caminho_arquivo)
        status_label.configure(text="Arquivo selecionado com sucesso!", text_color="green")
    else:
        status_label.configure(text="Erro: Nenhum arquivo selecionado.", text_color="red")

def obter_ultima_linha_area_impressao(sht):
    """
    Identifica a última linha da área de impressão da planilha
    """
    try:
        print_area = sht.api.PageSetup.PrintArea
        if print_area:
            print(f"Área de impressão definida: {print_area}")
            match = re.search(r'\$[A-Z]+\$(\d+)$', print_area)
            if match:
                ultima_linha_impressao = int(match.group(1))
                print(f"Última linha da área de impressão: {ultima_linha_impressao}")
                return ultima_linha_impressao

        print("Área de impressão não definida, procurando última linha com dados...")

        used_range = sht.used_range
        if used_range:
            ultima_linha_dados = used_range.last_cell.row
            print(f"Última linha com dados: {ultima_linha_dados}")
            return ultima_linha_dados

        print("Procurando última linha nas colunas específicas...")

        colunas_verificar = ['P', 'L', 'O', 'K', 'S']
        ultima_linha_encontrada = 0

        for coluna in colunas_verificar:
            try:
                for linha in range(LINHA_MAX_FALLBACK, 0, -1):
                    valor = sht.range(f"{coluna}{linha}").value
                    if valor is not None and str(valor).strip():
                        print(f"Última linha com dados na coluna {coluna}: {linha}")
                        ultima_linha_encontrada = max(ultima_linha_encontrada, linha)
                        break
            except Exception as e:
                print(f"\nErro ao acessar coluna {coluna}: {e}")
                print(f"\n{traceback.format_exc()}")
                continue

        if ultima_linha_encontrada > 0:
            return ultima_linha_encontrada

        print(f"Não foi possível determinar a última linha, usando {LINHA_MAX_FALLBACK} como padrão")
        return LINHA_MAX_FALLBACK

    except Exception as e:
        print(f"\nErro ao obter última linha da área de impressão: {e}")
        print(f"\n{traceback.format_exc()}")
        return LINHA_MAX_FALLBACK

def detectar_offsets(sht, coluna_codigo, coluna_dados, codigos, ultima_linha, candidatos):
    """
    Detecta automaticamente qual par (offset_comp, offset_alt) está em uso na planilha.

    Estratégia: encontra o primeiro código reconhecido e testa cada candidato
    verificando se a célula de dados no offset correspondente contém um valor numérico.
    O primeiro candidato que produzir um número válido é retornado.

    Retorna (offset_comp, offset_alt) se detectado, ou None se:
      - Nenhum código for encontrado na planilha, ou
      - Nenhum candidato produzir valor numérico (layout desconhecido).
    """
    for linha in range(4, ultima_linha + 1):
        valor = sht.range(f"{coluna_codigo}{linha}").value
        if not valor or str(valor).strip().upper() not in codigos:
            continue

        print(f"Detectando offsets a partir do código '{str(valor).strip()}' na linha {linha}...")
        for offset_comp, offset_alt in candidatos:
            for offset in (offset_comp, offset_alt):
                cell_val = sht.range(f"{coluna_dados}{linha + offset}").value
                try:
                    if cell_val is not None:
                        float(cell_val)
                        print(f"Layout detectado: offset_comp={offset_comp}, offset_alt={offset_alt} "
                              f"(valor numérico em {coluna_dados}{linha + offset})")
                        return offset_comp, offset_alt
                except (ValueError, TypeError):
                    continue

        print(f"Código encontrado na linha {linha}, mas nenhum candidato produziu valor numérico. "
              f"Layout não determinado.")
        return None

    print("Nenhum código reconhecido encontrado na planilha. Não foi possível detectar o layout.")
    return None


def _fechar_e_reabrir(wb, app_excel, caminho, salvar=True):
    """Fecha (e opcionalmente salva) o workbook e o reabre na mesma instância do Excel"""
    if wb is not None:
        if salvar:
            wb.save()
            print("Arquivo salvo com sucesso!")
            time.sleep(1)
        wb.close()
        print("Workbook fechado com sucesso!")

    if app_excel is not None:
        print("Reabrindo arquivo na mesma instância do Excel...")
        time.sleep(1)
        app_excel.books.open(caminho)
        print(f"Arquivo {caminho} reaberto com sucesso!")

def gerar_cotas_thread():
    """Função que executa o processamento em thread separada"""
    original_stdout = sys.stdout
    wb = None
    app_excel = None

    try:
        sys.stdout = DebugCapture(log_queue)

        caminho = entrada_caminho.get().strip()
        if caminho:
            if os.path.exists(caminho):
                try:
                    wb = xw.Book(caminho)
                    app_excel = wb.app

                    abas_disponiveis = [sheet.name for sheet in wb.sheets]
                    print(f"Abas disponíveis: {abas_disponiveis}")

                    aba_relatorio = None
                    nomes_possiveis = ["RELATORIO", "RELATÓRIO", "Relatorio", "Relatório", "RELATÓRIO GERAL", "RELATORIO GERAL"]

                    for nome in nomes_possiveis:
                        if nome in abas_disponiveis:
                            aba_relatorio = nome
                            break

                    if aba_relatorio is None:
                        for aba in abas_disponiveis:
                            if "RELATORIO" in aba.upper() or "RELATÓRIO" in aba.upper():
                                aba_relatorio = aba
                                break

                    if aba_relatorio is None:
                        app.after(0, lambda: status_label.configure(text=f"Erro: Aba 'RELATORIO' não encontrada. Abas disponíveis: {', '.join(abas_disponiveis)}", text_color="red"))
                        wb.close()
                        return

                    sht = wb.sheets[aba_relatorio]
                    app.after(0, lambda: status_label.configure(text=f"Processando aba: {aba_relatorio}", text_color="blue"))

                    print("Removendo formas geradas anteriormente...")
                    cotas_anteriormente_geradas = 0
                    for shape in sht.api.Shapes:
                        if shape.Name.startswith("Cota_"):
                            shape.Delete()
                            cotas_anteriormente_geradas += 1

                    print(f"{cotas_anteriormente_geradas} formas removidas.")

                    ultima_linha = obter_ultima_linha_area_impressao(sht)
                    print(f"Processando até a linha: {ultima_linha}")

                    cotas_geradas = 0

                    cfg = CONFIGURACOES_ABAS.get(aba_relatorio, CONFIGURACAO_PADRAO)
                    modelo_RF = cfg["modelo_RF"]
                    coluna_codigo = cfg["coluna_codigo"]
                    coluna_dados = cfg["coluna_dados"]
                    offset_comp = cfg["offset_comp"]
                    offset_alt = cfg["offset_alt"]
                    offset_larg = cfg["offset_larg"]

                    if "offset_candidatos" in cfg:
                        resultado = detectar_offsets(
                            sht, coluna_codigo, coluna_dados,
                            cfg["codigos"], ultima_linha, cfg["offset_candidatos"]
                        )
                        if resultado is None:
                            app.after(0, lambda: status_label.configure(
                                text="Erro: não foi possível detectar o layout da planilha. "
                                     "Verifique o log para mais detalhes.",
                                text_color="red"
                            ))
                            return
                        offset_comp, offset_alt = resultado
                    codigos_procurados = cfg["codigos"]
                    print(f"Configuração carregada para aba '{aba_relatorio}': modelo={modelo_RF}, "
                          f"código={coluna_codigo}, dados={coluna_dados}")

                    for linha_atual in range(4, ultima_linha + 1):
                        if stop_event.is_set():
                            print("Processamento interrompido pelo usuário.")
                            app.after(0, lambda: status_label.configure(
                                text=f"Interrompido. {cotas_geradas} cota(s) gerada(s) até o momento.",
                                text_color="orange"
                            ))
                            break

                        if linha_atual % 100 == 0:
                            print(f"Processando linha {linha_atual} de {ultima_linha}")

                        try:
                            codigo = sht.range(f"{coluna_codigo}{linha_atual}").value

                            if codigo:
                                codigo_str = str(codigo).strip()

                                if codigo_str.upper() in codigos_procurados:
                                    print(f"Código {codigo_str} encontrado na linha {coluna_codigo}{linha_atual}")

                                    linha_comp = linha_atual + offset_comp
                                    linha_alt = linha_atual + offset_alt

                                    if offset_larg is not None:
                                        linha_larg = linha_atual + offset_larg
                                    else:
                                        linha_larg = None

                                    try:
                                        comprimento_raw = sht.range(f"{coluna_dados}{linha_comp}").value
                                        altura_raw = sht.range(f"{coluna_dados}{linha_alt}").value

                                        if linha_larg is not None:
                                            largura_raw = sht.range(f"{coluna_dados}{linha_larg}").value
                                        else:
                                            largura_raw = None

                                        comprimento = None if comprimento_raw is None or str(comprimento_raw).strip() == '' else comprimento_raw
                                        altura = None if altura_raw is None or str(altura_raw).strip() == '' else altura_raw
                                        largura = None if largura_raw is None or str(largura_raw).strip() == '' else largura_raw

                                        print(f"Comprimento da célula {coluna_dados}{linha_comp}: {comprimento}")
                                        print(f"Altura da célula {coluna_dados}{linha_alt}: {altura}")
                                        print(f"Largura da célula {coluna_dados}{linha_larg}: {largura}")

                                        if comprimento is not None or altura is not None or largura is not None:
                                            try:
                                                comprimento_float = None
                                                altura_float = None
                                                largura_float = None
                                                comprimento_formatado = None
                                                altura_formatada = None
                                                largura_formatada = None

                                                if comprimento is not None and str(comprimento).strip():
                                                    comprimento_float = float(comprimento)
                                                    comprimento_formatado = f"{comprimento_float:.2f}".replace('.', ',')

                                                if altura is not None and str(altura).strip():
                                                    altura_float = float(altura)
                                                    altura_formatada = f"{altura_float:.2f}".replace('.', ',')

                                                if largura is not None and str(largura).strip():
                                                    largura_float = float(largura)
                                                    largura_formatada = f"{largura_float:.2f}".replace('.', ',')

                                                if comprimento_formatado or altura_formatada or largura_formatada:
                                                    gerar_seta_e_texto(sht, linha_atual, comprimento_formatado, altura_formatada, largura_formatada, modelo_RF)
                                                    cotas_geradas += 1

                                                    valores_encontrados = []
                                                    if comprimento_formatado:
                                                        valores_encontrados.append(f"Comprimento: {comprimento_formatado}m")
                                                    if altura_formatada:
                                                        valores_encontrados.append(f"Altura: {altura_formatada}m")
                                                    if largura_formatada:
                                                        valores_encontrados.append(f"Largura: {largura_formatada}m")

                                                    print(f"Cota {cotas_geradas} gerada para código {codigo_str} - {', '.join(valores_encontrados)}")
                                                else:
                                                    print(f"Nenhum valor válido encontrado para conversão - Comp='{comprimento}', Alt='{altura}', Larg='{largura}'")

                                            except (ValueError, TypeError) as e:
                                                print(f"Erro ao converter valores para float: Comp='{comprimento}', Alt='{altura}', Larg='{largura}' - {e}")
                                        else:
                                            print(f"Todos os valores são None - Comprimento: {comprimento}, Altura: {altura}, Largura: {largura} - Pulando esta linha")

                                    except Exception as e:
                                        print(f"Erro ao acessar células {coluna_dados}{linha_comp} ou {coluna_dados}{linha_alt} ou {coluna_dados}{linha_larg} : {e}")

                        except Exception as e:
                            print(f"Erro ao acessar célula {coluna_codigo}{linha_atual}: {e}")
                            print(traceback.format_exc())

                    if cotas_geradas > 0:
                        app.after(0, lambda: status_label.configure(text=f"Sucesso! {cotas_geradas} cotas geradas.", text_color="green"))
                        try:
                            _fechar_e_reabrir(wb, app_excel, caminho, salvar=True)
                            wb = None
                        except Exception as e:
                            print(f"\nErro ao salvar, fechar ou reabrir o arquivo: {e}")
                            print(f"\n{traceback.format_exc()}")
                            app.after(0, lambda: status_label.configure(text=f"Cotas geradas, mas erro ao salvar/fechar: {str(e)}", text_color="orange"))
                    else:
                        app.after(0, lambda: status_label.configure(text="Nenhum código encontrado da lista de códigos procurados.", text_color="orange"))
                        try:
                            _fechar_e_reabrir(wb, app_excel, caminho, salvar=False)
                            wb = None
                        except Exception as e:
                            print(f"\nErro ao fechar ou reabrir workbook: {e}")
                            print(f"\n{traceback.format_exc()}")

                except Exception as e:
                    app.after(0, lambda: status_label.configure(text=f"Erro ao processar arquivo: {str(e)}", text_color="red"))
                    print(f"\nErro detalhado: {e}")
                    print(f"\n{traceback.format_exc()}")
                finally:
                    try:
                        if wb is not None:
                            _fechar_e_reabrir(wb, app_excel, caminho, salvar=False)
                            wb = None
                    except Exception as e:
                        print(f"\nErro ao fechar workbook no finally: {e}")
                        print(f"\n{traceback.format_exc()}")
            else:
                app.after(0, lambda: status_label.configure(text="Erro: Caminho do arquivo não encontrado.", text_color="red"))
        else:
            app.after(0, lambda: status_label.configure(text="Erro: Caminho do arquivo não fornecido.", text_color="red"))

    except Exception as e:
        print(f"\nErro na thread: {e}")
        print(f"\n{traceback.format_exc()}")
        app.after(0, lambda: status_label.configure(text=f"Erro interno: {str(e)}", text_color="red"))
    finally:
        sys.stdout = original_stdout
        app.after(0, lambda: btn_gerar_cotas.configure(state="normal", text="Gerar Cotas"))
        app.after(0, lambda: btn_parar_cotas.configure(state="disabled"))

def parar_cotas():
    """Sinaliza para a thread de processamento que deve parar"""
    stop_event.set()
    btn_parar_cotas.configure(state="disabled", text="Parando...")
    print("Sinal de parada enviado. Aguardando fim da iteração atual...")

def gerar_cotas():
    """Função chamada pelo botão - inicia o processamento em thread separada"""
    stop_event.clear()

    debug_text_widget.pack(pady=10, padx=20, fill="both", expand=True)
    debug_text_widget.delete("1.0", "end")

    btn_gerar_cotas.configure(state="disabled", text="Processando...")
    btn_parar_cotas.configure(state="normal", text="Parar")

    thread = threading.Thread(target=gerar_cotas_thread, daemon=True)
    thread.start()

def gerar_seta_e_texto(sht, linha_p, comprimento, altura, largura, modelo_RF):
    if modelo_RF == "RJ/NI":
        linha_destino = linha_p + 6
        celula_s = sht.range(f"S{linha_destino}")
        print(f"Posicionando setas na célula S{linha_destino} (linha do código P{linha_p} + 6)")
    else:
        linha_destino = linha_p + 2
        celula_s = sht.range(f"P{linha_destino}")
        print(f"Posicionando setas na célula P{linha_destino} (linha do código L{linha_p} + 2)")

    posicao_x = celula_s.left
    posicao_y = celula_s.top

    if posicao_x is None or posicao_y is None:
        print(f"Não foi possível obter as coordenadas da célula S{linha_destino}.")
        return

    max_linha = sht.api.UsedRange.Rows.Count
    if linha_destino > max_linha:
        print(f"linha_destino ({linha_destino}) está fora do limite da planilha ({max_linha}).")
        return

    setas_criadas = []

    if altura is not None:
        try:
            print("Criando seta vertical (altura)...")
            arrow_vertical = sht.api.Shapes.AddLine(
                posicao_x + 10,
                posicao_y,
                posicao_x + 10,
                posicao_y + ALTURA_SETA_PX
            )
            arrow_vertical.Name = "Cota_Arrow_Vertical"
            arrow_vertical.Line.EndArrowheadStyle = 2
            arrow_vertical.Line.BeginArrowheadStyle = 2
            arrow_vertical.Line.ForeColor.RGB = COR_AZUL
            arrow_vertical.Line.Weight = ESPESSURA_LINHA

            text_v = sht.api.Shapes.AddTextbox(
                1,
                posicao_x + 15,
                posicao_y + 30,
                60,
                30
            )
            text_v.Name = "Cota_Text_Vertical"
            text_v.TextFrame2.TextRange.Text = f"{altura}m"
            text_v.TextFrame2.TextRange.Font.Size = TAMANHO_FONTE
            text_v.TextFrame2.TextRange.ParagraphFormat.Alignment = 1
            text_v.TextFrame2.VerticalAnchor = 1
            text_v.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = COR_AZUL
            text_v.Line.Visible = False
            text_v.Fill.Visible = False

            setas_criadas.append(f"altura: {altura}m")
            print("Seta vertical criada com sucesso!")

        except Exception as e:
            print(f"\nErro ao criar seta vertical: {e}")
            print(f"\n{traceback.format_exc()}")

    if comprimento is not None:
        try:
            print("Criando seta horizontal (comprimento)...")
            offset_vertical = ALTURA_SETA_PX if altura is not None else 20

            print(f"Criando seta horizontal com coordenadas:")
            print(f"  Início: ({posicao_x}, {posicao_y + offset_vertical})")
            print(f"  Fim: ({posicao_x + 100}, {posicao_y + offset_vertical})")

            arrow_horizontal = sht.api.Shapes.AddLine(
                posicao_x,
                posicao_y + offset_vertical,
                posicao_x + LARGURA_SETA_PX,
                posicao_y + offset_vertical
            )
            arrow_horizontal.Name = "Cota_Arrow_Horizontal"
            arrow_horizontal.Line.EndArrowheadStyle = 2
            arrow_horizontal.Line.BeginArrowheadStyle = 2
            arrow_horizontal.Line.ForeColor.RGB = COR_AZUL
            arrow_horizontal.Line.Weight = ESPESSURA_LINHA

            text_h = sht.api.Shapes.AddTextbox(
                1,
                posicao_x + 35,
                posicao_y + offset_vertical + 5,
                60,
                30
            )
            text_h.Name = "Cota_Text_Horizontal"
            text_h.TextFrame2.TextRange.Text = f"{comprimento}m"
            text_h.TextFrame2.TextRange.Font.Size = TAMANHO_FONTE
            text_h.TextFrame2.TextRange.ParagraphFormat.Alignment = 1
            text_h.TextFrame2.VerticalAnchor = 1
            text_h.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = COR_AZUL
            text_h.Line.Visible = False
            text_h.Fill.Visible = False

            setas_criadas.append(f"comprimento: {comprimento}m")
            print("Seta horizontal criada com sucesso!")

        except Exception as e:
            print(f"\nErro ao criar seta horizontal: {e}")
            print(f"\n{traceback.format_exc()}")

    if largura is not None:
        try:
            print("Criando seta vertical (largura)...")
            offset_horizontal = 10 if altura is None else 90

            if altura is not None and comprimento is not None:
                offset_vertical = 0
            elif comprimento is not None:
                offset_vertical = 50
            else:
                offset_vertical = 50

            arrow_largura = sht.api.Shapes.AddLine(
                posicao_x + offset_horizontal,
                posicao_y + offset_vertical,
                posicao_x + offset_horizontal,
                posicao_y + offset_vertical + ALTURA_SETA_PX
            )
            arrow_largura.Name = "Cota_Arrow_Largura_Seta_Vertical"
            arrow_largura.Line.EndArrowheadStyle = 2
            arrow_largura.Line.BeginArrowheadStyle = 2
            arrow_largura.Line.ForeColor.RGB = COR_AZUL
            arrow_largura.Line.Weight = ESPESSURA_LINHA

            text_l = sht.api.Shapes.AddTextbox(
                1,
                posicao_x + offset_horizontal + 10,
                posicao_y + offset_vertical + 30,
                60,
                30
            )
            text_l.Name = "Cota_Text_Largura"
            text_l.TextFrame2.TextRange.Text = f"{largura}m"
            text_l.TextFrame2.TextRange.Font.Size = TAMANHO_FONTE
            text_l.TextFrame2.TextRange.ParagraphFormat.Alignment = 1
            text_l.TextFrame2.VerticalAnchor = 1
            text_l.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = COR_AZUL
            text_l.Line.Visible = False
            text_l.Fill.Visible = False

            setas_criadas.append(f"largura: {largura}m")
            print("Seta de largura criada com sucesso!")

        except Exception as e:
            print(f"\nErro ao criar seta de largura: {e}")
            print(f"\n{traceback.format_exc()}")

    if setas_criadas:
        print(f"Cota gerada com sucesso: {', '.join(setas_criadas)}")
    else:
        print("Nenhuma seta foi criada - valores inválidos")


# Criando a janela principal
app = ctk.CTk()
app.geometry("600x600")
app.title("Gerar Cotas para Excel")

titulo = ctk.CTkLabel(app, text="Gerar Cotas para Excel", font=("Arial", 20))
titulo.pack(pady=20)

label_caminho = ctk.CTkLabel(app, text="Caminho do arquivo Excel:")
label_caminho.pack(pady=(10, 5))

entrada_caminho = ctk.CTkEntry(app, placeholder_text="Selecione o arquivo Excel", width=350)
entrada_caminho.pack(pady=10)

btn_selecionar_arquivo = ctk.CTkButton(app, text="Aperte para encontrar o arquivo", command=selecionar_arquivo)
btn_selecionar_arquivo.pack(pady=10)

btn_gerar_cotas = ctk.CTkButton(app, text="Gerar Cotas", command=gerar_cotas)
btn_gerar_cotas.pack(pady=15)

btn_parar_cotas = ctk.CTkButton(app, text="Parar", command=parar_cotas, state="disabled", fg_color="red", hover_color="darkred")
btn_parar_cotas.pack(pady=(0, 10))

status_label = ctk.CTkLabel(app, text="", font=("Arial", 12))
status_label.pack(pady=5)

debug_text_widget = ctk.CTkTextbox(app, height=200, width=550, font=("Arial", 11), state="normal")
debug_text_widget.pack_forget()

rodape = ctk.CTkLabel(app, text="Dawhen © 2025 - Todos os direitos reservados", font=("Arial", 10))
rodape.pack(side="bottom", pady=15)

update_debug_display()

app.mainloop()
