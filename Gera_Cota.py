import customtkinter as ctk
from tkinter import filedialog
import xlwings as xw
import os
import threading
import queue
import sys
from io import StringIO
import traceback
import time

# Configuração inicial do tema do CustomTkinter
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# Variáveis globais para threading e logging
log_queue = queue.Queue()
debug_text_widget = None

class DebugCapture:
    """Classe para capturar prints e redirecionar para a interface"""
    def __init__(self, queue_obj):
        self.queue = queue_obj
        self.original_stdout = sys.__stdout__  # Usar sys.__stdout__ em vez de sys.stdout
        
    def write(self, text):
        if text.strip():  # Só adiciona se não for string vazia
            try:
                self.queue.put(text.strip())
            except Exception as e:
                print(f"\nErro ao adicionar à queue: {e}")
                print(f"\n{traceback.format_exc()}")  # Exibe o traceback completo
                pass  # Se não conseguir adicionar à queue, ignora
        
        # Garantir que original_stdout existe antes de usar
        if self.original_stdout:
            try:
                self.original_stdout.write(text)
            except Exception as e:
                print(f"\nErro ao escrever no console: {e}")
                print(f"\n{traceback.format_exc()}")  # Exibe o traceback completo
                pass  # Se não conseguir escrever no console, ignora
        
    def flush(self):
        if self.original_stdout and hasattr(self.original_stdout, 'flush'):
            try:
                self.original_stdout.flush()
            except:
                pass

def log_debug(message):
    """Função para adicionar mensagens de debug"""
    if debug_text_widget:
        try:
            log_queue.put(message)
        except:
            print(message)  # Fallback para console
    else:
        print(message)

def update_debug_display():
    """Atualiza a caixa de texto com as mensagens de debug"""
    try:
        while True:
            message = log_queue.get_nowait()
            if debug_text_widget:
                debug_text_widget.insert("end", message + "\n")
                debug_text_widget.see("end")  # Scroll para o final
    except queue.Empty:
        pass
    
    # Agendar próxima verificação
    if debug_text_widget:
        app.after(100, update_debug_display)

# Função para selecionar o arquivo Excel
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
        entrada_caminho.delete(0, 'end')  # Limpa o campo de entrada
        entrada_caminho.insert(0, caminho_arquivo)  # Insere o caminho selecionado
        status_label.configure(text="Arquivo selecionado com sucesso!", text_color="green")
    else:
        status_label.configure(text="Erro: Nenhum arquivo selecionado.", text_color="red")

def obter_ultima_linha_area_impressao(sht):
    """
    Identifica a última linha da área de impressão da planilha
    """
    try:
        # Método 1: Verificar se há área de impressão definida
        print_area = sht.api.PageSetup.PrintArea
        if print_area:
            print(f"Área de impressão definida: {print_area}")
            # Extrair a última linha da área de impressão
            # Exemplo: "$A$1:$Z$100" -> linha 100
            import re
            match = re.search(r'\$[A-Z]+\$(\d+)$', print_area)
            if match:
                ultima_linha_impressao = int(match.group(1))
                print(f"Última linha da área de impressão: {ultima_linha_impressao}")
                return ultima_linha_impressao
        
        # Método 2: Se não há área de impressão, usar a última linha com dados
        print("Área de impressão não definida, procurando última linha com dados...")
        
        # Encontrar a última linha com dados em qualquer coluna
        used_range = sht.used_range
        if used_range:
            ultima_linha_dados = used_range.last_cell.row
            print(f"Última linha com dados: {ultima_linha_dados}")
            return ultima_linha_dados
        
        # Método 3: Procurar especificamente nas colunas que interessam
        print("Procurando última linha nas colunas específicas...")
        
        # Para aba RELATORIO - verificar coluna P
        # Para aba RELATÓRIO GERAL - verificar coluna L
        colunas_verificar = ['P', 'L', 'O', 'K', 'S']
        ultima_linha_encontrada = 0
        
        for coluna in colunas_verificar:
            try:
                # Procurar a última célula não vazia na coluna
                for linha in range(2000, 0, -1):  # De 2000 para 1, decrementando
                    valor = sht.range(f"{coluna}{linha}").value
                    if valor is not None and str(valor).strip():
                        print(f"Última linha com dados na coluna {coluna}: {linha}")
                        ultima_linha_encontrada = max(ultima_linha_encontrada, linha)
                        break
            except Exception as e:
                print(f"\nErro ao acessar coluna {coluna}: {e}")
                print(f"\n{traceback.format_exc()}")  # Exibe o traceback completo
                continue
        
        if ultima_linha_encontrada > 0:
            return ultima_linha_encontrada
        
        # Método 4: Fallback - retornar 2000 como padrão
        print("Não foi possível determinar a última linha, usando 2000 como padrão")
        return 2000
        
    except Exception as e:
        print(f"\nErro ao obter última linha da área de impressão: {e}")
        print(f"\n{traceback.format_exc()}")  # Exibe o traceback completo
        return 2000

# Função para gerar as cotas (executada em thread separada)
def gerar_cotas_thread():
    """Função que executa o processamento em thread separada"""
    original_stdout = sys.stdout  # Salvar stdout original
    wb = None  # Inicializar variável wb
    app_excel = None  # Variável para armazenar a instância do Excel

    try:
        # Redirecionar prints para a caixa de debug
        sys.stdout = DebugCapture(log_queue)

        caminho = entrada_caminho.get().strip()  # Acessa o caminho da entrada de texto
        if caminho:
            if os.path.exists(caminho):
                try:
                    wb = xw.Book(caminho)
                    app_excel = wb.app  # Armazenar a instância do Excel
                    
                    # Listar todas as abas disponíveis para debug
                    abas_disponiveis = [sheet.name for sheet in wb.sheets]
                    print(f"Abas disponíveis: {abas_disponiveis}")
                    
                    # Tentar encontrar a aba correta
                    aba_relatorio = None
                    nomes_possiveis = ["RELATORIO", "RELATÓRIO", "Relatorio", "Relatório", "RELATÓRIO GERAL", "RELATORIO GERAL"]
                    
                    for nome in nomes_possiveis:
                        if nome in abas_disponiveis:
                            aba_relatorio = nome
                            break

                    # Se não encontrou pelos nomes exatos, procura por substring
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

                    # Não apagar as formas VBA (alterar a exclusão para as cotas geradas)
                    print("Removendo formas geradas anteriormente...")
                    cotas_anteriormente_geradas = 0
                    for shape in sht.api.Shapes:
                        if shape.Name.startswith("Cota_"):  # Só remove as formas geradas pela função
                            shape.Delete()
                            cotas_anteriormente_geradas += 1

                    print(f"{cotas_anteriormente_geradas} formas removidas.")

                    # Obter a última linha da área de impressão
                    ultima_linha = obter_ultima_linha_area_impressao(sht)
                    print(f"Processando até a linha: {ultima_linha}")

                    cotas_geradas = 0

                    # Definir configurações baseadas na aba
                    if aba_relatorio == "RELATORIO":
                        # Códigos que devemos procurar na coluna P
                        codigos_procurados = ["17.1", "17.3", "17.4", "17.6", "17.7", "17.8", "17.10", "17.11", "29.2", "29.7"]
                        modelo_RF = "RJ/NI"
                        coluna_codigo = "P"
                        coluna_dados = "O"
                        offset_comp = 2  # O(n+2)
                        offset_alt = 3   # O(n+3)
                        offset_larg = 5  # O(n+5)
                        print(f"Configuração: RELATORIO, modelo {modelo_RF} - Códigos na coluna P, dados na coluna S")
                    elif aba_relatorio == "RELATÓRIO GERAL":
                        # Códigos que devemos procurar na coluna L
                        codigos_procurados = [
                            "17.1 CORRIMÃO", "17.1 ESCUDO", "17.1 PIQUETE", "17.1 BICICLETÁRIO",
                            "17.3 PAREDE", "17.4 PAREDE", "17.4 MARQUISE", "17.4 FORRO", "17.6 PAREDE",
                            "17.6 RODAPÉ", "17.6 PILAR", "17.6 MURETA", "17.6 MURO", "17.6 MARQUISE",
                            "17.6 FORRO", "17.7 PORTA", "17.8 VAGAS", "17.8 TÁTIL", "17.9 LETREIRO",
                            "17.9 TOTEM", "17.9 LIXEIRA", "17.11 PAREDE", "17.11 RODAPÉ", "17.11 PILAR",
                            "17.11 MURETA", "17.11  MURO", "17.11 MARQUISE", "17.11  FORRO"
                        ]
                        modelo_RF = "Demais_RFs"
                        coluna_codigo = "L"
                        coluna_dados = "K"
                        offset_comp = 3  # K(n+3)
                        offset_alt = 4   # K(n+4)
                        offset_larg = None  # Não existe largura para esse modelo
                        print(f"Configuração: RELATÓRIO GERAL, modelo {modelo_RF} - Códigos na coluna L, dados na coluna K")
                    else:
                        # Para outras abas, usar configuração padrão (RELATORIO)
                        coluna_codigo = "P"
                        coluna_dados = "S"
                        offset_comp = 2
                        offset_alt = 3
                        print(f"Configuração padrão para aba '{aba_relatorio}' - Códigos na coluna P, dados na coluna S")

                    # Pesquisar linha por linha, da linha 4 até a última linha da área de impressão
                    for linha_atual in range(4, ultima_linha + 1):
                        if linha_atual % 100 == 0:  # Log a cada 100 linhas
                            print(f"Processando linha {linha_atual} de {ultima_linha}")
                            
                        try:
                            # Verificar se há código na coluna definida pela configuração
                            codigo = sht.range(f"{coluna_codigo}{linha_atual}").value
                            
                            if codigo:
                                codigo_str = str(codigo).strip()
                                
                                # Verifica se o código está na lista de códigos procurados
                                if codigo_str.upper() in codigos_procurados:
                                    print(f"Código {codigo_str} encontrado na linha {coluna_codigo}{linha_atual}")
                                    
                                    # Calcular as linhas correspondentes para comprimento e altura
                                    linha_comp = linha_atual + offset_comp
                                    linha_alt = linha_atual + offset_alt
                                    
                                    # Só calcular linha_larg se offset_larg não for None
                                    if offset_larg is not None:
                                        linha_larg = linha_atual + offset_larg
                                    else:
                                        linha_larg = None

                                    # Pegar os valores de comprimento e altura na coluna definida
                                    try:
                                        comprimento_raw = sht.range(f"{coluna_dados}{linha_comp}").value
                                        altura_raw = sht.range(f"{coluna_dados}{linha_alt}").value
                                        
                                        # Só acessar largura se linha_larg não for None
                                        if linha_larg is not None:
                                            largura_raw = sht.range(f"{coluna_dados}{linha_larg}").value
                                        else:
                                            largura_raw = None
                                        
                                        # Normalizar valores vazios para None
                                        comprimento = None if comprimento_raw is None or str(comprimento_raw).strip() == '' else comprimento_raw
                                        altura = None if altura_raw is None or str(altura_raw).strip() == '' else altura_raw
                                        largura = None if largura_raw is None or str(largura_raw).strip() == '' else largura_raw
                                        
                                        print(f"Comprimento da célula {coluna_dados}{linha_comp}: {comprimento}")
                                        print(f"Altura da célula {coluna_dados}{linha_alt}: {altura}")
                                        print(f"Largura da célula {coluna_dados}{linha_larg}: {largura}")

                                        # Verificar se os valores não são None antes de tentar converter
                                        if comprimento is not None or altura is not None or largura is not None:
                                            # Tentar converter para float
                                            try:
                                                comprimento_float = None
                                                altura_float = None
                                                largura_float = None
                                                comprimento_formatado = None
                                                altura_formatada = None
                                                largura_formatada = None
                                                
                                                # Converter comprimento se existir e não for string vazia
                                                if comprimento is not None and str(comprimento).strip():
                                                    comprimento_float = float(comprimento)
                                                    comprimento_formatado = f"{comprimento_float:.2f}".replace('.', ',')
                                                
                                                # Converter altura se existir e não for string vazia
                                                if altura is not None and str(altura).strip():
                                                    altura_float = float(altura)
                                                    altura_formatada = f"{altura_float:.2f}".replace('.', ',')

                                                # Converter largura se existir e não for string vazia
                                                if largura is not None and str(largura).strip():
                                                    largura_float = float(largura)
                                                    largura_formatada = f"{largura_float:.2f}".replace('.', ',')

                                                # Verificar se pelo menos um valor foi convertido com sucesso
                                                if comprimento_formatado or altura_formatada or largura_formatada:
                                                    # Gerar cota com valores formatados (pode ser apenas um dos valores)
                                                    gerar_seta_e_texto(sht, linha_atual, comprimento_formatado, altura_formatada, largura_formatada, modelo_RF)
                                                    cotas_geradas += 1
                                                    
                                                    # Log específico para cada caso
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
                            print(traceback.format_exc())  # Exibe o traceback completo
                    
                    if cotas_geradas > 0:
                        app.after(0, lambda: status_label.configure(text=f"Sucesso! {cotas_geradas} cotas geradas.", text_color="green"))
                        # Salvar o arquivo
                        try:
                            wb.save()
                            print("Arquivo salvo com sucesso!")
                            time.sleep(1)  # Aguarda 1 segundo para garantir que o salvamento seja concluído
                            
                            # Fechar apenas o workbook, não o Excel
                            wb.close()
                            wb = None
                            print("Workbook fechado com sucesso!")
                            
                            # Reabrir o workbook na mesma instância do Excel
                            print("Reabrindo arquivo na mesma instância do Excel...")
                            time.sleep(1)  # Pequeno atraso antes de reabrir
                            wb_novo = app_excel.books.open(caminho)
                            print(f"Arquivo {caminho} reaberto com sucesso!")
                            
                        except Exception as e:
                            print(f"\nErro ao salvar, fechar ou reabrir o arquivo: {e}")
                            print(f"\n{traceback.format_exc()}")
                            app.after(0, lambda: status_label.configure(text=f"Cotas geradas, mas erro ao salvar/fechar: {str(e)}", text_color="orange"))
                    else:
                        app.after(0, lambda: status_label.configure(text="Nenhum código encontrado da lista de códigos procurados.", text_color="orange"))
                        # Fechar workbook mesmo se não há cotas geradas
                        try:
                            wb.close()
                            wb = None
                            print("Workbook fechado com sucesso!")
                            
                            # Reabrir o workbook na mesma instância do Excel
                            print("Reabrindo arquivo na mesma instância do Excel...")
                            time.sleep(1)  # Pequeno atraso antes de reabrir
                            wb_novo = app_excel.books.open(caminho)
                            print(f"Arquivo {caminho} reaberto com sucesso!")
                            
                        except Exception as e:
                            print(f"\nErro ao fechar ou reabrir workbook: {e}")
                            print(f"\n{traceback.format_exc()}")
                except Exception as e:
                    app.after(0, lambda: status_label.configure(text=f"Erro ao processar arquivo: {str(e)}", text_color="red"))
                    print(f"\nErro detalhado: {e}")
                    print(f"\n{traceback.format_exc()}")  # Exibe o traceback completo
                finally:
                    # Sempre fechar o workbook se ele ainda estiver aberto
                    try:
                        if wb is not None:
                            wb.close()
                            wb = None
                            print("Workbook fechado no finally!")
                            
                            # Tentar reabrir mesmo no caso de erro
                            if app_excel is not None:
                                try:
                                    print("Tentando reabrir arquivo após erro na mesma instância...")
                                    time.sleep(1)
                                    wb_novo = app_excel.books.open(caminho)
                                    print(f"Arquivo {caminho} reaberto com sucesso após erro!")
                                except Exception as e2:
                                    print(f"\nErro ao reabrir arquivo após erro: {e2}")
                                    print(f"\n{traceback.format_exc()}")
                                
                    except Exception as e:
                        print(f"\nErro ao fechar workbook no finally: {e}")
                        print(f"\n{traceback.format_exc()}")  # Exibe o traceback completo
            else:
                app.after(0, lambda: status_label.configure(text="Erro: Caminho do arquivo não encontrado.", text_color="red"))
        else:
            app.after(0, lambda: status_label.configure(text="Erro: Caminho do arquivo não fornecido.", text_color="red"))

    except Exception as e:
        # Se houver erro na thread, restaurar stdout e mostrar erro
        print(f"\nErro na thread: {e}")
        print(f"\n{traceback.format_exc()}")  # Exibe o traceback completo
        app.after(0, lambda: status_label.configure(text=f"Erro interno: {str(e)}", text_color="red"))
    finally:
        # Sempre restaurar stdout original
        sys.stdout = original_stdout
        app.after(0, lambda: btn_gerar_cotas.configure(state="normal", text="Gerar Cotas"))

def fechar_e_abrir_arquivo(caminho):
    """Função para fechar e reabrir o arquivo Excel após o processamento"""
    try:
        print(f"Tentando reabrir arquivo: {caminho}")
        if os.path.exists(caminho):
            # Tentar reabrir o arquivo
            wb_novo = xw.Book(caminho)
            print(f"Arquivo {caminho} reaberto com sucesso.")
        else:
            print(f"Arquivo {caminho} não encontrado ao tentar reabrir.")
    except Exception as e:
        print(f"\nErro ao reabrir o arquivo: {e}")
        print(f"\n{traceback.format_exc()}")  # Exibe o traceback completo

def gerar_cotas():
    """Função chamada pelo botão - inicia o processamento em thread separada"""
    # Mostrar caixa de debug e limpar conteúdo anterior
    debug_text_widget.pack(pady=10, padx=20, fill="both", expand=True)
    debug_text_widget.delete("1.0", "end")
    
    # Desabilitar botão durante processamento
    btn_gerar_cotas.configure(state="disabled", text="Processando...")
    
    # Iniciar thread de processamento
    thread = threading.Thread(target=gerar_cotas_thread, daemon=True)
    thread.start()

# Função para gerar a seta e a caixa de texto com as medidas
def gerar_seta_e_texto(sht, linha_p, comprimento, altura, largura, modelo_RF):
    if modelo_RF == "RJ/NI":
        # Calcular a linha onde as setas serão posicionadas: linha do código + 6
        linha_destino = linha_p + 6
        celula_s = sht.range(f"S{linha_destino}")
        print(f"Posicionando setas na célula S{linha_destino} (linha do código P{linha_p} + 6)")
    else:
        # Calcular a linha onde as setas serão posicionadas: linha do código + 5
        linha_destino = linha_p + 5
        celula_s = sht.range(f"P{linha_destino}")
        print(f"Posicionando setas na célula P{linha_destino} (linha do código L{linha_p} + 5)")
    
    # Obter as coordenadas da célula em pixels
    posicao_x = celula_s.left
    posicao_y = celula_s.top

    if posicao_x is None or posicao_y is None:
        print(f"Não foi possível obter as coordenadas da célula S{linha_destino}.")
        return

    max_linha = sht.api.UsedRange.Rows.Count
    if linha_destino > max_linha:
        print(f"linha_destino ({linha_destino}) está fora do limite da planilha ({max_linha}).")
        return

    # Verificar quais setas criar baseado nos valores disponíveis
    setas_criadas = []

    # Criar seta vertical (para altura) apenas se altura estiver disponível
    if altura is not None:
        try:
            print("Criando seta vertical (altura)...")
            arrow_vertical = sht.api.Shapes.AddLine(
                posicao_x + 10,
                posicao_y,
                posicao_x + 10,
                posicao_y + 60
            )
            arrow_vertical.Name = "Cota_Arrow_Vertical"  # Nome para identificação
            arrow_vertical.Line.EndArrowheadStyle = 2
            arrow_vertical.Line.BeginArrowheadStyle = 2
            arrow_vertical.Line.ForeColor.RGB = 0x0000FF  # Cor Azul
            arrow_vertical.Line.Weight = 1.5

            # Adicionando o texto da altura ao lado da seta vertical
            text_v = sht.api.Shapes.AddTextbox(
                1,
                posicao_x + 15,
                posicao_y + 20,
                50,
                20
            )
            text_v.Name = "Cota_Text_Vertical"  # Nome para identificação
            text_v.TextFrame2.TextRange.Text = f"{altura}m"
            text_v.TextFrame2.TextRange.Font.Size = 10
            text_v.TextFrame2.TextRange.ParagraphFormat.Alignment = 1
            text_v.TextFrame2.VerticalAnchor = 1
            text_v.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0x0000FF  # Cor Azul
            text_v.Line.Visible = False
            text_v.Fill.Visible = False
            
            setas_criadas.append(f"altura: {altura}m")
            print("Seta vertical criada com sucesso!")
            
        except Exception as e:
            print(f"\nErro ao criar seta vertical: {e}")
            print(f"\n{traceback.format_exc()}")  # Exibe o traceback completo

    # Criar seta horizontal (para comprimento) apenas se comprimento estiver disponível
    if comprimento is not None:
        try:
            print("Criando seta horizontal (comprimento)...")
            # Se já existe seta vertical, posicionar a horizontal mais abaixo
            offset_vertical = 80 if altura is not None else 20

            print(f"Criando seta horizontal com coordenadas:")
            print(f"  Início: ({posicao_x}, {posicao_y + offset_vertical})")
            print(f"  Fim: ({posicao_x + 100}, {posicao_y + offset_vertical})")

            arrow_horizontal = sht.api.Shapes.AddLine(
                posicao_x,
                posicao_y + offset_vertical,
                posicao_x + 100,
                posicao_y + offset_vertical
            )
            arrow_horizontal.Name = "Cota_Arrow_Horizontal"  # Nome para identificação
            arrow_horizontal.Line.EndArrowheadStyle = 2
            arrow_horizontal.Line.BeginArrowheadStyle = 2
            arrow_horizontal.Line.ForeColor.RGB = 0x0000FF  # Cor Azul
            arrow_horizontal.Line.Weight = 1.5

            # Adicionando o texto do comprimento abaixo da seta horizontal
            text_h = sht.api.Shapes.AddTextbox(
                1,
                posicao_x + 25,
                posicao_y + offset_vertical + 5, # 5 pixels abaixo da seta
                60,
                20
            )
            text_h.Name = "Cota_Text_Horizontal"  # Nome para identificação
            text_h.TextFrame2.TextRange.Text = f"{comprimento}m"
            text_h.TextFrame2.TextRange.Font.Size = 10
            text_h.TextFrame2.TextRange.ParagraphFormat.Alignment = 1
            text_h.TextFrame2.VerticalAnchor = 1
            text_h.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0x0000FF  # Cor Azul
            text_h.Line.Visible = False
            text_h.Fill.Visible = False
            
            setas_criadas.append(f"comprimento: {comprimento}m")
            print("Seta horizontal criada com sucesso!")
            
        except Exception as e:
            print(f"\nErro ao criar seta horizontal: {e}")
            print(f"\n{traceback.format_exc()}")  # Exibe o traceback completo
    
    if largura is not None:
        try:
            print("Criando seta vertical (largura)...")
            # Se já existe seta vertical, posicionar a seta de largura mais à direita
            offset_horizontal = 10 if altura is None else 90
            
            # Se altura e comprimento existirem, ajustar a posição vertical
            if altura is not None and comprimento is not None:
                offset_vertical = 0
            # Se comprimento existir, ajustar a posição vertical
            elif comprimento is not None:
                offset_vertical = 50
            else:
                offset_vertical = 50

            arrow_largura = sht.api.Shapes.AddLine(
                posicao_x + offset_horizontal,
                posicao_y + offset_vertical,
                posicao_x + offset_horizontal,
                posicao_y + offset_vertical + 60
            )
            arrow_largura.Name = "Cota_Arrow_Largura_Seta_Vertical"  # Nome para identificação
            arrow_largura.Line.EndArrowheadStyle = 2
            arrow_largura.Line.BeginArrowheadStyle = 2
            arrow_largura.Line.ForeColor.RGB = 0x0000FF  # Cor Azul
            arrow_largura.Line.Weight = 1.5

            # Adicionando o texto da largura ao lado da seta vertical
            text_l = sht.api.Shapes.AddTextbox(
                1,
                posicao_x + offset_horizontal + 10,  # 10 pixels à direita da seta
                posicao_y + offset_vertical + 20,  # Centralizado verticalmente
                60,
                20
            )
            text_l.Name = "Cota_Text_Largura"  # Nome para identificação
            text_l.TextFrame2.TextRange.Text = f"{largura}m"
            text_l.TextFrame2.TextRange.Font.Size = 10
            text_l.TextFrame2.TextRange.ParagraphFormat.Alignment = 1
            text_l.TextFrame2.VerticalAnchor = 1
            text_l.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0x0000FF  # Cor Azul
            text_l.Line.Visible = False
            text_l.Fill.Visible = False
            
            setas_criadas.append(f"largura: {largura}m")
            print("Seta de largura criada com sucesso!")
            
        except Exception as e:
            print(f"\nErro ao criar seta de largura: {e}")
            print(f"\n{traceback.format_exc()}")  # Exibe o traceback completo

    # Log do resultado final
    if setas_criadas:
        print(f"Cota gerada com sucesso: {', '.join(setas_criadas)}")
    else:
        print("Nenhuma seta foi criada - valores inválidos")


# Criando a janela principal
app = ctk.CTk()
app.geometry("600x600")
app.title("Gerar Cotas para Excel")

# Título
titulo = ctk.CTkLabel(app, text="Gerar Cotas para Excel", font=("Arial", 20))
titulo.pack(pady=20)

# Campo para o caminho do arquivo Excel
label_caminho = ctk.CTkLabel(app, text="Caminho do arquivo Excel:")
label_caminho.pack(pady=(10, 5))

# Definir a entrada de texto (campo de entrada) para o caminho do arquivo
entrada_caminho = ctk.CTkEntry(app, placeholder_text="Selecione o arquivo Excel", width=350)
entrada_caminho.pack(pady=10)

# Botão para selecionar o arquivo
btn_selecionar_arquivo = ctk.CTkButton(app, text="Aperte para encontrar o arquivo", command=selecionar_arquivo)
btn_selecionar_arquivo.pack(pady=10)

# Botão para gerar as cotas
btn_gerar_cotas = ctk.CTkButton(app, text="Gerar Cotas", command=gerar_cotas)
btn_gerar_cotas.pack(pady=15)

# Status do processo
status_label = ctk.CTkLabel(app, text="", font=("Arial", 12))
status_label.pack(pady=5)

# Caixa de texto para debug (inicialmente oculta)
debug_text_widget = ctk.CTkTextbox(app, height=200, width=550, font=("Arial", 11), state="normal")
debug_text_widget.pack_forget()  # Inicialmente oculta

# Rodapé
rodape = ctk.CTkLabel(app, text="Dawhen © 2025 - Todos os direitos reservados", font=("Arial", 10))
rodape.pack(side="bottom", pady=15)

# Iniciar o loop de atualização de debug
update_debug_display()

# Rodar a aplicação
app.mainloop()
