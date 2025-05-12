import tkinter as tk
import threading
from tkinter import messagebox, filedialog, ttk
import os
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, numbers
from openpyxl.formula.translate import Translator

abas_do_excel = ["COMPRAS SEFAZ", "COMPRAS ALTERDATA", "COMPRAS PRODUTOS", "VENDAS SEFAZ", "VENDAS ALTERDATA", "VENDAS PRODUTOS", "CHECK", "PRODUTOS", "APURAÇÃO"]

arquivos_selecionados = {"COMPRAS SEFAZ": None, "COMPRAS ALTERDATA": None, "COMPRAS PRODUTOS": None, "VENDAS SEFAZ": None, "VENDAS ALTERDATA": None, "VENDAS PRODUTOS": None,
}

carregando_popup = None

def mostrar_carregando():
    global carregando_popup, barra_progresso

    carregando_popup = tk.Toplevel()
    carregando_popup.title("Carregando")
    carregando_popup.geometry("250x100")
    carregando_popup.resizable(False, False)

    tk.Label(carregando_popup, text="Gerando Excel, aguarde...").pack(pady=10)

    barra_progresso = ttk.Progressbar(carregando_popup, mode="indeterminate", length=200)
    barra_progresso.pack(pady=5)
    barra_progresso.start(10)

    carregando_popup.grab_set()

def criar_excel_thread():
    import threading
    thread = threading.Thread(target=criar_excel)
    thread.start()

def formatar_valores_contabil(aba, colunas, linhas):
    """
    Aplica a formatação contábil (moeda) às células especificadas.

    :param aba: aba do Excel onde será aplicada a formatação.
    :param colunas: lista de letras das colunas (ex: ['B', 'C', 'D']).
    :param linhas: lista de números das linhas (ex: [2, 3, 4]).
    """
    for col in colunas:
        for linha in linhas:
            celula = aba[f"{col}{linha}"]
            celula.number_format = '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-'

def botao_excel(_janela, _frame, _nome, _linha):
    caminho_var = tk.StringVar()
    caminho_var.set("Nada selecionado")

    def escolher_arquivo():
        caminho = filedialog.askopenfilename(title=f'Escolha o arquivo {_nome}.xlsx', 
        filetypes=[("Arquivos do Excel", "*.xlsx"), ("Todos os arquivos", "*.*")])

        if caminho:
            arquivos_selecionados[_nome] = caminho
            caminho_var.set(os.path.basename(caminho))  # Mostra só o nome do arquivo

    botao = tk.Button(_frame, text=_nome, width=20, command=escolher_arquivo)
    botao.grid(row=_linha, column=0, padx=5, pady=2)

    label = tk.Label(_frame, textvariable=caminho_var, anchor="w", width=40)
    label.grid(row=_linha, column=1, padx=5, pady=2)

def botao_para_agrupar_exceis(_janela, _x, _y):
    global caixa_nome_excel

    # Escolhendo o nome do excel
    label = tk.Label(_janela, text="Nome do novo Excel:")
    label.place(x=_x, y=_y)

    caixa_nome_excel = tk.Entry(_janela, width=30)
    caixa_nome_excel.place(x=_x + 130, y=_y)

    botao_gerar = tk.Button(_janela, text="Gerar Excel", command=criar_excel_thread)
    botao_gerar.place(x=_x + 130, y=_y + 30)

def criar_abas_excel(_wb):
    for aba in abas_do_excel:
        _wb.create_sheet(aba)   

def copiar_planilhas(_aba, _wb):
    arquivo = arquivos_selecionados.get(_aba)

    if arquivo:
        excel_origem = load_workbook(arquivo)
        aba_excel_origem = excel_origem.active
        aba_destino = _wb[_aba]

        #Copiando os dados
        for row in aba_excel_origem.iter_rows():
            for cell in row:
                if cell.value is not None:
                    aba_destino[cell.coordinate].value = cell.value

def rotulo_coluna(_wb, _aba_destino, _rotulo):
    aba = _wb[_aba_destino]
    
    ultima_coluna = aba.max_column + 1  # Nova coluna onde as fórmulas vão
    celula_rotulo = aba.cell(row=1, column=ultima_coluna)
    celula_rotulo.value = _rotulo

def adicionar_formula_procv(_wb, _aba_destino, _aba_matriz, _valor_procurado, _coluna_busca, _rotulo):
    aba = _wb[_aba_destino]

    # Determina a próxima coluna livre (sem alterar durante o loop)
    coluna_destino = aba.max_column + 1

    # Adiciona o rótulo na primeira linha
    aba.cell(row=1, column=coluna_destino).value = _rotulo

    # Escreve a fórmula nas linhas restantes
    for linha in range(2, aba.max_row + 1):
        celula_ref = f"{_valor_procurado}{linha}"  # Ex: C2, C3, etc.
        formula = (
            f'=IFERROR(VLOOKUP({celula_ref}, \'{_aba_matriz}\'!{_coluna_busca}:{_coluna_busca}, 1, FALSE), "ERRO")'
        )
        aba.cell(row=linha, column=coluna_destino).value = formula

def adicionar_formula_cancelada(_wb, _aba_destino, _coluna_ref, _rotulo):
    aba = _wb[_aba_destino]

    # Determina a próxima coluna livre
    coluna_destino = aba.max_column + 1

    # Adiciona o rótulo na primeira linha
    aba.cell(row=1, column=coluna_destino).value = _rotulo

    # Adiciona a fórmula SE nas linhas seguintes
    for linha in range(2, aba.max_row + 1):
        celula_ref = f"{_coluna_ref}{linha}"  # Ex: N2, N3, etc.
        formula = f'=IF({celula_ref}="AUTORIZADA", "NÃO", "SIM")'
        aba.cell(row=linha, column=coluna_destino).value = formula

def formula_check(_wb, _celula, _formula):
    aba = _wb["CHECK"]

    # Fórmula: subtrai o valor de _formula_2 de _formula_1
    aba[_celula].value = f"={_formula}"

def formula_generica(_wb, _aba, _celula, _formula):
    aba = _wb[_aba]

    # Fórmula: subtrai o valor de _formula_2 de _formula_1
    aba[_celula].value = f"={_formula}"

def formula_somases(_wb, _aba, _linha, _coluna, _aba_coluna, _aba_matriz, _aba_matriz_coluna, _aba_check):
    aba = _wb["CHECK"]
    formula = f'=SUMIFS(\'{_aba_matriz}\'!{_aba_coluna}:{_aba_coluna}, \'{_aba_matriz}\'!{_aba_matriz_coluna}:{_aba_matriz_coluna}, \'{_aba}\'!{_aba_check})'
    aba.cell(row=_linha, column=_coluna).value = formula

def tabela_check(_wb, _mesclar_inicio, _mesclar_final, _rotulo, _cabecalho_linha, _linha):
    aba = _wb["CHECK"]

    rotulos_linhas = ["NÃO", "SIM", "TOTAL"]
    rotulos_cabecalho = ["CANCELADAS", "SEFAZ", "ALTERDATA", "PRODUTO", "DESCONTO", "ST Ñ APROVEITADO", "ST APROVEITADO"]

    aba.merge_cells(f"{_mesclar_inicio}:{_mesclar_final}")
    aba[_mesclar_inicio].value = _rotulo
    aba[_mesclar_inicio].alignment = Alignment(horizontal="center", vertical="center")

    #Cabeçalho
    for col, texto in enumerate(rotulos_cabecalho, start=1):
        aba.cell(row=_cabecalho_linha, column=col).value = texto

    #LINHAS
    for i, texto in enumerate(rotulos_linhas):
        aba.cell(row=_linha + i, column=1).value = texto

    #FORMULAS COMPRAS
    formula_check(_wb, "B6", "=B5-C5")
    formula_check(_wb, "D6", "=C5-D5")
    formula_check(_wb, "F6", "=D6+E5-F5-G5")

    #FORMULAS VENDAS
    formula_check(_wb, "B13", "=B12-C12")
    formula_check(_wb, "D13", "=C12-D12")
    formula_check(_wb, "F13", "=D13+E12-F12-G12")

    # Fórmulas para somar os valores
    abas = ["COMPRAS", "VENDAS"]
    linhas_iniciais = [3, 10]
    abas_sefaz = ["COMPRAS SEFAZ", "VENDAS SEFAZ"]
    abas_alterdata = ["COMPRAS ALTERDATA", "VENDAS ALTERDATA"]
    abas_produtos = ["COMPRAS PRODUTOS", "VENDAS PRODUTOS"]
    colunas_iniciais_check = [2, 2]

    for i in range(len(abas)):
        linha_check = linhas_iniciais[i]
        coluna_check = colunas_iniciais_check[i]

        # Definir a célula de referência para cada linha
        if abas[i] == "COMPRAS":
            celula_referencia_1 = f"A{linha_check}"  # A3 para COMPRAS
            celula_referencia_2 = f"A{linha_check + 1}"  # A4 para COMPRAS
        else:
            celula_referencia_1 = f"A{linha_check}"  # A10 para VENDAS
            celula_referencia_2 = f"A{linha_check + 1}"  # A11 para VENDAS

        # SEFAZ
        formula_somases(_wb, "CHECK", linha_check, coluna_check,'P', abas_sefaz[i], 'Y', celula_referencia_1)

        # ALTERDATA
        formula_somases(_wb, "CHECK", linha_check, coluna_check + 1,'J', abas_alterdata[i], 'I', celula_referencia_1)

        # PRODUTOS
        formula_somases(_wb, "CHECK", linha_check, coluna_check + 2,'M', abas_produtos[i], 'I', celula_referencia_1)
        formula_somases(_wb, "CHECK", linha_check, coluna_check + 3,'N', abas_produtos[i], 'I', celula_referencia_1)
        formula_somases(_wb, "CHECK", linha_check, coluna_check + 4,'O', abas_produtos[i], 'I', celula_referencia_1)
        formula_somases(_wb, "CHECK", linha_check, coluna_check + 5,'P', abas_produtos[i], 'I', celula_referencia_1)

        # SEFAZ para a segunda célula de referência
        formula_somases(_wb, "CHECK", linha_check + 1, coluna_check,'P', abas_sefaz[i], 'Y', celula_referencia_2)

        # ALTERDATA para a segunda célula de referência
        formula_somases(_wb, "CHECK", linha_check + 1, coluna_check + 1,'J', abas_alterdata[i], 'I', celula_referencia_2)

        # PRODUTOS para a segunda célula de referência
        formula_somases(_wb, "CHECK", linha_check + 1, coluna_check + 2,'M', abas_produtos[i], 'I', celula_referencia_2)
        formula_somases(_wb, "CHECK", linha_check + 1, coluna_check + 3,'N', abas_produtos[i], 'I', celula_referencia_2)
        formula_somases(_wb, "CHECK", linha_check + 1, coluna_check + 4,'O', abas_produtos[i], 'I', celula_referencia_2)
        formula_somases(_wb, "CHECK", linha_check + 1, coluna_check + 5,'P', abas_produtos[i], 'I', celula_referencia_2)

    #SOMANDOS AS CANCELADAS SIM OU NÃO
    colunas_soma = ['B', 'C', 'D', 'E', 'F', 'G']

    for idx, coluna in enumerate(colunas_soma, 1):  # '1' é a posição inicial para o índice
        formula_check(_wb, f"{coluna}5", f"=SUM({coluna}3:{coluna}4)")
        formula_check(_wb, f"{coluna}12", f"=SUM({coluna}10:{coluna}11)")

    #ALUGUEL
    aba.merge_cells("A15:A16")
    aba.cell(row=15, column=1).value = "ALUGUEL"
    aba.cell(row=16, column=2).value = 0
    aba.cell(row=15, column=2).value = "VALOR"
    aba.cell(row=15, column=3).value = "PIS"
    aba.cell(row=15, column=4).value = "COFINS"
    aba.cell(row=16, column=3).value = "=b16 * 1.65%"
    aba.cell(row=16, column=4).value = "=b16 * 7.6%"

    #ENERGIA
    aba.cell(row=17, column=1).value = "ENERGIA"
    aba.cell(row=17, column=2).value = "=SUMIFS('COMPRAS ALTERDATA'!J:J,'COMPRAS ALTERDATA'!H:H,1253)"
    aba.cell(row=17, column=3).value = "=b17 * 1.65%"
    aba.cell(row=17, column=4).value = "=b17 * 7.6%"

    #ERROS
    rotulos_erros = ["ERRO", "SEFAZ", "ALTERDATA", "PRODUTOS"]
    rotulos_erros_linhas = ["COMPRAS", "VENDAS"]

    for col, texto in enumerate(rotulos_erros, start=1):
        aba.cell(row=19, column=col).value = texto

    for i, texto in enumerate(rotulos_erros_linhas):
        aba.cell(row=20 + i, column=1).value = texto

    coluna = ['W', 'Q', 'AA']
    coluna_celula = ['B', 'C', 'D']
    abas_compras = ["COMPRAS SEFAZ", "COMPRAS ALTERDATA", "COMPRAS PRODUTOS"]
    abas_vendas = ["VENDAS SEFAZ", "VENDAS ALTERDATA", "VENDAS PRODUTOS", "CHECK", "PRODUTOS", "APURAÇÃO"]

    for aba_compra, aba_venda, col, celula in zip(abas_compras, abas_vendas, coluna, coluna_celula):
        formula_check(_wb, f"{celula}20", f"=COUNTIF('{aba_compra}'!{col}:{col}, \"ERRO\")")
        formula_check(_wb, f"{celula}21", f"=COUNTIF('{aba_venda}'!{col}:{col}, \"ERRO\")")
    
    # Aplicar formatação contábil para todas as linhas com produtos
    aba = _wb["CHECK"]
    linhas = [3, 4, 5, 6, 10, 11, 12, 13, 16, 17]
    formatar_valores_contabil(aba, colunas_soma, linhas)

def tabela_produtos(_wb):
    rotulos = ["NOME", "ORIGEM", "COMPRAS", "PIS", "COFINS", "ICMS", "", "VENDAS BRUTAS", "VENDAS 6929", "VENDAS 5929", "VENDAS LIQUIDAS", "PIS", "COFINS", "ICMS"]
    aba_origem_compras = _wb["COMPRAS PRODUTOS"]
    aba_origem_vendas = _wb["VENDAS PRODUTOS"]
    aba_destino = _wb["PRODUTOS"]

    produtos_compras = {cell.value for cell in aba_origem_compras['F'][1:] if cell.value}
    produtos_vendas = {cell.value for cell in aba_origem_vendas['F'][1:] if cell.value}
    todos_produtos = produtos_compras.union(produtos_vendas)

    for i, produto in enumerate(sorted(todos_produtos), start=2):
        aba_destino.cell(row=i, column=1).value = produto

        origem = "AMBOS" if produto in produtos_compras and produto in produtos_vendas else \
                 "COMPRAS" if produto in produtos_compras else "VENDAS"
        aba_destino.cell(row=i, column=2).value = origem

        # COMPRAS
        formula_generica(_wb, "PRODUTOS", f"C{i}", f"=SUMIFS('COMPRAS PRODUTOS'!M:M,'COMPRAS PRODUTOS'!F:F,A{i})")
        formula_generica(_wb, "PRODUTOS", f"D{i}", f"=SUMIFS('COMPRAS PRODUTOS'!T:T,'COMPRAS PRODUTOS'!F:F,A{i})")
        formula_generica(_wb, "PRODUTOS", f"E{i}", f"=SUMIFS('COMPRAS PRODUTOS'!U:U,'COMPRAS PRODUTOS'!F:F,A{i})")
        formula_generica(_wb, "PRODUTOS", f"F{i}", f"=SUMIFS('COMPRAS PRODUTOS'!V:V,'COMPRAS PRODUTOS'!F:F,A{i})")

        # VENDAS
        formula_generica(_wb, "PRODUTOS", f"H{i}", f"=SUMIFS('VENDAS PRODUTOS'!M:M,'VENDAS PRODUTOS'!F:F,A{i})")
        formula_generica(_wb, "PRODUTOS", f"I{i}", f"=SUMIFS('VENDAS PRODUTOS'!M:M,'VENDAS PRODUTOS'!F:F,A{i}, 'VENDAS PRODUTOS'!G:G, 5929)")
        formula_generica(_wb, "PRODUTOS", f"J{i}", f"=SUMIFS('VENDAS PRODUTOS'!M:M,'VENDAS PRODUTOS'!F:F,A{i}, 'VENDAS PRODUTOS'!G:G, 6929)")
        formula_generica(_wb, "PRODUTOS", f"K{i}", f"=H{i}-I{i}-J{i}")
        formula_generica(_wb, "PRODUTOS", f"L{i}", f"=SUMIFS('VENDAS PRODUTOS'!T:T,'VENDAS PRODUTOS'!F:F,A{i})")
        formula_generica(_wb, "PRODUTOS", f"M{i}", f"=SUMIFS('VENDAS PRODUTOS'!U:U,'VENDAS PRODUTOS'!F:F,A{i})")
        formula_generica(_wb, "PRODUTOS", f"N{i}", f"=SUMIFS('VENDAS PRODUTOS'!V:V,'VENDAS PRODUTOS'!F:F,A{i})")

    for j, rotulo in enumerate(rotulos, start=1):
        aba_destino.cell(row=1, column=j).value = rotulo

    # Aplicar formatação contábil para todas as linhas com produtos
    aba = _wb["PRODUTOS"]
    colunas = ['C', 'D', 'E', 'F', 'H', 'I', 'J', 'K', 'L', 'M', 'N']
    linhas = range(2, aba.max_row + 1)
    formatar_valores_contabil(aba, colunas, linhas)

def tabela_apuracao(_wb):
    aba = _wb["APURAÇÃO"]
    rotulos_horizontais = ["TIPO", "VALOR", "PIS", "COFINS", "ICMS"]
    rotulos_verticais = ["COMPRAS", "VENDAS", "TOTAL", "TIPO"]

    # Cabeçalho horizontal (linha 1)
    for col, titulo in enumerate(rotulos_horizontais, start=1):
        aba.cell(row=1, column=col).value = titulo

    # Cabeçalho vertical (coluna A)
    for lin, tipo in enumerate(rotulos_verticais, start=2):
        aba.cell(row=lin, column=1).value = tipo

    linha_compras_apuracao = ['C', 'D', 'E', 'F']  # Colunas da aba PRODUTOS para COMPRAS
    linha_vendas_apuracao = ['K', 'L', 'M', 'N']   # Colunas da aba PRODUTOS para VENDAS
    linha_apuracao = ['B', 'C', 'D', 'E']          # Colunas da aba APURAÇÃO

    # Formatando os valores contábeis
    formatar_valores_contabil(_wb["APURAÇÃO"], linha_apuracao, [2, 3, 4])

    for col_ap, col_compra, col_venda in zip(linha_apuracao, linha_compras_apuracao, linha_vendas_apuracao):
        if col_ap in ['C', 'D']:
            formula = (f'=SUM(PRODUTOS!{col_compra}:{col_compra}) + SUM(CHECK!{col_ap}16, CHECK!{col_ap}17)')
        else:
            formula = f'=SUM(PRODUTOS!{col_compra}:{col_compra})'

        formula_generica(_wb, "APURAÇÃO", f'{col_ap}2', formula)

        # VENDAS (linha 3)
        formula_generica(_wb, "APURAÇÃO", f'{col_ap}3', f'=-SUM(PRODUTOS!{col_venda}:{col_venda})')

        # TOTAL (linha 4 = COMPRAS + VENDAS)
        formula_generica(_wb, "APURAÇÃO", f'{col_ap}4', f'={col_ap}2+{col_ap}3')

        # Tipo (linha 5 = CRÉDITO ou DÉBITO)
        formula_generica(_wb, "APURAÇÃO", f'{col_ap}5', f'=IF({col_ap}4>=0, "CRÉDITO", "DÉBITO")')

    # Valor (linha 5) com texto alternativo
    formula_generica(_wb, "APURAÇÃO", 'B5', f'=IF(B4>=0, "FATUROU MENOS", "FATUROU MAIS")')

def criar_excel():
    nome_arquivo = caixa_nome_excel.get().strip()
    diretorio = filedialog.askdirectory()

    if not (diretorio and nome_arquivo):
        messagebox.showwarning("AVISO", "Por favor, escolha um nome e um diretório válidos para o arquivo.")
        return

    caminho_arquivo = os.path.join(diretorio, f'{nome_arquivo}.xlsx')

    if os.path.exists(caminho_arquivo):
        resposta = messagebox.askyesno("Arquivo Existente", f"O arquivo '{nome_arquivo}.xlsx' já existe. Deseja substituir?")
        if not resposta:
            return

    # Agora o try engloba tudo que pode lançar erro
    try:
        global carregando_popup
        mostrar_carregando()  # <- Chamada da função para exibir a tela de carregamento

        wb = Workbook()
        wb.remove(wb.active)

        criar_abas_excel(wb)

        for aba in arquivos_selecionados:
            copiar_planilhas(aba, wb)

        abas_formulas = ["COMPRAS", "VENDAS"]

        for aba in abas_formulas:
            adicionar_formula_procv(wb, f'{aba} SEFAZ', f"{aba} ALTERDATA", 'C', 'C', "ALTERDATA")
            adicionar_formula_procv(wb, f'{aba} SEFAZ', f"{aba} PRODUTOS", 'C', 'B', "PRODUTOS")
            adicionar_formula_cancelada(wb, f'{aba} SEFAZ', "N", "CANCELADAS")

            adicionar_formula_procv(wb, f'{aba} ALTERDATA', f"{aba} SEFAZ", 'C', 'C', "SEFAZ")
            adicionar_formula_procv(wb, f'{aba} ALTERDATA', f"{aba} PRODUTOS", 'C', 'B', "PRODUTOS")

            adicionar_formula_procv(wb, f'{aba} PRODUTOS', f"{aba} SEFAZ", 'B', 'C', "SEFAZ")
            adicionar_formula_procv(wb, f'{aba} PRODUTOS', f"{aba} ALTERDATA", 'B', 'C', "ALTERDATA")

        tabela_check(wb, "A1", "G1", "COMPRAS", 2, 3)
        tabela_check(wb, "A8", "G8", "VENDAS", 9, 10)

        tabela_produtos(wb)
        tabela_apuracao(wb)

        wb.save(caminho_arquivo)
        messagebox.showinfo("Sucesso", f"Arquivo '{nome_arquivo}.xlsx' criado com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao criar o arquivo: {e}")

    finally:
        if carregando_popup:
            barra_progresso.stop()
            carregando_popup.destroy()

def interface():
    global aluguel_valor
    # Criando a janela principal
    janela = tk.Tk()
    janela.title("Interface de Agrupamento de Excel")
    janela.geometry("700x450")

    # Frame para COMPRAS
    frame_compras = tk.LabelFrame(janela, text="COMPRAS", padx=10, pady=10)
    frame_compras.place(x=20, y=20)

    # Frame para VENDAS
    frame_vendas = tk.LabelFrame(janela, text="VENDAS", padx=10, pady=10)
    frame_vendas.place(x=20, y=160)

    # Frame para ALUGUEL
    frame_aluguel = tk.LabelFrame(janela, text="ALUGUEL", padx=10, pady=10)
    frame_aluguel.place(x=20, y=300)

    # Adicionando campos de entrada para valor de aluguel
    tk.Label(frame_aluguel, text="Valor do Aluguel:").grid(row=0, column=0)
    valor_aluguel = tk.Entry(frame_aluguel)
    valor_aluguel.grid(row=0, column=1)
    
    # Ajustando o espaçamento para as abas do Excel
    botao_excel(janela, frame_compras, "COMPRAS SEFAZ", 0)
    botao_excel(janela, frame_compras, "COMPRAS ALTERDATA", 1)
    botao_excel(janela, frame_compras, "COMPRAS PRODUTOS", 2)

    # Abas do Excel para VENDAS
    botao_excel(janela, frame_vendas, "VENDAS SEFAZ", 0)
    botao_excel(janela, frame_vendas, "VENDAS ALTERDATA", 1)
    botao_excel(janela, frame_vendas, "VENDAS PRODUTOS", 2)

    # Botão para agrupar as planilhas
    botao_para_agrupar_exceis(janela, 150, 380)  

    # Iniciar o loop da interface
    janela.mainloop()

# Iniciando o MAIN
if __name__ == "__main__":
    interface()