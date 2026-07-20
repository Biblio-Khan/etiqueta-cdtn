import os
import json
from tkinter import dialog
import pandas as pd
import io  # <--- Importante
from nicegui import ui
from reportlab.lib.pagesizes import A4, mm
from reportlab.pdfgen import canvas
import barcode
from barcode.writer import ImageWriter
from reportlab.lib.utils import ImageReader

ARQUIVO_DADOS = "biblioteca.json"

# --- FUNÇÕES DE DADOS ---
def carregar_dados():
    if os.path.exists(ARQUIVO_DADOS):
        with open(ARQUIVO_DADOS, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []

if os.path.exists('biblioteca.json'):
    with open('biblioteca.json', 'r', encoding='utf-8') as f:
        biblioteca = json.load(f)
else:
    biblioteca = []

def salvar_dados(dados):
    with open('biblioteca.json', 'w', encoding='utf-8') as f:
        json.dump(dados, f, ensure_ascii=False, indent=4)

biblioteca = carregar_dados()

# --- FUNÇÕES DE EXPORTAÇÃO E IMPORTAÇÃO ---
def exportar_para_excel():
    if not biblioteca:
        ui.notify("Estante vazia!", type='warning')
        return
    
    # Gera o arquivo
    df = pd.DataFrame(biblioteca)
    nome_arquivo = "inventario_biblioteca.xlsx"
    df.to_excel(nome_arquivo, index=False)
    
    # FORÇA O DOWNLOAD NO NAVEGADOR
    ui.download(nome_arquivo) 
    ui.notify(f"Download iniciado: {nome_arquivo}")

def baixar_etiquetas():
    caminho_pdf = gerar_pdf_etiquetas(preparar_dados_para_pdf(biblioteca))
    ui.download(caminho_pdf)

# --- 1. NOVA FUNÇÃO: Atualiza SÓ a lista de livros ---
@ui.refreshable
def renderizar_grid(termo_pesquisa=""):
    # Lógica de filtragem
    if termo_pesquisa == "":
        livros_filtrados = biblioteca
    else:
        termo = termo_pesquisa.lower()
        livros_filtrados = []
        for livro in biblioteca:
            titulo = str(livro.get('titulo', '')).lower()
            autor = str(livro.get('autor', '')).lower()
            tombo = str(livro.get('tombo', '')).lower()
            
            if termo in titulo or termo in autor or termo in tombo:
                livros_filtrados.append(livro)

    # Grid com os livros
    with ui.grid(columns=2).classes('w-full gap-4'):
        for livro in livros_filtrados:
            with ui.row().classes('p-3 border rounded items-center justify-between bg-white shadow-sm'):
                with ui.column():
                    ui.label(f"{livro.get('titulo', 'Sem Título')}").classes('font-bold text-sm')
                    ui.label(f"Autor: {livro.get('autor', 'Desconhecido')}").classes('text-xs text-gray-600')
                
                # ATENÇÃO AQUI: O botão de deletar agora chama renderizar_grid.refresh()
                ui.button(icon='delete', on_click=lambda l=livro: (biblioteca.remove(l), salvar_dados(biblioteca), renderizar_grid.refresh()), color='red').props('flat dense')


# --- 2. SUA FUNÇÃO ORIGINAL AJUSTADA: Desenha o cabeçalho e chama o grid ---
def renderizar_lista(): 
    if lista_estante is None:
        return
        
    lista_estante.clear()
    with lista_estante:
        with ui.row().classes('w-full justify-between items-center mb-4'):
            ui.label("Livros na Estante").classes('text-h5 font-bold')
            
            # O input agora só manda o @ui.refreshable recarregar o grid!
            ui.input(
                'Pesquisar...', 
                on_change=lambda e: renderizar_grid.refresh(e.value)
            ).props('clearable outlined dense').classes('w-64')
            
            with ui.row():
                ui.button("Baixar Estante", on_click=exportar_para_excel, icon='download').classes('bg-blue-600')
                ui.button("Imprimir Etiquetas", on_click=baixar_etiquetas, icon='print').classes('bg-orange-500')
            
                with ui.dialog() as dialog, ui.card():
                    ui.label("Deseja apagar toda a estante? Esta ação não pode ser desfeita.")
                    with ui.row():
                        ui.button("SIM, APAGAR", on_click=lambda: (biblioteca.clear(), salvar_dados(biblioteca), renderizar_lista(), dialog.close()), color='red')
                        ui.button("NÃO", on_click=dialog.close, color='grey')

                ui.button("Apagar Tudo", icon='delete_sweep', color='red', on_click=dialog.open)
        
        # Desenha os livros logo abaixo do cabeçalho
        renderizar_grid()

import io # Certifique-se de ter este import no topo do arquivo

def limpar_estante():
    global biblioteca
    biblioteca = [] # Limpa a lista na memória
    salvar_dados(biblioteca) # Atualiza o arquivo JSON para ficar vazio
    renderizar_lista() # Atualiza a tela
    ui.notify("Estante limpa com sucesso!", type='info')

# Adicione a palavra 'async' antes de 'def'
async def processar_upload(event):
    try:
        # 1. Lê o conteúdo
        conteudo = await event.file.read()
        
        # 2. Lê a planilha. O 'header=0' garante que a 1ª linha seja o nome das colunas
        df = pd.read_excel(io.BytesIO(conteudo), dtype={'cdu': str, 'tombo': str, 'exemplar': str})

        # 3. Limpa espaços em branco nos nomes das colunas e deixa tudo minúsculo
        # Usando .strip() para remover espaços acidentais como " titulo "
        df.columns = [str(c).strip().lower() for c in df.columns]
        
        # 4. Remove linhas completamente vazias
        df = df.dropna(how='all')
        
        # --- ORDENAÇÃO ALFABÉTICA ---
        # Verifica se a coluna 'titulo' existe antes de ordenar
        if 'titulo' in df.columns:
            # Organiza pelo título. 
            # astype(str).str.lower() garante que "A" e "a" sejam tratados iguais.
            # na_position='last' joga livros sem título para o final da lista.
            df = df.sort_values(
                by='titulo', 
                ascending=True, 
                key=lambda col: col.astype(str).str.lower(),
                na_position='last'
            )
        
        # 5. Converte para dicionários
        novos_dados = df.to_dict('records')
        
        # 6. Adiciona e atualiza
        global biblioteca
        biblioteca.extend(novos_dados)
        salvar_dados(biblioteca)
        renderizar_lista()
        
        ui.notify(f"Sucesso! {len(novos_dados)} livros importados e organizados.", type='positive')
        
    except Exception as e:
        ui.notify(f"Erro ao processar o arquivo: {str(e)}", type='negative')
        
        
        
    except Exception as e:
        ui.notify(f"Erro: {e}", type='negative')
        print(f"Erro detalhado: {e}")

# 1. Nova função de limpeza (fica antes)
def preparar_dados_para_pdf(biblioteca):
    dados_formatados = []
    for livro in biblioteca:
        # Garante que, ao converter para dicionário, o CDU já seja tratado como string
        item = {k: (str(v) if k == 'cdu' else v) for k, v in livro.items()}
        dados_formatados.append(item)
    return dados_formatados

def gerar_barcode_memoria(numero_tombo):
    buffer = io.BytesIO()
    CODE128 = barcode.get_barcode_class('code128')
    codigo = CODE128(str(numero_tombo), writer=ImageWriter())
    
    # write_text=False remove o número que está saindo "atropelado"
    codigo.write(buffer, options={
        "write_text": False, 
        "module_width": 0.25, 
        "module_height": 10.0
    })
    buffer.seek(0)
    return buffer

# --- FUNÇÃO DE ETIQUETA ---
def gerar_pdf_etiquetas(books):
    arquivo_pdf = "etiquetas.pdf"
    if os.path.exists(arquivo_pdf):
        try: os.remove(arquivo_pdf)
        except PermissionError: pass
    
    c = canvas.Canvas(arquivo_pdf, pagesize=A4)
    LARGURA_ETQ, ALTURA_ETQ = 80 * mm, 25 * mm
    GAP_X, GAP_Y = 20 * mm, 5 * mm 
    MARGEM_X, MARGEM_Y = 15 * mm, 270 * mm
    coluna, linha = 0, 0
    
    def imprimir_campo(canvas, x, y, prefixo, valor, fonte, tamanho):
        valor_limpo = str(valor or '').replace('.ed', '').replace('.ex', '').strip()
        if valor_limpo and valor_limpo.lower() not in ['none', 'nan', '']:
            canvas.setFont(fonte, tamanho)
            canvas.drawString(x, y, f"{prefixo}{valor_limpo}")
            return True
        return False

    for book in books:
        x = MARGEM_X + (coluna * (LARGURA_ETQ + GAP_X))
        y = MARGEM_Y - (linha * (ALTURA_ETQ + GAP_Y))
        c.setLineWidth(0.2)
        c.rect(x, y, LARGURA_ETQ, ALTURA_ETQ)
        
        # --- LADO ESQUERDO ---
        y_lombada = y + 20*mm
        cdu_valor = str(book.get('cdu', ''))
        if len(cdu_valor) < 5 and '.' in cdu_valor: cdu_valor = cdu_valor.zfill(5)
        c.setFont("Helvetica-Bold", 7)
        c.drawString(x + 2*mm, y_lombada, cdu_valor); y_lombada -= 3*mm
        c.setFont("Helvetica", 6)
        c.drawString(x + 2*mm, y_lombada, book.get('cutter', '')); y_lombada -= 3*mm
        if imprimir_campo(c, x + 2*mm, y_lombada, "ed.", book.get('edicao'), "Helvetica", 6): y_lombada -= 3*mm
        if imprimir_campo(c, x + 2*mm, y_lombada, "ex.", book.get('exemplar'), "Helvetica", 6): y_lombada -= 3*mm
        c.setFont("Helvetica-Bold", 6)
        c.drawString(x + 2*mm, y + 4*mm, "BIB"); c.drawString(x + 2*mm, y + 1*mm, "CDTN")
        
        # --- LADO DIREITO ---
        DIV = 18 * mm 
        CAPA_X = x + DIV + 2*mm
        y_capa = y + 20*mm
        # Textos do Livro
        c.setFont("Helvetica-Bold", 7)
        c.drawString(CAPA_X, y_capa, f"N.chamada: {str(book.get('cdu', ''))}"); y_capa -= 3*mm
        
        c.setFont("Helvetica", 7)
        
        # --- PREPARA O AUTOR ---
        autor_str = str(book.get('autor', ''))
        # Se for maior que 45 letras, corta no 42 e adiciona os 3 pontinhos
        if len(autor_str) > 45:
            autor_str = autor_str[:42] + "..."
        c.drawString(CAPA_X, y_capa, autor_str); y_capa -= 3*mm
        
        # --- PREPARA O TÍTULO ---
        titulo_str = str(book.get('titulo', ''))
        if len(titulo_str) > 45:
            titulo_str = titulo_str[:42] + "..."
        c.drawString(CAPA_X, y_capa, titulo_str); y_capa -= 3*mm
        # ADICIONE ESTA LINHA: Garante que, se o título for muito longo, ele pare antes do código de barras
        y_capa = max(y_capa, y + 11*mm)
        
        # ADICIONE ESTA LINHA: Garante que, se o título for muito longo, ele pare antes do código de barras
        y_capa = max(y_capa, y + 11*mm)
        
        # --- CÓDIGO DE BARRAS E NÚMERO DO TOMBO ---
        tb = book.get('tombo', '')
        if tb:
            buffer_img = gerar_barcode_memoria(tb)
            
            # Subimos o código de barras para 5.5mm (estava em 4.5mm)
            c.drawImage(ImageReader(buffer_img), CAPA_X, y + 5.5*mm, width=22*mm, height=7*mm)
            
            # Subimos o número do tombo para 3.5mm (estava em 2.5mm)
            c.setFont("Helvetica", 5) 
            c.drawCentredString(CAPA_X + 11*mm, y + 3.5*mm, str(tb))
        
        # --- Rodapé da Capa (Agora não terá conflito) ---
        c.setFont("Helvetica-Bold", 5.5)
        c.drawString(CAPA_X, y + 1*mm, "BIBLIOTECA CDTN")
        c.drawRightString(x + 78*mm, y + 1*mm, "BiblioKhan")
        
        c.setFont("Helvetica-Bold", 5.5)
        c.drawString(CAPA_X, y + 1*mm, "BIBLIOTECA CDTN")
        c.drawRightString(x + 78*mm, y + 1*mm, "BiblioKhan")
        
        coluna += 1
        if coluna > 1: coluna = 0; linha += 1
        if linha > 9: c.showPage(); linha = 0; coluna = 0
    c.save()
    return arquivo_pdf

# --- INTERFACE ---
@ui.page('/')
def index():
    ui.colors(primary='#1E40AF', secondary='#F97316')
    
    with ui.column().classes('items-center w-full p-6'):
        # A linha (ui.row) vai conter a imagem E a coluna de texto lado a lado
        with ui.row().classes('items-center justify-center'):
            # Imagem com uma margem à direita (mr-4) para não grudar no texto
            ui.image("https://github.com/Biblio-Khan/etiqueta-cdtn/blob/main/Logotipo-CDTN.png?raw=true").classes('w-20 mr-4')
            
            # Coluna com o título e subtítulo
            with ui.column():
                ui.label("Sistema de Etiquetas: BIB CDTN").classes('text-h4 font-bold text-orange-600')
                ui.label("By BiblioKhan").classes('text-sm text-gray-500 italic')
            
        ui.separator().classes('w-full my-4')
        
        ui.label("Importar via Planilha").classes('text-h6')
        ui.upload(label="Selecione o arquivo Excel/CSV", on_upload=processar_upload, auto_upload=True).props('accept=".xlsx, .csv"')
        ui.separator().classes('w-full my-4')

    # 2. Formulário de Cadastro
    # Guardamos os inputs em variáveis locais da função index
    with ui.column().classes('w-full p-6'):
        with ui.grid(columns=4).classes('w-full'):
            cdu_in = ui.input("CDU")
            cutter_in = ui.input("Cutter")
            autor_in = ui.input("Autor")
            titulo_in = ui.input("Título")
            edicao_in = ui.input("Edição")
            exemplar_in = ui.input("Exemplar")
            patrimonio_in = ui.input("Patrimônio")
            tombo_in = ui.input("Tombo")
        
        # Função interna para capturar os valores e limpar os inputs
        def add():
            if not cdu_in.value: 
                ui.notify("O campo CDU é obrigatório!", type='warning')
                return
            
            novo_livro = {
                'cdu': cdu_in.value, 'cutter': cutter_in.value, 'autor': autor_in.value, 
                'titulo': titulo_in.value, 'edicao': edicao_in.value, 
                'exemplar': exemplar_in.value, 'patrimonio': patrimonio_in.value, 
                'tombo': tombo_in.value
            }
            biblioteca.append(novo_livro)
            salvar_dados(biblioteca)
            renderizar_lista()
            
            # Limpa os campos após adicionar
            for i in [cdu_in, cutter_in, autor_in, titulo_in, edicao_in, exemplar_in, patrimonio_in, tombo_in]:
                i.value = ''
        
        ui.button("Adicionar à Estante", on_click=add).classes('w-full mt-4')

    # 3. Container da Estante
    global lista_estante
    lista_estante = ui.column().classes('w-full p-6')
    
    # Renderiza imediatamente ao abrir a página
    renderizar_lista()
      
            
if __name__ in {"__main__", "__mp_main__"}:
    ui.run(
        port=8081, 
        reload=False, 
        show=False,
        title="Etiqueta Biblioteca CDTN",
        favicon='https://raw.githubusercontent.com/Biblio-Khan/etiqueta-cdtn/refs/heads/main/favicon.ico'
    )
