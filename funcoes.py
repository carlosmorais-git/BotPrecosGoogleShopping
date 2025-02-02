
import pandas as pd
import tkinter as tk
from tkinter import  messagebox,ttk
import webbrowser

# Bloco Formata Moeda
def formata_moeda(entry_variavel):
    '''Função que formata o valor digitado no entry como uma moeda (ex: 1234 -> 12,34)'''
    
    valor = entry_variavel.get().replace(",", "").replace(".", "")
    
    # Verifica se o valor é numérico
    if valor.isdigit():
        valor = int(valor)
        # Formata o valor como moeda (2 casas decimais, separador de milhar)
        formato_val = f"{valor // 100},{valor % 100:02d}"
        entry_variavel.set(formato_val)

        # # Move o cursor para o final do texto
        # entry.icursor(tk.END)

    # Se o valor estiver vazio, define como "0,00"
    elif not valor:
        entry_variavel.set("0,00")
        # entry.icursor(tk.END)
        
def limpar_click(event, entry_variavel):
    if entry_variavel.get() == "0,00":
        entry_variavel.set("")

def voltar_padrao(event, entry_variavel):
    '''Função que restaura o valor padrão "0,00" se o entry ficar vazio ao perder o foco'''
    if not entry_variavel.get():
        entry_variavel.set("0,00")

def iniciar_no_final(event,entry):
    
    '''Garante que a digitação ocorra sempre no final do texto'''
    if event.keysym in ("BackSpace", "Delete"):
        # Posiciona o cursor sempre no final do texto
        entry.icursor(tk.END)
        return
    else:
        # Insere o caractere no final
        entry.insert(tk.END, event.char)
    
    # Posiciona o cursor sempre no final do texto
    entry.icursor(tk.END)
    return "break"  # Evita o comportamento padrão de mover o cursor

def formatar_entry(entry):
    '''Garante que toda entrada de texto será formatada por moedas'''
    entry_variavel = tk.StringVar(value="0,00")

    # Vincula a StringVar ao Entry
    entry.config(textvariable=entry_variavel)
    entry.bind("<FocusIn>", lambda event: limpar_click(event, entry_variavel))
    entry.bind("<Key>", lambda event: iniciar_no_final(event, entry))
    entry.bind("<FocusOut>", lambda event: voltar_padrao(event, entry_variavel))
   
    # Configura um callback para ser executado sempre que a variavel associada for modificada.
    # O metodo trace_add monitora mudanças no valor da variavel do tipo StringVar.
    # Quando a variavel e alterada, a funcao formata_moeda e chamada, passando entry_variavel como argumento.
    entry_variavel.trace_add("write", lambda *args: formata_moeda(entry_variavel))


# Posiciona Entry e Botoes da tela
def posicionar_OBJ_tela(valor_x,valor_y,w,h,padding=None):
    '''Configura a posicao da entrada de texto de entry e posicionamento dos botoes'''
    if padding:# Se for entry
        text_x = valor_x + padding
        text_y = valor_y + padding
        text_largura = w - 2 * padding
        text_altura = h - 2 * padding
        
    else:# Se for button
        text_x = valor_x 
        text_y = valor_y
        text_largura = w + 10
        text_altura = h + 10

    return (text_x,text_y,text_largura,text_altura)

# Centralizar a janela
def centralizar_janela(janela, largura, altura,up=None):
    '''Centraliza a exibcao de uma tela'''
    x, y = largura, altura
    janela.geometry(f"{x}x{y}+{int((janela.winfo_screenwidth() / 2) - (x / 2))}+{int((janela.winfo_screenheight() / 2) - (y / 2))}")

def remove_formatacao_preco(preco):
    '''Remove todos os caracteres que não são números, ponto, vírgula ou sinal negativo e converte para ponto flutuante'''
    
    # Converte o valor para string
    preco = str(preco)
   
    # Remove todos os caracteres que não são números, ponto, vírgula ou sinal negativo
    # preco = re.sub(r'[^\d,.-]', '', preco)

    # Substitui o ponto (separador de milhares) por uma string vazia e ajusta a vírgula como decimal
    return float(preco.replace(',', '').replace('R$',''))

## Converter um arquivo xlsx para HTML
def exibir_meus_resultados_na_web(tabela_resultados):
   
    try:
        # Verificar se a entrada é um DataFrame ou um caminho para arquivo Excel
        

        if 'DataFrame' in str(type(tabela_resultados)):
            tabela = tabela_resultados # Já é um DataFrame

        else:
            tabela = pd.read_excel(r'{}'.format(tabela_resultados)) # Ler o arquivo Excel
                
        tabela['Preco'] = tabela['Preco'].apply(remove_formatacao_preco)

        tabela['Preco'] = tabela['Preco'].apply('R${:,.2f}'.format)

       
        # Verifica se a coluna "Link" existe
        if "Link" in tabela.columns:
            # Formata apenas os links que ainda não estão no formato de link HTML
            tabela['Link'] = tabela['Link'].apply(
                lambda x: f'<a href="{x}" target="_blank">Ver Produto</a>' if isinstance(x, str) and not x.startswith('<a href') else x
            )


        # Converter um DataFrame para HTML
        tabela_html = tabela.to_html(index=False, escape=False, border=1)

        # CSS para estilizar a tabela
        css_styles = """
        <style>
        
        body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 12px;
            background-color: #fff;
        }
        header {
            background: #FDFCE6;
            color: #fff;
            padding: 1em;
            text-align: center;
        }

        table {
            width: 100%;
            border-collapse: collapse;
        }
        table, th, td {
            border: 1px solid #ddd;
            }

        th, td {
            padding: 8px;
            text-align: left;
        }
            
        a {
            color: blue;
            text-decoration: none;
        }
        a:hover {
            text-decoration: underline;
        }
        </style>
        """

        # Montar HTML final com CSS
        # Esse trecho monta o meu arquivo html já exibindo todos os meus produtos 
        html_content = f"""
        <html>
            <head>
                <meta charset="utf-8">
                <title>Tabela Resultados</title>
                {css_styles}
            </head>
            <body>
                <header>
                       <img src="imagens/banner_gloogle-shope.png" alt="Imagem local" width="480" height="130">
                </header>
                <h2 style="text-align: left;">Resultados da Busca</h2>
                {tabela_html}
                
            </body>
        </html>
        """

        # Salvar como arquivo HTML
        arquivo_html = "base_html_resultados_.html"
        with open(arquivo_html, "w", encoding="utf-8") as f:
            f.write(html_content)

        # Abrir o HTML no navegador
        webbrowser.open(arquivo_html)
    
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar o arquivo: {str(e)}")

def limpar_tree(tree):
    for i in tree.get_children():        
        tree.delete(i)

def converte_float(numero):
    return float(numero.replace(',','.'))


def listar_tree(treeview_listar_produto):

    style = ttk.Style()
    style.configure("Custom.Treeview", font=("Arial", 10),
                    background='white',foreground="black",fieldbackground="red")
    style.configure("Treeview.Heading", font = "Arial 8")

    treeview_listar_produto.config(style='Custom.Treeview')
    valores_inserir_lista = [treeview_listar_produto.item(linha)["values"] for linha in treeview_listar_produto.get_children()]
      

    limpar_tree(treeview_listar_produto)

    # modificar o cabeçario dinamico
    colunasProdutos = ['Produtos','Min','Max']
    largura_colunas_treeview = [140,60,60]
    treeview_listar_produto.config(columns=list(colunasProdutos),show="headings")

    # cor linhas treeview
    treeview_listar_produto.tag_configure('oddrow', background='white')
    treeview_listar_produto.tag_configure('evenrow', background='#D9D9D9')

    for i,coluna in enumerate(colunasProdutos):
        treeview_listar_produto.heading(coluna, text=coluna)
        treeview_listar_produto.column(coluna,width=largura_colunas_treeview[i],anchor='center')
    

    # Adiciona cor as linhas do Treeview
    for index,valor in enumerate(valores_inserir_lista):

        if index % 2 == 0:# ocorrencia de repeticao 2 em 2 ou 3 em 3
            treeview_listar_produto.insert("", "end", values=(valor[0],converte_float(valor[1]),converte_float(valor[2])), tags=('evenrow',))
        else:
            treeview_listar_produto.insert("", "end", values=(valor[0], converte_float(valor[1]),converte_float(valor[2])), tags=('oddrow',))

