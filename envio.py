import tkinter as tk
from tkinter import ttk
from tkinter import simpledialog, messagebox
from tkcalendar import Calendar, DateEntry
from datetime import datetime
import mysql.connector
from mysql.connector import Error
import pandas as pd

# Criar janela Tkinter
root = tk.Tk()
root.title("Scanner de Códigos de Barras")

# Criar Treeview para exibir os dados em forma de tabela
tree = ttk.Treeview(root, columns=("ID", "Nota", "Rastreio", "Date", "Lote"), show="headings")
tree.heading("ID", text="ID")
tree.heading("Nota", text="Nota")
tree.heading("Rastreio", text="Rastreio")
tree.heading("Date", text="Date")
tree.heading("Lote", text="Lote")
tree.pack(fill="both", expand=True)

# Função para carregar os dados da tabela
def load_table():
    try:
        connection = mysql.connector.connect(
            host='192.168.1.11',
            database='scanner_codigos_barras',
            user='root',
            password='Luc@1081'
        )

        if connection.is_connected():
            cursor = connection.cursor()

            # Limpar a tabela
            tree.delete(*tree.get_children())

            # Selecionar todos os registros da tabela
            cursor.execute("SELECT * FROM codigos")
            rows = cursor.fetchall()

            # Preencher a tabela com os dados do banco de dados
            for row in rows:
                tree.insert("", tk.END, values=row)

    except Error as e:
        print("Erro ao conectar ao MySQL ou carregar dados da tabela:", e)
    finally:
        # Fechar a conexão apenas se ela estiver aberta
        if connection.is_connected():
            cursor.close()
            connection.close()

# Função para adicionar os códigos de barras e a data na tabela e no banco de dados MySQL
def add_to_table_and_database(nota, rastreio):
    try:
        connection = mysql.connector.connect(
            host='192.168.1.11',
            database='scanner_codigos_barras',
            user='root',
            password='Luc@1081'
        )

        if connection.is_connected():
            cursor = connection.cursor()

            # Verificar se a nota já existe na tabela
            cursor.execute("SELECT * FROM codigos WHERE Nota=%s", (nota,))
            existing_row = cursor.fetchone()
            if existing_row:
                print("Nota já existe na tabela. Não foi adicionada novamente.")
                return  # Sai da função se a nota já existir na tabela

            # Inserir os dados na tabela
            sql_insert_query = """INSERT INTO codigos (Nota, Rastreio) 
                                  VALUES (%s, %s)"""
            insert_tuple = (nota, rastreio)
            cursor.execute(sql_insert_query, insert_tuple)
            connection.commit()
            print("Código de barras inserido no banco de dados com sucesso")

            # Adicionar os dados à Treeview
            date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            tree.insert("", tk.END, values=(nota, rastreio, date))

    except Error as e:
        print("Erro ao conectar ao MySQL ou inserir dados:", e)
    finally:
        if connection.is_connected():
            cursor.close()
            connection.close()

# Função para lidar com a entrada do usuário
def handle_input(event):
    # Ler os códigos de barras
    nota = entry_nota.get()
    rastreio = entry_rastreio.get()

    # Verificar se os códigos foram escaneados
    if nota and rastreio:
        # Adicionar os códigos de barras na tabela e no banco de dados
        add_to_table_and_database(nota, rastreio)

        # Limpar as entradas
        entry_nota.delete(0, tk.END)
        entry_rastreio.delete(0, tk.END)

# Função para pesquisar por código de barras semelhante no banco de dados
def search_barcode():
    search_text = entry_search.get()
    if search_text:
        try:
            connection = mysql.connector.connect(
                host='192.168.1.11',
                database='scanner_codigos_barras',
                user='root',
                password='Luc@1081'
            )

            if connection.is_connected():
                cursor = connection.cursor()

                # Realizar a pesquisa no banco de dados
                cursor.execute("SELECT c.* FROM codigos c INNER JOIN lotes l ON c.Lote = l.id WHERE c.Nota LIKE %s OR c.Rastreio LIKE %s OR l.nome LIKE %s", ('%' + search_text + '%', '%' + search_text + '%', '%' + search_text + '%'))
                rows = cursor.fetchall()

                # Limpar a tabela
                tree.delete(*tree.get_children())

                # Preencher a tabela com os resultados da pesquisa
                for row in rows:
                    tree.insert("", tk.END, values=row)

        except Error as e:
            print("Erro ao conectar ao MySQL ou realizar pesquisa:", e)
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

# Função para excluir um item selecionado da tabela e do banco de dados
def delete_item():
    selected_item = tree.selection()
    if selected_item:
        item_id = tree.item(selected_item)['values'][0]  # Obtém o ID do item selecionado
        try:
            connection = mysql.connector.connect(
                host='192.168.1.11',
                database='scanner_codigos_barras',
                user='root',
                password='Luc@1081'
            )

            if connection.is_connected():
                cursor = connection.cursor()

                # Excluir o item do banco de dados
                cursor.execute("DELETE FROM codigos WHERE ID=%s", (item_id,))
                connection.commit()

                # Excluir o item da tabela
                tree.delete(selected_item)

        except Error as e:
            print("Erro ao conectar ao MySQL ou excluir item:", e)
        finally:
            if connection.is_connected():
                cursor.close()
                connection.close()

# Função para obter os nomes dos lotes do banco de dados
def get_lote_names_from_db():
    lote_names = []
    try:
        connection = mysql.connector.connect(
            host='192.168.1.11',
            database='scanner_codigos_barras',
            user='root',
            password='Luc@1081'
        )

        if connection.is_connected():
            cursor = connection.cursor()

            # Selecionar todos os nomes dos lotes
            cursor.execute("SELECT nome FROM lotes")
            rows = cursor.fetchall()

            # Preencher a lista com os nomes dos lotes
            for row in rows:
                lote_names.append(row[0])

    except Error as e:
        print("Erro ao conectar ao MySQL ou carregar nomes dos lotes:", e)
    finally:
        if connection.is_connected():
            cursor.close()
            connection.close()

    return lote_names

# Função para obter informações de exportação (lote e data)
def get_export_info():
    lote_names = get_lote_names_from_db()
    if not lote_names:
        print("Nenhum lote encontrado no banco de dados.")
        return None, None

    def on_submit():
        selected_lote = combo_lote.get()
        if not selected_lote:
            messagebox.showwarning("Aviso", "Selecione um lote.")
            return
        if selected_date.get() == 'Todos os dias':
            selected_date_value = 'Todos os dias'
        else:
            selected_date_value = cal.get_date()
        root_temp.selected_lote = selected_lote
        root_temp.selected_date = selected_date_value
        root_temp.destroy()

    root_temp = tk.Toplevel(root)
    root_temp.title("Selecionar Lote e Data")

    # Criar combobox para os lotes
    tk.Label(root_temp, text="Selecione o Lote:").pack(padx=10, pady=5)
    combo_lote = ttk.Combobox(root_temp, values=lote_names)
    combo_lote.pack(padx=10, pady=5)

    # Criar DateEntry para a data
    tk.Label(root_temp, text="Selecione a Data:").pack(padx=10, pady=5)
    cal = DateEntry(root_temp, date_pattern='dd/MM/yyyy')
    cal.pack(padx=10, pady=5)

    # Combobox para selecionar "Todos os dias"
    tk.Label(root_temp, text="Ou selecione:").pack(padx=10, pady=5)
    selected_date = ttk.Combobox(root_temp, values=['Todos os dias'])
    selected_date.pack(padx=10, pady=5)

    # Botão para submeter a seleção
    btn_submit = tk.Button(root_temp, text="Submit", command=on_submit)
    btn_submit.pack(pady=10)

    root_temp.wait_window()
    if hasattr(root_temp, 'selected_lote') and hasattr(root_temp, 'selected_date'):
        return root_temp.selected_lote, root_temp.selected_date
    return None, None

def export_to_excel():
    try:
        # Obter as informações de exportação do usuário
        lote_name, date = get_export_info()
        if not lote_name or not date:
            print("Nenhum nome de lote ou data fornecidos.")
            return

        # Conectar ao banco de dados e carregar os dados
        connection = mysql.connector.connect(
            host='192.168.1.11',
            database='scanner_codigos_barras',
            user='root',
            password='Luc@1081'
        )

        if connection.is_connected():
            cursor = connection.cursor()

            if date == 'Todos os dias':
                # Selecionar todos os registros da tabela para todos os lotes
                cursor.execute("SELECT Nota, Rastreio, Date_added FROM codigos")
            else:
                # Selecionar todos os registros da tabela para o lote e data especificados
                cursor.execute("SELECT Nota, Rastreio, Date_added FROM codigos WHERE Lote IN (SELECT id FROM lotes WHERE nome = %s) AND DATE(Date_added) = %s", (lote_name, date))
            rows = cursor.fetchall()

            # Criar um DataFrame Pandas com os dados
            df = pd.DataFrame(rows, columns=["Nota", "Rastreio", "Date_added"])

            # Exportar para um arquivo Excel, substituindo o arquivo existente
            date_str = date if date == 'Todos os dias' else date.strftime("%Y-%m-%d")
            df.to_excel(f"{lote_name}_{date_str}_dados_codigos.xlsx", index=False)

            print(f"Dados exportados para o arquivo Excel: {lote_name}_{date_str}_dados_codigos.xlsx")

    except Error as e:
        print("Erro ao conectar ao MySQL ou exportar para Excel:", e)
    finally:
        # Fechar a conexão apenas se ela estiver aberta
        if connection.is_connected():
            cursor.close()
            connection.close()

# Função para exportar os dados quando o botão de exportar é pressionado
def export_button_pressed():
    export_to_excel()

# Criar o botão de exportação
btn_export_excel = tk.Button(root, text="Exportar para Excel", command=export_button_pressed)
btn_export_excel.pack()

# Criar rótulos (labels) para indicar as caixas de entrada
label_nota = tk.Label(root, text="Nota:")
label_nota.pack(padx=10, pady=5)

# Caixa de entrada para o número da nota
entry_nota = tk.Entry(root)
entry_nota.pack(padx=10, pady=5)
entry_nota.bind("<Return>", handle_input)
entry_nota.focus()

# Rótulo para indicar a caixa de entrada do rastreio
label_rastreio = tk.Label(root, text="Rastreio:")
label_rastreio.pack(padx=10, pady=5)

# Caixa de entrada para o código de rastreio
entry_rastreio = tk.Entry(root)
entry_rastreio.pack(padx=10, pady=5)
entry_rastreio.bind("<Return>", handle_input)

# Caixa de entrada para pesquisa por código de barras
entry_search = tk.Entry(root)
entry_search.pack(side=tk.LEFT, padx=10, pady=5)

# Botão para acionar a pesquisa
btn_search = tk.Button(root, text="Search", command=search_barcode)
btn_search.pack(side=tk.LEFT, padx=5, pady=5)

# Botão para excluir o item selecionado
btn_delete = tk.Button(root, text="Delete", command=delete_item)
btn_delete.pack(side=tk.LEFT, padx=5, pady=5)

# Botão de atualização (refresh) para recarregar a tabela
btn_refresh = tk.Button(root, text="Refresh", command=load_table)
btn_refresh.pack()

# Carregar os dados da tabela ao iniciar o programa
load_table()

# Executar o loop principal da janela Tkinter
root.mainloop()