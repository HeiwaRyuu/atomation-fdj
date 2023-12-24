import datetime as dt
from datetime import timedelta
import pyautogui
from src import *
import time
from bs4 import BeautifulSoup
import xlwings as xw
import pyperclip
import os
from playwright.sync_api import *
import tkinter as tk
import threading
from tkinter import messagebox

MAX_TRIES = 20
HIGH_CONFIDENCE=0.9
STANDARD_CONFIDENCE = 0.75
STANDARD_SLEEP = 1
STANDARD_PADY = 10
STANDARD_PADX = 10
STANDARD_WIDTH = 350
STANDARD_HEIGHT = 200
STANDARD_GEOMETRY = f'{STANDARD_WIDTH}x{STANDARD_HEIGHT}'

def fetch_first_day_month():
    today = dt.datetime.today()
    first_day = today.replace(day=1)
    return first_day

def move_to(start, how_far, where='x'):
    x = start[0]
    y = start[1]
    if where == 'x':
        pyautogui.moveTo(start[0]+how_far, start[1])
        x = start[0]+how_far
    else:
        pyautogui.moveTo(start[0], start[1]+how_far)
        y = start[1]+how_far
    return (x, y)

def parse_data(data):
    parsed_data = data.split('|')
    entrada = float(parsed_data[2].replace(' ', '').replace('.', '').replace(',', '.'))
    saida = float(parsed_data[4].replace(' ', '').replace('.', '').replace(',', '.'))
    return entrada, saida

def parse_data_tipos_de_pagamento(data):
    parsed_data = data.split(' ')
    parsed_data = list(filter(lambda a: a != '', parsed_data))
    print(parsed_data)
    parsed_data = float(parsed_data[1].replace(' ', '').replace('.', '').replace(',', '.'))
    return parsed_data
        
def write_in_excel(excel_file, entrada, saida, apuracao, venda, deposito, count):
    excel_app = xw.App(visible=True)
    excel_book = excel_app.books.open(excel_file)
    sheet = excel_book.sheets['Plan1']

    sheet[f'B{count+2}'].value = saida
    sheet[f'C{count+2}'].value = apuracao
    sheet[f'E{count+2}'].value = entrada
    sheet[f'F{count+2}'].value = venda
    sheet[f'G{count+2}'].value = deposito
    
    excel_book.save()
    excel_book.close()
    excel_app.quit()

def parse_apuracao(apuracao):
    apuracao = float(apuracao.replace(',', '.'))
    return apuracao

def inserir_data(date, skip_find=False, go_right=True):
    if skip_find:
        pyautogui.press('tab')

    pyautogui.typewrite(date.strftime('%d'))
    if go_right:
        pyautogui.press('right')
    pyautogui.typewrite(date.strftime('%m'))
    if go_right:
        pyautogui.press('right')
    pyautogui.typewrite(date.strftime('%Y'))

def delete_files(file_path):
    files = os.listdir(file_path)
    for file in files:
        print(file_path+file)
        os.unlink(file_path+file)
        
def get_dict_id(dct, item):
    for id, ele in enumerate(dct):
        if item == ele:
            print(ele)
            return id
    return False

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("AUTOMAÇÃO BIG")
        self.master.geometry(STANDARD_GEOMETRY)
        self.master.resizable(False, False)
        self.stop_threads = False
        self.create_widgets()

    def center(self):
        self.master.update_idletasks()
        width = self.master.winfo_width()
        height = self.master.winfo_height()
        x = (self.master.winfo_screenwidth() // 2) - (width // 2)
        y = (self.master.winfo_screenheight() // 2) - (height // 2)
        self.master.geometry('{}x{}+{}+{}'.format(width, height, x, y))

    def create_widgets(self):  
        self.store_to_start_from_label = tk.Label(self.master, text="Comecar a partir de:")
        self.store_to_start_from = tk.StringVar(self.master)
        options = ["PEDRO", "FERREIRA", "JUNDIAI", "EIRELI", "IF"]
        self.omlojas = tk.OptionMenu(self.master, self.store_to_start_from, *options)
        self.omlojas.config(width=20)
        self.run_single_checkbox_label = tk.Label(self.master, text="Rodar apenas uma loja?")
        self.run_single_checkbox_var = tk.IntVar()
        self.run_single_checkbox = tk.Checkbutton(self.master, variable=self.run_single_checkbox_var)
        self.days_to_feth_label = tk.Label(self.master, text="Dias para buscar:")
        self.days_to_feth_label_entry = tk.Entry(self.master)
        self.automation_button = tk.Button(self.master, text="COMEÇAR", command=self.start_automation)
        self.quit_btn = tk.Button(self.master, text="PARAR", command=self.quit)

        self.store_to_start_from_label.grid(row=0, column=0, pady=STANDARD_PADY, padx=STANDARD_PADX, sticky='w')
        self.omlojas.grid(row=0, column=1, pady=STANDARD_PADY, padx=STANDARD_PADX, sticky='w')
        self.run_single_checkbox_label.grid(row=1, column=0, pady=STANDARD_PADY, padx=STANDARD_PADX)
        self.run_single_checkbox.grid(row=1, column=1, pady=STANDARD_PADY, padx=STANDARD_PADX, sticky='w')
        self.days_to_feth_label.grid(row=2, column=0, pady=STANDARD_PADY, padx=STANDARD_PADX, sticky='w')
        self.days_to_feth_label_entry.grid(row=2, column=1, pady=STANDARD_PADY, padx=STANDARD_PADX, sticky='w')
        self.automation_button.grid(row=3, column=0, pady=STANDARD_PADY, padx=STANDARD_PADX, sticky='w')
        self.quit_btn.grid(row=3, column=1, pady=STANDARD_PADY, padx=STANDARD_PADX, sticky='w')
        
        self.set_widgets()
        self.center()

    def set_widgets(self):
        self.store_to_start_from.set("PEDRO")
        self.days_to_feth_label_entry.insert(0, '1')

    def start_automation(self):
        self.stop_threads = False
        new_thread = threading.Thread(target=self.automacao_big)
        new_thread.start()

    def find_img(self, img_path, img_name, click=True, confidence=STANDARD_CONFIDENCE):
        try_count = 0
        max_tries = MAX_TRIES
        while(try_count < max_tries):
            try:
                img = list(pyautogui.locateAllOnScreen(img_path, confidence=confidence))
                if img:
                    if img_name == 'escritorio' and len(img) > 2:
                        img = img[2]
                    elif img_name == 'escritorio':
                        img = img[1]
                    else:
                        img = img[0]
                    pyautogui.moveTo(img)
                    if click:
                        pyautogui.click()
                    time.sleep(STANDARD_SLEEP)
                    return img
            except Exception as e:
                print(f"Erro ao tentar buscar a imagem e clicar: {e}")
                break
            else:
                print(f'IMAGEM {img_name} NAO ENCONTRADA: TENTATIVA {try_count} de {max_tries}')
            try_count += 1
        messagebox.showerror("Erro", f"Erro ao tentar encontrar a imagem {img_name}!\nO programa foi finalizado!")
        self.stop_threads = True
        return False

    def save_file_as_htm(self, file_path, today):
        filename=f'{file_path}{today.strftime("%d-%m-%Y")}-1'
        extension='.htm'
        try:
            time.sleep(STANDARD_SLEEP*5)
            self.find_img(img_path=os.getcwd() + ARQUIVO_MENU, img_name='arquivo menu')
            self.find_img(img_path=os.getcwd() + SALVAR_COMO, img_name='salvar como btn')
            pyautogui.typewrite(filename)
            self.find_img(img_path=os.getcwd() + PADRAO_FILE_TYPE, img_name='padrao file type')
            self.find_img(img_path=os.getcwd() + WEB_PAGE_FILE_TYPE, img_name='web page file type')
            self.find_img(img_path=os.getcwd() + SAVE_FILE, img_name='save file')
            
            ## CLOSING SAVE FILE PAGE
            self.find_img(img_path=os.getcwd() + CLOSE_SAVE_FILE, img_name='close save file')
        except Exception as e:
            print(f"Erro ao salvar o file como html: {e}")
            messagebox.showerror("Erro", "Erro ao salvar o file:\n" + file_path.split('\\')[-1])
        return filename+extension
    
    def read_file_data_tipos_pagamento(self, file):
        try:
            total = False
            dinheiro = False
            with open(file, 'r') as f:
                webpage = f.read()
            soup = BeautifulSoup(webpage)
            filtered = soup.find_all('font', attrs = {'class' : 'f0'})
            total=0
            dinheiro=0
            for element in filtered:
                if 'Total: ' in element.text:
                    total = parse_data_tipos_de_pagamento(element.text)
                if 'DINHEIRO' in element.text:
                    dinheiro = parse_data_tipos_de_pagamento(element.text)
        except Exception as e:
            print(f"Erro ao tentar ler os dados de total e dinheiro: {e}")
            messagebox.showerror("Erro", "Erro ao ler os dados de total e dinheiro do arquivo:\n" + file.split('\\')[-1])
        return total, dinheiro
    
    def read_file_data(self, file):
        try:
            entrada = False
            saida = False
            with open(file, 'r') as f:
                webpage = f.read()
            soup = BeautifulSoup(webpage)
            filtered = soup.find_all('font', attrs = {'class' : 'f0'})
            entrada=0
            saida=0
            for element in filtered:
                if 'TOTAL --' in element.text:
                    entrada, saida = parse_data(element.text)
        except Exception as e:
            print(f"Erro ao tentar ler os dados entrada e saída: {e}")
            messagebox.showerror("Erro", "Erro ao ler os dados de entrada e saída do arquivo " + file.split('\\')[-1])
        return entrada, saida

    def fetch_lojas(self):
        starting_store = self.store_to_start_from.get()
        lojas = {"PEDRO":PEDRO, "FERREIRA":FERREIRA, "JUNDIAI":JUNDIAI, "EIRELI":EIRELI, "IF":IF}
        if self.run_single_checkbox_var.get():
            id = get_dict_id(lojas, starting_store)
            lojas = lojas[starting_store]
            lojas = [lojas]
        else:
            id = get_dict_id(lojas, starting_store)
            if id:
                lojas = list(lojas.values())[id:]
            else:
                lojas = list(lojas.values())
        return lojas, id
    
    def fetch_total_do_imposto(self):
        apuracao = False
        try_count = 0
        max_tries = MAX_TRIES
        while ((not (apuracao)) or (try_count < max_tries)):
            elemento = self.find_img(img_path=os.getcwd() + RECEITA_TOTAL, img_name='total do imposto', click=False)
            if elemento:
                last_post = move_to(start=elemento, how_far=40, where='x')
                move_to(start=last_post, how_far=32, where='y')
                print("MOUSE FINAL POS: ", pyautogui.position())
                print('double_click')
                pyautogui.doubleClick()
                pyautogui.hotkey('ctrl', 'c')
                apuracao = parse_apuracao(pyperclip.paste())
                break
            try_count += 1

        return apuracao

    def quit(self):
        self.stop_threads = True
        messagebox.showinfo("Programa Finalizado", "O programa foi finalizado pelo usuário!")

    def automacao_big(self):
        today=dt.datetime.today()
        file_path=os.getcwd() + '\\arquivos\\'
        excel_file=os.getcwd() + '\\dados_diarios.xlsx'
        count = 0
        lojas, starting_index = self.fetch_lojas()
        count += starting_index
        if not self.find_img(img_path=os.getcwd() + ICONE_SISTEMA_BIG, img_name='icone_sistema_big', click=False): return
        pyautogui.move(0, -20)
        pyautogui.click()
        if self.stop_threads:
            self.stop_threads = False
            return
        ## RESETANDO A INTERFACE PARA ESCRITORIO
        if not self.find_img(img_path=os.getcwd() + CASINHA, img_name='casinha'): return
        if self.stop_threads:
            self.stop_threads = False
            return
        if not self.find_img(img_path=os.getcwd() + DROPDOWN, img_name='dropdown', confidence=HIGH_CONFIDENCE): return
        if self.stop_threads:
            self.stop_threads = False
            return
        if not self.find_img(img_path=os.getcwd() + ESCRITORIO, img_name='escritorio'): return
        if self.stop_threads:
            self.stop_threads = False
            return
        if not self.find_img(img_path=os.getcwd() + OK_BTN, img_name='ok_btn'): return
        if self.stop_threads:
            self.stop_threads = False
            return

        for loja in lojas:
            ## NOTINHA COMPARATIVOS ENTRADAS E SAÍDAS
            if not self.find_img(img_path=os.getcwd() + CASINHA, img_name='casinha'): return
            if self.stop_threads:
                self.stop_threads = False
                return
            if not self.find_img(img_path=os.getcwd() + DROPDOWN, img_name='dropdown', confidence=HIGH_CONFIDENCE): return
            if self.stop_threads:
                self.stop_threads = False
                return
            if not self.find_img(img_path=os.getcwd() + loja, img_name=f'{loja}'): return
            if self.stop_threads:
                self.stop_threads = False
                return
            if not self.find_img(img_path=os.getcwd() + OK_BTN, img_name='ok_btn'): return
            if self.stop_threads:
                self.stop_threads = False
                return
            if not self.find_img(img_path=os.getcwd() + GERENCIAL_MENU, img_name='gerencial_menu'): return
            if self.stop_threads:
                self.stop_threads = False
                return
            self.find_img(img_path=os.getcwd() + ENTRADAS_E_COMPRAS, img_name='entradas_e_compras')
            if self.stop_threads:
                self.stop_threads = False
                return
            if not self.find_img(img_path=os.getcwd() + COMPARATIVO_DE_ENTRADAS, img_name='comparativo_entradas'): return
            if self.stop_threads:
                self.stop_threads = False
                return
            if not self.find_img(img_path=os.getcwd() + DATA_DA_NOTA, img_name='data da notas'): return
            if self.stop_threads:
                self.stop_threads = False
                return

            if not self.find_img(img_path=os.getcwd() + CALENDAR_BIG, img_name='data inicial big'): return
            pyautogui.click()
            pyautogui.typewrite(fetch_first_day_month().strftime('%d%m%Y'))

            if not self.find_img(img_path=os.getcwd() + CALENDAR_BIG, img_name='data final big'): return
            pyautogui.click()
            pyautogui.typewrite(today.strftime('%d%m%Y'))
            # Shortcut para visualizar
            pyautogui.press('F3')      
            if self.stop_threads:
                self.stop_threads = False
                return
            filename = self.save_file_as_htm(file_path, today)
            if self.stop_threads:
                self.stop_threads = False
                return
            entrada, saida = self.read_file_data(filename)
            if self.stop_threads:
                self.stop_threads = False
                return
            
            ## CLOSING COMPARATIVO
            if not self.find_img(img_path=os.getcwd() + CLOSE_COMPARATIVO, img_name='close comparativo window'): return
            if self.stop_threads:
                self.stop_threads = False
                return
            print("Pegando Apuração...")
            if not self.find_img(img_path=os.getcwd() + FISCAL_MENU, img_name='Fiscal'): return
            if self.stop_threads:
                self.stop_threads = False
                return
            if not self.find_img(img_path=os.getcwd() + ESCRITURAR, img_name='escriturar'): return
            if self.stop_threads:
                self.stop_threads = False
                return
            if not self.find_img(img_path=os.getcwd() + APURAR, img_name='apurar'): return
            if self.stop_threads:
                self.stop_threads = False
                return
            inserir_data(date=fetch_first_day_month())
            if self.stop_threads:
                self.stop_threads = False
                return
            inserir_data(date=today, skip_find=True)
            if self.stop_threads:
                self.stop_threads = False
                return
            if not self.find_img(img_path=os.getcwd() + PESQUISAR_APURACAO, img_name='pesquisar apuracao'): return
            if self.stop_threads:
                self.stop_threads = False
                return
            ## CALCULAR
            time.sleep(STANDARD_SLEEP*5)
            pyautogui.press('F3')
            apuracao = self.fetch_total_do_imposto()
            if self.stop_threads:
                self.stop_threads = False
                return
            if not self.find_img(img_path=os.getcwd() + CLOSE_APURACAO, img_name='fechar apuracao'): return
            if self.stop_threads:
                self.stop_threads = False
                return
            ## PEGANDO RELATORIO TIPOS DE PAGAMENTO
            if not self.find_img(img_path=os.getcwd() + FINANCEIRO_MENU, img_name='Financeiro Menu'): return
            if self.stop_threads:
                self.stop_threads = False
                return
            if not self.find_img(img_path=os.getcwd() + RELATORIO_TIPOS_DE_PAGAMENTO, img_name='Relatorio Tipos de Pagamento'): return
            if self.stop_threads:
                self.stop_threads = False
                return
            if not self.find_img(img_path=os.getcwd() + NAO_POP_UP, img_name='Relatorio Tipos de Pagamento'): return
            if self.stop_threads:
                self.stop_threads = False
                return
            day = int(self.days_to_feth_label_entry.get())
            inserir_data(date=today-timedelta(days=day), go_right=False)
            if self.stop_threads:
                self.stop_threads = False
                return
            inserir_data(date=today-timedelta(days=day), skip_find=True, go_right=False)
            if self.stop_threads:
                self.stop_threads = False
                return
            if not self.find_img(img_path=os.getcwd() + SINTETICO, img_name='Relatorio Tipos de Pagamento'): return
            if self.stop_threads:
                self.stop_threads = False
                return
            if not self.find_img(img_path=os.getcwd() + VIZUALIZAR_RELATORIO_TIPOS_DE_PAGAMENTO, img_name='Relatorio Tipos de Pagamento'): return
            if self.stop_threads:
                self.stop_threads = False
                return
            filename = self.save_file_as_htm(file_path, today)
            if self.stop_threads:
                self.stop_threads = False
                return
            if not self.find_img(img_path=os.getcwd() + CANCELAR_RELATORIO_TIPOS_DE_PAGAMENTO, img_name='Cancelar Relatorio Tipos de'): return
            if self.stop_threads:
                self.stop_threads = False
                return
            total, dinheiro = self.read_file_data_tipos_pagamento(filename)
            if self.stop_threads:
                self.stop_threads = False
                return
            ## ESCREVENDO NO EXCEL
            try:
                write_in_excel(excel_file, entrada, saida, apuracao, venda=total, deposito=dinheiro, count=count)
            except Exception as e:
                print(f"Excessão ao tentar escrever no Excel: {e}")
            if self.stop_threads:
                self.stop_threads = False
                return
            try:
                print(f"Entrada: {entrada}")
                print(f"Saida: {saida}")
                print(f"Apuração: {apuracao}")
                print(f"Processo concluído para loja {loja}!")
                count += 1
                delete_files(file_path)
                print("Arquivos deletados!")
                print(f"Processo concluído para a loja {loja}!")
            except Exception as e:
                print(f"Excessão ao tentar printar as variáveis finais: {e}")

        messagebox.showinfo("Processo concluído!", f"Processo concluído para todas as lojas!")

if __name__ == "__main__":
    root = tk.Tk()
    app = Application(master=root)
    app.mainloop()