import sys
import tkinter as tk
from tkinter import ttk, Scrollbar
# from PIL import Image, ImageTk  # Importação correta
from tkinter import filedialog, messagebox
from pathlib import Path
import json
from datetime import datetime
from roby import Automation


class ConfigApp():
    def __init__(self):
        super().__init__()
        self.root = tk.Tk()
       
        print('Pont11')
        self.entrada_var = tk.StringVar()
        print('Pont12')
        self.saida_var = tk.StringVar()
        # self.saida_var.set("Caminho padrão")  # Antes de criar o Entry
        self.template_var = tk.StringVar()
        # self.template_var.set("Caminho padrão")  # Antes de criar o Entry
        print('Pont13')
        self.tipo_template_var = tk.StringVar(value="Escolha um nome de saida")
        print('Pont14')
        self.tela()
        self.frame()
        self.widgets()
        self.auto = Automation()
        # self.carregar_configuracoes()
        # self.entrada_var.set(self.ent)  # Antes de criar o Entry
    
    # def carregar_configuracoes(self):
    #     self.config_data = self.config_app.auto.carregar_configuracoes()
    #     print(self.config_data)
    #
    #     self.ent = self.config_data.get("input_dir", "")
    #     self.sai = self.config_data.get("output_dir", "")
    #     self.tem = self.config_data.get("template", "")
    #     self.tfile = self.config_data.get("template_type")
    #     self.cx = self.config_data.get("check", 0)

    def tela(self):
        self.root.title("FileMerge")
        self.root.configure(background="darkblue")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        self.style = ttk.Style()
        self.style.configure("Gold.TButton", background="gold", foreground="black", font=("arial", 6, "bold"))
        self.root.maxsize(900, 700)
        self.root.minsize(700, 400)

    def frame(self):
        self.frame_1 = tk.Frame(self.root, bd=4, bg="lightgray", highlightbackground="gold", highlightthickness=3)
        self.frame_1.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.46)
        self.frame_2 = tk.Frame(self.root, bd=4, bg="white", highlightbackground="black", highlightthickness=2)
        self.frame_2.place(relx=0.02, rely=0.5, relwidth=0.96, relheight=0.46)
        self.text_log = tk.Text(self.frame_2, bg="white", fg="black", wrap=tk.WORD, state="disabled")
        self.text_log.place(relx=0.01, rely=0.01, relwidth=0.98, relheight=0.98)
        self.text_log.tag_config("timestamp", foreground="gray")
        self.text_log.tag_config("success", foreground="green")
        self.text_log.tag_config("error", foreground="red")
        self.text_log.tag_config("warning", foreground="orange")
        self.text_log.tag_config("info", foreground="blue")

    def widgets(self):
        print('Pont1')
        print(self.entrada_var)
        self.btn_sourcein = ttk.Button(self.frame_1, text="FINDING...", style="Gold.TButton", command=self.entrada )
        self.btn_sourcein.place(relx=0.58, rely=0.1, relwidth=0.1, relheight=0.1)
        self.lbl_sourcein = tk.Label(self.frame_1, text="SOURCE IN:", bd=2, bg="darkgray", font=("arial", 6, "bold"),highlightbackground="gold", highlightthickness=2)
        self.lbl_sourcein.place(relx=0.05, rely=0.1, relwidth=0.15, relheight=0.1)
        self.en_sourcein = ttk.Entry(self.frame_1,  width=50,textvariable=self.entrada_var)
        self.en_sourcein.place(relx=0.18, rely=0.1, relwidth=0.4, relheight=0.1)
        # self.root.update()
        
        print('Pont2')
        print(self.saida_var)
        self.btn_sourceout = ttk.Button(self.frame_1, text="FINDING...", style="Gold.TButton", command=self.saida)
        self.btn_sourceout.place(relx=0.58, rely=0.3, relwidth=0.1, relheight=0.1)
        self.lbl_sourceout = tk.Label(self.frame_1, text="SOURCE OUT:", bd=2, bg="darkgray", font=("arial", 6, "bold"),highlightbackground="gold", highlightthickness=2) 
        self.lbl_sourceout.place(relx=0.05, rely=0.3, relwidth=0.15, relheight=0.1)
        self.en_sourceout = ttk.Entry(self.frame_1, width=50, textvariable=self.saida_var)
        self.en_sourceout.place(relx=0.18, rely=0.3, relwidth=0.4, relheight=0.1)
        
        print('Pont3')
        print(self.template_var)
        self.btn_template = ttk.Button(self.frame_1, text="FINDING...", style="Gold.TButton",command=self.template)
        self.btn_template.place(relx=0.58, rely=0.5, relwidth=0.1, relheight=0.1)
        self.lbl_template = tk.Label(self.frame_1, text="TEMPLATE:", bd=2, bg="darkgray", font=("arial", 6, "bold"), highlightbackground="gold", highlightthickness=2)
        self.lbl_template.place(relx=0.05, rely=0.5, relwidth=0.15, relheight=0.1)
        self.en_template = ttk.Entry(self.frame_1, width=50, textvariable=self.template_var)
        self.en_template.place(relx=0.18, rely=0.5, relwidth=0.4, relheight=0.1)
        
        print('Pont4')
        print(self.tipo_template_var)
        self.lbl_output = tk.Label(self.frame_1, text="OUTPUT:", bd=2, bg="darkgray", font=("arial", 6, "bold"),highlightbackground="gold", highlightthickness=2)
        self.lbl_output.place(relx=0.05, rely=0.7, relwidth=0.15, relheight=0.1)
        self.en_output = ttk.Combobox(self.frame_1, values=["UniOutput.xlsx", "EMPTY_1", "EMPTY_2", "EMPTY_3", "EMPTY_4", "EMPTY_5"],textvariable=self.tipo_template_var)
        self.en_output.place(relx=0.18, rely=0.7, relwidth=0.4, relheight=0.1)
        
        print('Pont5')
        self.btn_salvar = ttk.Button(self.frame_1, text="SALVAR", style="Gold.TButton", command=self.salvar)
        self.btn_salvar.place(relx=0.3, rely=0.85, relwidth=0.1, relheight=0.1)

        self.btn_limpar = ttk.Button(self.frame_1, text="LIMPAR", style="Gold.TButton", command=self.limpar_tela)
        self.btn_limpar.place(relx=0.1, rely=0.85, relwidth=0.1, relheight=0.1)
        print('Pont7')
        self.btn_cancelar = ttk.Button(self.frame_1, text="CANCELAR", style="Gold.TButton", command=self.fecha)
        self.btn_cancelar.place(relx=0.5, rely=0.85, relwidth=0.1, relheight=0.1)

        self.merge_var = tk.IntVar()  # Variável para o estado do Merge
        self.check_var = tk.IntVar()  # Variável para o estado do Check
        print('Pont8')
        self.ck_merge = tk.Checkbutton(self.frame_1, text="MERGE", variable=self.merge_var, onvalue=1, offvalue=0, bd=2,bg="darkgray", font=("arial", 6, "bold"), selectcolor="gold", activebackground="gold", activeforeground="white")
        self.ck_merge.place(relx=0.75, rely=0.2, relwidth=0.1, relheight=0.1)
        self.ck_check = tk.Checkbutton(self.frame_1, text="CHECK & MERGE", variable=self.check_var, onvalue=1,offvalue=0, bd=2, bg="darkgray", font=("arial", 6, "bold"), selectcolor="gold",activebackground="gold", activeforeground="white")
        self.ck_check.place(relx=0.75, rely=0.4, relwidth=0.15, relheight=0.1)
       
        self.text_log.insert("end", "Aqui está uma mensagem de log.\n", "timestamp")
        self.scrollbar = Scrollbar(self.frame_2, orient="vertical")
        self.text_log.configure(yscroll=self.scrollbar.set)
        self.scrollbar.place(relx=0.96, rely=0.1, relwidth=0.03, relheight=0.85)
        print('Pont9')
        
        
    def limpar_tela(self):
        print('Pont10')
        self.en_sourcein.delete(0, tk.END)
        self.en_sourceout.delete(0, tk.END)
        self.en_template.delete(0, tk.END)
        self.en_output.delete(0, tk.END)
        self.merge_var.set(0)
        self.check_var.set(0)
        

    def fecha(self):
        print('Pont11')
        """Fecha a aplicação com confirmação e tratamento de erros"""
        try:
            if messagebox.askyesno("Sair", "Deseja realmente sair?"):
                if hasattr(self, 'pre_fechamento'):
                    self.pre_fechamento()  # Método para limpeza
                self.root.destroy()
                sys.exit(0)

        except KeyboardInterrupt:
            print("\nOperação cancelada pelo usuário")
            sys.exit(0)

        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao fechar: {e}")
            sys.exit(1)
            
    
    def entrada(self):
        self.auto.procurar_entrada()
        self.entrada_var.set(self.auto.entrada_var.get())
    
    def saida(self):
        self.auto.procurar_saida()
        self.saida_var.set(self.auto.saida_var.get())
    
    def template(self):
        self.auto.procurar_template()
        self.template_var.set(self.auto.template_var.get())
    
    def salvar(self):
        try:
            self.auto.salvar_config()
        except Exception as e:
            print( {str(e)})