import tkinter as tk
from tkinter import ttk, Scrollbar, filedialog, messagebox
from pathlib import Path
import json


class Automation():
    def __init__(self):
        # Inicializa as variáveis da classe
        self.text_log = None
        # print('Pontx')
        self.config_file = Path("config.json")
        self.config_data = self.carregar_configuracoes()
        # print('Ponty')
        # Variáveis do Tkinter
        self.entrada_var = tk.StringVar(value=self.config_data.get("input_dir", ""))
        self.saida_var = tk.StringVar(value=self.config_data.get("output_dir"))
        self.template_var = tk.StringVar(value=self.config_data.get("template"))
        self.tipo_template_var = tk.StringVar(value=self.config_data.get("template_type", "Output.xlsx"))
        self.check_var = tk.IntVar(value=self.config_data.get("check"))
        self.merge_var = tk.IntVar(value=self.config_data.get("merge"))
        
        
    def procurar_entrada(self):
        print('Pontz')
        # print(self.entrada_var)
        # print(self.entrada_var.get())
        # self.entrada_var.set("Caminho 2")  # Antes de criar o Entry
        # print(self.entrada_var.get())
        try:
            path = filedialog.askdirectory(title="Selecione a pasta de entrada")
            if path:                
                self.entrada_var.set(path)            
                self.adicionar_log(f"Pasta de entrada selecionada: {path}")
        except Exception as e:
           self.adicionar_log(f"Erro ao selecionar pasta: {str(e)}", "erro")
        print("Valor definido:", self.entrada_var)  # Debug
        print("Valor definido:", self.entrada_var.get())  # Debug.get
        
    def procurar_saida(self):
        print('Pontw')
        try:
            path = filedialog.askdirectory(title="Selecione a pasta de saída")
            if path:
                self.saida_var.set(path)
                self.adicionar_log(f"Pasta de saída selecionada: {path}")
        except Exception as e:
            self.adicionar_log(f"Erro ao selecionar pasta: {str(e)}", "erro")
        print("Valor definido:", self.saida_var)  # Debug
        print("Valor definido:", self.saida_var.get())  # Debug.get
    
    def procurar_template(self):
        print('PontA')
        try:
            filetypes = [("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
            path = filedialog.askopenfilename(title="Selecione o template", filetypes=filetypes)
            if path:
                self.template_var.set(path)
                self.adicionar_log(f"Template selecionado: {path}")
        except Exception as e:
            self.adicionar_log(f"Erro ao selecionar template: {str(e)}", "erro")
    

    
    def salvar_config(self):
        print('PontD')
        config = {
            "input_dir": self.entrada_var.get(),
            "output_dir": self.saida_var.get(),
            "template": self.template_var.get(),
            "template_type": self.tipo_template_var.get(),
            "check": self.check_var.get(),
            "merge": self.merge_var.get()
                                                                        }
        # Validação
        errors = []
        if not Path(config["input_dir"]).is_dir():
            errors.append("Pasta de entrada inválida")
        if not Path(config["output_dir"]).is_dir():
            errors.append("Pasta de saída inválida")
        if not Path(config["template"]).is_file():
            errors.append("Arquivo template não encontrado")
        if errors:
            for error in errors:
                self.adicionar_log(error, "erro")
            return
        try:
            with open(self.config_file, 'w') as f:
                json.dump(config, f, indent=4)

            self.adicionar_log("Configurações salvas com sucesso!", "info")
            self.adicionar_log(f"Arquivo salvo : " + self.saida_var.get()+"//"+self.tipo_template_var.get())

        except Exception as e:
            self.adicionar_log(f"Erro ao salvar configurações: {str(e)}", "erro")
            
    def carregar_configuracoes(self):
        print('PontB')
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                self.adicionar_log("Configurações carregadas com sucesso")
                return config
        except Exception as e:
            self.adicionar_log(f"Erro ao carregar configurações: {str(e)}", "erro")
        # Valores padrão caso o arquivo não exista
        return {
            "input_dir": "",
            "output_dir": "",
            "template": "",
            "template_type": "Output.xlsx",
            "check": 0,
            "merge": 0,
        }
        print('PontC')
            
            
            
            
            
            

    def adicionar_log(self, mensagem, tipo="info"):
        """
        Adiciona uma mensagem ao log com base no tipo (info, erro, etc.).
        """
        if self.text_log:
            self.text_log.configure(state="normal")
            if tipo == "erro":
                self.text_log.insert("end", f"[ERRO] {mensagem}\n", "error")
            elif tipo == "info":
                self.text_log.insert("end", f"[INFO] {mensagem}\n", "info")
            else:
                self.text_log.insert("end", f"{mensagem}\n")
            self.text_log.configure(state="disabled")
        else:
            print(f"[{tipo.upper()}] {mensagem}")

