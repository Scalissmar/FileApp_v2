from roby import Automation
from pathlib import Path
from gui import ConfigApp
import mac2
import mac1
import json



class Main:
	def __init__(self):
		self.config_app = ConfigApp()
		self.setup_application()
		# self.carregar_configuracoes()
		# self.mac()
			# = ConfigApp()
		
	def setup_application(self):
		# Configurações adicionais da aplicação podem vir aqui
		self.config_app.root.title("Aplicação Principal")

	def carregar_configuracoes(self):
		self.config_data = self.config_app.auto.carregar_configuracoes()
		print(self.config_data)
		
		self.ent = self.config_data.get("input_dir", "")
		self.sai = self.config_data.get("output_dir", "")
		self.tem = self.config_data.get("template", "")
		self.tfile = self.config_data.get("template_type")
		self.cx = self.config_data.get("check", 0)

	def mac (self):
		ent =  self.config_data.get("input_dir", "")
		sai = self.config_data.get("output_dir", "")
		tem = self.config_data.get("template", "")
		tfile = self.config_data.get("template_type", "Output.xlsx")
		cx = self.config_data.get("check", 0)

		print(self.ent)
		print(self.sai)
		print(self.tem)
		print(self.tfile)
		print(self.cx)

		# if cx = 1:
		# 	self.consolidar_com_formato(template_path=tem, input_dir=ent, output_path=sai, output_file=telFile)
		# else:
		# 	self.validar_e_consolidar(template_path=tem, input_dir=ent, output_path=sai, output_file=telFile)
	
	def run(self):
		self.config_app.root.mainloop()


if __name__ == "__main__":
	app = Main()
	app.run()
