from pathlib import Path
import shutil
from datetime import datetime
import sys
import os


def registrar_inconsistencia(writer, arquivo, aba, tipo_erro, detalhe):
    """Registra erros na aba de inconsistências de forma segura"""
    try:
        dados = {
            'Data_Hora': [datetime.now().strftime('%d/%m/%Y %H:%M:%S')],
            'Arquivo': [arquivo],
            'Aba': [aba],
            'Tipo_Erro': [tipo_erro],
            'Detalhe': [detalhe]
        }

        df_erro = pd.DataFrame(dados)

        # Verificar se a aba existe
        if 'Inconsistências' not in writer.sheets:
            df_erro.to_excel(writer, sheet_name='Inconsistências', index=False)
        else:
            df_erro.to_excel(
                writer,
                sheet_name='Inconsistências',
                index=False,
                header=False,
                startrow=writer.sheets['Inconsistências'].max_row
            )
    except Exception as e:
        print(f"Erro ao registrar inconsistência: {str(e)}")


def validar_e_consolidar(template_path, input_dir, output_path, output_file):
    # Converter para objetos Path e validar caminhos
    template_path = Path(template_path)
    input_dir = Path(input_dir)
    output_path = Path(output_path + "/" + output_file)

    # Validação inicial de caminhos
    if not template_path.exists():
        raise FileNotFoundError(f"Template não encontrado: {template_path}")

    if not input_dir.exists():
        raise NotADirectoryError(f"Diretório de entrada não existe: {input_dir}")

    # # Criar diretório de saída se necessário
    # output_path.parent.mkdir(parents=True, exist_ok=True)

    # Copiar template usando pandas para evitar problemas com links externos
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for sheet in pd.ExcelFile(template_path).sheet_names:
                df = pd.read_excel(template_path, sheet_name=sheet)
                df.to_excel(writer, sheet_name=sheet, index=False)
    except Exception as e:
        print(f"Erro ao copiar template: {str(e)}")
        # finally:
        # 	writer.save()  # Garante que o arquivo seja salvo e fechado
        # 	writer.close()  # Fecha explicitamente o writer
        sys.exit(1)

    # Carregar estrutura do template
    estrutura_template = ler_estrutura_template(template_path)

    # Processar arquivos
    with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        for arquivo in input_dir.glob('*.xlsx'):
            if arquivo.name == template_path.name:
                continue

            print(f"Processando: {arquivo.name}")

            # Validar arquivo
            erros = validar_arquivo(arquivo, estrutura_template)

            if erros:
                for erro in erros:
                    registrar_inconsistencia(writer, arquivo.name, *erro)
            else:
                consolidar_arquivo(arquivo, writer, estrutura_template)

    print("\nProcesso concluído com sucesso!")
    print(f"Arquivo consolidado: {output_path}")


def ler_estrutura_template(template_path):
    """Obtém a estrutura completa do template"""
    estrutura = {}
    try:
        with pd.ExcelFile(template_path) as xls:
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name, nrows=0)
                estrutura[sheet_name] = {
                    'colunas': list(df.columns),
                    'dtypes': df.dtypes.to_dict()
                }
    except Exception as e:
        print(f"Erro ao ler template: {str(e)}")
        sys.exit(1)
    return estrutura


def validar_arquivo(arquivo_path, estrutura_template):
    """Validação detalhada do arquivo"""
    erros = []
    try:
        with pd.ExcelFile(arquivo_path) as xls:
            # Verificar todas as abas do template
            for sheet_template in estrutura_template:
                if sheet_template not in xls.sheet_names:
                    erros.append((sheet_template, 'ABA FALTANTE', 'Aba não encontrada no arquivo'))

            # Verificar cada aba existente
            for sheet_name in xls.sheet_names:
                if sheet_name not in estrutura_template:
                    erros.append((sheet_name, 'ABA EXTRA', 'Aba não existe no template'))
                    continue

                # Validar estrutura
                df = pd.read_excel(xls, sheet_name=sheet_name, nrows=0)
                colunas_arquivo = list(df.columns)
                colunas_template = estrutura_template[sheet_name]['colunas']

                if colunas_arquivo != colunas_template:
                    diff = list(set(colunas_template) - set(colunas_arquivo))
                    erros.append((sheet_name, 'COLUNAS FALTANTES', f'Colunas ausentes: {diff}'))

        # Verificar tipos de dados (opcional)
        # df_full = pd.read_excel(xls, sheet_name=sheet_name)
        # for col, dtype in estrutura_template[sheet_name]['dtypes'].items():
        #     if df_full[col].dtype != dtype:
        #         erros.append((sheet_name, 'TIPO DE DADO', f'Coluna {col}: Esperado {dtype}, Encontrado {df_full[col].dtype}'))

    except Exception as e:
        erros.append(('GERAL', 'ERRO DE LEITURA', str(e)))

    return erros


def consolidar_arquivo(arquivo_path, writer, estrutura_template):
    """Consolida arquivos válidos"""
    try:
        with pd.ExcelFile(arquivo_path) as xls:
            for sheet_name in estrutura_template:
                df = pd.read_excel(xls, sheet_name=sheet_name)

                # Remover linhas vazias
                df = df.dropna(how='all')

                # Escrever dados
                df.to_excel(
                    writer,
                    sheet_name=sheet_name,
                    index=False,
                    header=False,
                    startrow=writer.sheets[sheet_name].max_row
                )
    except Exception as e:
        print(f"Erro ao consolidar {arquivo_path.name}: {str(e)}")

# # Configurações corrigidas
# template = r'C:\FTP_TRANSFER\IN\Template.xlsx'
# entrada = r'C:\FTP_TRANSFER\IN'
# saida = r'C:\FTP_TRANSFER\OUT\UniOutput.xlsx'
#
# # print(f"template: {template_path}")
# # print(f"entrada: {input_dir}")
# # print(f"saida: {output_path}")
#
# # # Execução segura
# # if __name__ == "__main__":
# # 	try:
# # 		validar_e_consolidar(template, entrada, saida)
# # 	except Exception as e:
# # 		print(f"Erro não tratado: {str(e)}")
# # 		sys.exit(1)