from pathlib import Path
import shutil
import sys


def consolidar_com_formato(template_path, input_dir, output_path, output_file):
    """
    Consolida arquivos Excel mantendo a formatação do template

    Args:
        template_path (str): Caminho do arquivo template
        input_dir (str): Diretório com arquivos para consolidar
        output_path (str): Diretório de saída
        output_file (str): Nome do arquivo de saída
    """
    try:
        # 1. Preparar caminhos
        saida = Path(output_path) / output_file

        # 2. Copiar template para arquivo final
        shutil.copy(template_path, saida)

        # 3. Carregar workbook de saída
        book = load_workbook(saida)

        # 4. Processar arquivos de entrada
        for arquivo in Path(input_dir).glob('*.xlsx'):
            if arquivo.name == Path(template_path).name:
                continue

            with pd.ExcelFile(arquivo) as xls:
                for sheet_name in xls.sheet_names:
                    if sheet_name not in book.sheetnames:
                        print(f"Aviso: Aba '{sheet_name}' ignorada (não existe no template)")
                        continue

                    # Ler dados (pula 1 linha de cabeçalho)
                    df = pd.read_excel(xls, sheet_name=sheet_name, skiprows=1)

                    if not df.empty:
                        sheet = book[sheet_name]
                        # Converter valores para evitar problemas com tipos
                        for row in df.itertuples(index=False, name=None):
                            sheet.append([str(item) if pd.isna(item) else item for item in row])

        # 5. Salvar alterações
        book.save(saida)
        print(f"Consolidação concluída! Arquivo salvo em: {saida}")
        return True

    except Exception as e:
        print(f"Erro na consolidação: {str(e)}", file=sys.stderr)
        return False

#
# if __name__ == "__main__":
#     # Exemplo de uso
#     # config = {
#     #     'template': 'caminho/template.xlsx',
#     #     'input_dir': 'caminho/entrada',
#     #     'output_dir': 'caminho/saida',
#     #     'template_type': 'consolidado.xlsx'
#     # }
#     #
#     # if consolidar_com_formato(
#     #         template_path=config['template'],
#     #         input_dir=config['input_dir'],
#     #         output_path=config['output_dir'],
#     #         output_file=config['template_type']
#     # ):
#     #     sys.exit(0)
#     # else:
#     #     sys.exit(1)