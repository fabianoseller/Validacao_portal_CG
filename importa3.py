import openpyxl
import re

def find_value(sheet, field):
    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and field in cell.value:
                row_values = [c.value for c in row]
                # Remove valores None e strings vazias
                row_values = [v for v in row_values if v is not None and v != '']
                # Retorna o valor após o campo, ignorando células vazias
                if len(row_values) > row_values.index(cell.value) + 1:
                    return row_values[row_values.index(cell.value) + 1]
    return None

# Caminho completo para o arquivo Excel
file_path = 'C:/Docs/Extrato Mensal.xlsx'

try:
    # Carrega o arquivo Excel
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    # Campos a serem extraídos
    fields = {
        'Empr.': 'Nome',
        'Vínculo:': 'Vínculo',
        'Cargo:': 'Cargo',
        'Salário:': 'Salário',
        'Líquido:': 'Líquido'
    }

    # Dicionário para armazenar os dados extraídos
    extracted_data = {}

    # Extrai os dados
    for field, key in fields.items():
        value = find_value(sheet, field)
        if value:
            # Remove espaços em branco extras e caracteres especiais
            if isinstance(value, str):
                value = re.sub(r'\s+', ' ', value).strip()
            extracted_data[key] = value

    # Imprime os dados extraídos
    print("\nDados extraídos:")
    for key, value in extracted_data.items():
        print(f"{key}: {value}")

    # Fecha o arquivo Excel
    workbook.close()

except FileNotFoundError:
    print(f"Erro: O arquivo '{file_path}' não foi encontrado.")
except Exception as e:
    print(f"Erro ao processar o arquivo: {e}")
