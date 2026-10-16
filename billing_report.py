import logging
import os
import sys
import xml.etree.ElementTree as ET

import pandas as pd

from xml_utils import (
    NFE_NS, build_error, build_summary, detect_layout, parse_text, to_float, write_errors_sheet,
)

# Configure the logging
logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s - %(message)s',
    datefmt='%H:%M:%S'
)

NAMESPACES = {'nfe': NFE_NS}

TYPE_MAP = {'0': 'Entrada', '1': 'Saída'}

FINALITY_MAP = {
    '1': 'Normal',
    '2': 'Complementar',
    '3': 'Ajuste',
    '4': 'Devolução'
}


def extract_cnpj_or_cpf(parent):
    if parent is None:
        return None
    cpf = parse_text(parent, 'nfe:CPF', NAMESPACES)
    cnpj = parse_text(parent, 'nfe:CNPJ', NAMESPACES)
    return cpf or cnpj


def process_xml_file(root):
    return {
        'Modelo': parse_text(root, './/nfe:ide/nfe:mod', NAMESPACES),
        'Número': parse_text(root, './/nfe:nNF', NAMESPACES),
        'Serie': parse_text(root, './/nfe:serie', NAMESPACES),
        'Tipo': TYPE_MAP.get(parse_text(root, './/nfe:tpNF', NAMESPACES), 'Desconhecido'),
        'Finalidade': FINALITY_MAP.get(parse_text(root, './/nfe:finNFe', NAMESPACES), 'Desconhecida'),
        'Natureza da Operação': parse_text(root, './/nfe:natOp', NAMESPACES),
        'Dt_Emissão': parse_text(root, './/nfe:dhEmi', NAMESPACES),
        'CPF_CNPJ_Emit': extract_cnpj_or_cpf(root.find('.//nfe:emit', NAMESPACES)),
        'Rz_Emit': parse_text(root, './/nfe:emit/nfe:xNome', NAMESPACES),
        'IE_Emit': parse_text(root, './/nfe:emit/nfe:IE', NAMESPACES),
        'CPF_CNPJ_Dest': extract_cnpj_or_cpf(root.find('.//nfe:dest', NAMESPACES)),
        'Rz_Dest': parse_text(root, './/nfe:dest/nfe:xNome', NAMESPACES),
        'IE_Dest': parse_text(root, './/nfe:dest/nfe:indIEDest', NAMESPACES),
        'UF_Dest': parse_text(root, './/nfe:dest/nfe:enderDest/nfe:UF', NAMESPACES),
        'Valor': to_float(parse_text(root, './/nfe:vNF', NAMESPACES)),
        'Chave': parse_text(root, './/nfe:chNFe', NAMESPACES),
        'Status': parse_text(root, './/nfe:xMotivo', NAMESPACES),
    }


def identify_note(root):
    return parse_text(root, './/nfe:nNF', NAMESPACES), parse_text(root, './/nfe:serie', NAMESPACES)


def process_folder(folder_path):
    all_data = []
    errors = []

    for filename in os.listdir(folder_path):
        if not filename.endswith('.xml'):
            continue

        file_path = os.path.join(folder_path, filename)
        root = None
        try:
            root = ET.parse(file_path).getroot()
            layout = detect_layout(root)

            if layout in ('nfse', 'evento'):
                continue  # Nota de serviço ou evento de NFS-e: tratados pelo nfse_report.py
            if layout != 'nfe':
                logging.warning(f'Ignorado (não é NF-e nem NFS-e): {filename}')
                continue

            logging.info(f'Processing: {file_path}')
            all_data.append(process_xml_file(root))
        except Exception as e:
            logging.error(f'Error processing {filename}: {e}')
            errors.append(build_error(filename, *identify_note(root), e))

    return all_data, errors


def write_model_sheets(writer, sub_df, detail_name, total_name):
    """Escreve a aba de detalhe (uma linha por nota) e a aba de total para um modelo."""
    sub_df.to_excel(writer, sheet_name=detail_name, index=False)
    build_summary(sub_df, 'Tipo', ['Valor']).to_excel(writer, sheet_name=total_name, index=False)


def save_to_excel(data, errors, output_path):
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        if data:
            df = pd.DataFrame(data)
            write_model_sheets(writer, df[df['Modelo'] == '55'], 'NF-e 55', 'Total NF-e 55')
            write_model_sheets(writer, df[df['Modelo'] == '65'], 'NFC-e 65', 'Total NFC-e 65')
        write_errors_sheet(writer, errors)


def main():
    if len(sys.argv) < 2:
        logging.warning("Uso: python billing_report.py <caminho_da_pasta>")
        return

    folder_path = sys.argv[1]
    data, errors = process_folder(folder_path)

    if data or errors:
        output_file = os.path.join(folder_path, 'notas_fiscais.xlsx')
        save_to_excel(data, errors, output_file)
        logging.info(f'\nArquivo salvo em: {output_file}')
        if errors:
            logging.warning(f'{len(errors)} arquivo(s) com erro — veja a aba "Erros" da planilha.')
    else:
        logging.warning("Nenhum dado encontrado.")


if __name__ == '__main__':
    main()
