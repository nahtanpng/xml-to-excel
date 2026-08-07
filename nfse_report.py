import logging
import os
import sys
import xml.etree.ElementTree as ET

import pandas as pd
from xlsxwriter.utility import xl_col_to_name

from xml_utils import NFSE_NS, build_summary, detect_layout, parse_text, to_float

# Configure the logging
logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s - %(message)s',
    datefmt='%H:%M:%S'
)

NAMESPACES = {'nfse': NFSE_NS}

# CSRF = contribuições sociais retidas na fonte: CSLL 1% + PIS 0,65% + COFINS 3%.
# Só dá para separar nas três alíquotas se o total retido bater com 4,65% do
# faturamento da nota; fora disso o valor sai destacado para conferência manual.
CSRF_RATE = 0.0465
CSRF_TOLERANCE = 0.01  # diferença aceita, em reais

VALUE_COLUMNS = ['VALOR', 'ISS', 'INSS', 'CSLL', 'PIS', 'COFINS', 'CSRF', 'IRR', 'VALOR LIQUIDO']
DETAIL_COLUMNS = [
    'NF', 'STATUS', 'PRESTADOR', 'TOMADOR', 'DESCRIÇÃO', 'SERVIÇO', 'VALOR', 'ISS', 'RET',
    'INSS', 'CSLL', 'PIS', 'COFINS', 'CSRF', 'IRR', 'VALOR LIQUIDO',
]


def extract_federal_taxes(dps):
    """Lê o grupo tribFed. Notas sem retenção federal não trazem esse grupo."""
    trib_fed = dps.find('nfse:valores/nfse:trib/nfse:tribFed', NAMESPACES)
    if trib_fed is None:
        return {'INSS': 0.0, 'CSLL': 0.0, 'PIS': 0.0, 'COFINS': 0.0, 'CSRF': 0.0, 'IRR': 0.0}

    # O leiaute novo não traz mais um flag de retenção dentro de <piscofins> — vPis/vCofins
    # vêm sempre crus. vRetCSLL é o único campo com "Ret" no nome (valor retido, não só
    # apurado) e os três tributos sempre vêm juntos quando há retenção, então ele serve de
    # indicador confiável para os três.
    csll = to_float(parse_text(trib_fed, 'nfse:vRetCSLL', NAMESPACES))
    pis = to_float(parse_text(trib_fed, 'nfse:piscofins/nfse:vPis', NAMESPACES))
    cofins = to_float(parse_text(trib_fed, 'nfse:piscofins/nfse:vCofins', NAMESPACES))

    if csll == 0.0:
        pis = cofins = 0.0

    return {
        'INSS': to_float(parse_text(trib_fed, 'nfse:vRetCP', NAMESPACES)),
        'CSLL': csll,
        'PIS': pis,
        'COFINS': cofins,
        # Somado depois do filtro de retenção: o CSRF tem que refletir o que
        # realmente vai para a planilha, não o que está cru no XML.
        'CSRF': csll + pis + cofins,
        'IRR': to_float(parse_text(trib_fed, 'nfse:vRetIRRF', NAMESPACES)),
    }


def extract_parties(inf, dps):
    """Nome do prestador (emit, em infNFSe — infDPS/prest só traz CNPJ) e do tomador
    (toma, em infDPS; opcional na nota)."""
    return {
        'PRESTADOR': parse_text(inf, 'nfse:emit/nfse:xNome', NAMESPACES),
        'TOMADOR': parse_text(dps, 'nfse:toma/nfse:xNome', NAMESPACES),
    }


def extract_service_tax(inf, dps):
    """ISS calculado (infNFSe/valores/vISSQN) e se foi retido pelo tomador (tpRetISSQN=2)."""
    retention = parse_text(dps, 'nfse:valores/nfse:trib/nfse:tribMun/nfse:tpRetISSQN', NAMESPACES)
    return {
        'ISS': to_float(parse_text(inf, 'nfse:valores/nfse:vISSQN', NAMESPACES)),
        'RET': 'Retido' if retention == '2' else 'Não retido',
    }


def detect_cancellation_target(root):
    """Se o XML for um evento de cancelamento de NFS-e, devolve o nNFSe da nota cancelada.

    O tipo do evento fica no nome de um elemento que varia por motivo (ex.: <e101101>), então
    o critério é o texto de <xDesc> em vez do nome da tag.
    """
    desc = parse_text(root, './/nfse:xDesc', NAMESPACES)
    if not desc or 'cancelamento' not in desc.lower():
        return None
    return parse_text(root, 'nfse:infEvento/nfse:nDFSe', NAMESPACES)


def apply_cancellation(entry):
    """Zera os valores em reais e marca a nota como cancelada. Colunas descritivas
    (PRESTADOR, TOMADOR, DESCRIÇÃO, SERVIÇO, RET) continuam preenchidas para contexto.
    """
    for column in VALUE_COLUMNS:
        entry[column] = 0.0
    entry['STATUS'] = 'Cancelada'


def apply_csrf_split(taxes, value, note_number):
    """Decide se o CSRF pode ser separado em CSLL/PIS/COFINS.

    Bate com 4,65% do faturamento: mantém os três valores do XML (fiéis à nota e
    já somando exatamente o CSRF). Não bate: zera os três e deixa só o CSRF
    cheio, que sai destacado na planilha para conferência manual.

    Devolve True quando a nota está conforme.
    """
    csrf = taxes['CSRF']
    if csrf == 0.0:
        return True  # nota sem retenção de CSRF — não é divergência

    expected = round(value * CSRF_RATE, 2)
    if abs(round(csrf, 2) - expected) <= CSRF_TOLERANCE:
        return True

    logging.warning(
        f'NF {note_number}: CSRF de {csrf:.2f} não bate com 4,65% de {value:.2f} '
        f'(esperado {expected:.2f}). Importada sem separar CSLL/PIS/COFINS.'
    )
    taxes['CSLL'] = taxes['PIS'] = taxes['COFINS'] = 0.0
    return False


def process_xml_file(root):
    # Atenção: <valores> aparece 4x neste leiaute (em infNFSe, em infDPS e nos
    # dois blocos IBSCBS). Todos os caminhos abaixo são ancorados de propósito —
    # um './/' pegaria o valor errado.
    inf = root.find('nfse:infNFSe', NAMESPACES)
    dps = inf.find('nfse:DPS/nfse:infDPS', NAMESPACES)

    note_number = parse_text(inf, 'nfse:nNFSe', NAMESPACES)
    value = to_float(parse_text(dps, 'nfse:valores/nfse:vServPrest/nfse:vServ', NAMESPACES))
    taxes = extract_federal_taxes(dps)
    csrf_ok = apply_csrf_split(taxes, value, note_number)
    competence = parse_text(dps, 'nfse:dCompet', NAMESPACES) or ''

    # Somar taxes.values() subtrairia CSLL/PIS/COFINS duas vezes, já que o CSRF
    # engloba os três. O CSRF também é o valor certo na nota divergente, onde as
    # três colunas saem zeradas mas a retenção aconteceu.
    retained = taxes['INSS'] + taxes['CSRF'] + taxes['IRR']

    return {
        'NF': note_number,
        'STATUS': 'Normal',
        **extract_parties(inf, dps),
        'DESCRIÇÃO': parse_text(dps, 'nfse:serv/nfse:cServ/nfse:xDescServ', NAMESPACES),
        'SERVIÇO': parse_text(inf, 'nfse:xTribMun', NAMESPACES),
        'VALOR': value,
        **extract_service_tax(inf, dps),
        **taxes,
        'VALOR LIQUIDO': value - retained,
        # Colunas auxiliares: não vão para a aba de detalhe
        'Mês Competência': competence[:7],
        'Dt_Emissão': parse_text(dps, 'nfse:dhEmi', NAMESPACES),
        '_csrf_ok': csrf_ok,
    }


def process_folder(folder_path):
    note_entries = []
    canceled_notes = set()

    for filename in os.listdir(folder_path):
        if not filename.endswith('.xml'):
            continue

        file_path = os.path.join(folder_path, filename)
        try:
            root = ET.parse(file_path).getroot()
            layout = detect_layout(root)

            if layout == 'nfe':
                continue  # Nota de produto: tratada pelo billing_report.py
            if layout == 'evento':
                target = detect_cancellation_target(root)
                if target:
                    canceled_notes.add(target)
                continue
            if layout != 'nfse':
                logging.warning(f'Ignorado (não é NFS-e nem NF-e): {filename}')
                continue

            logging.info(f'Processing: {file_path}')
            note_entries.append(process_xml_file(root))
        except Exception as e:
            logging.error(f'Error processing {filename}: {e}')

    # Duas fases porque o evento de cancelamento pode aparecer antes ou depois do XML
    # da nota na pasta — não dá para marcar durante o primeiro loop.
    for entry in note_entries:
        if entry['NF'] in canceled_notes:
            apply_cancellation(entry)

    return note_entries


def write_csrf_column(writer, df):
    """Escreve CSRF como fórmula =SOMA(CSLL:COFINS) da própria linha, para o usuário
    conferir a conta direto na planilha. Notas que não fecham em 4,65% saem em vermelho.
    """
    red = writer.book.add_format({'font_color': 'red', 'bold': True, 'num_format': '#,##0.00'})
    normal = writer.book.add_format({'num_format': '#,##0.00'})
    worksheet = writer.sheets['NFS-e']
    csrf_col = DETAIL_COLUMNS.index('CSRF')
    csll_letter = xl_col_to_name(DETAIL_COLUMNS.index('CSLL'))
    cofins_letter = xl_col_to_name(DETAIL_COLUMNS.index('COFINS'))

    for offset, (_, row) in enumerate(df.iterrows()):
        excel_row = offset + 2  # +1 cabeçalho, +1 índice do Excel começa em 1
        formula = f'=SUM({csll_letter}{excel_row}:{cofins_letter}{excel_row})'
        fmt = red if not row['_csrf_ok'] else normal
        worksheet.write_formula(offset + 1, csrf_col, formula, fmt, row['CSRF'])


def highlight_canceled_rows(writer, df):
    """Pinta o fundo da linha inteira das notas com STATUS = Cancelada (estilo "Ruim" do
    Excel), para o usuário identificar de cara sem precisar ler a coluna STATUS.
    """
    if not (df['STATUS'] == 'Cancelada').any():
        return

    red_fill = writer.book.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    worksheet = writer.sheets['NFS-e']
    status_col = xl_col_to_name(DETAIL_COLUMNS.index('STATUS'))
    last_row = len(df)  # 0-indexado: linha 0 é o cabeçalho, última linha de dado é len(df)
    last_col = len(DETAIL_COLUMNS) - 1

    worksheet.conditional_format(1, 0, last_row, last_col, {
        'type': 'formula',
        'criteria': f'=${status_col}2="Cancelada"',
        'format': red_fill,
    })


def save_to_excel(data, output_path):
    df = pd.DataFrame(data).sort_values('Dt_Emissão')

    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        df[DETAIL_COLUMNS].to_excel(writer, sheet_name='NFS-e', index=False)
        write_csrf_column(writer, df)
        highlight_canceled_rows(writer, df)
        summary = build_summary(df, 'Mês Competência', VALUE_COLUMNS)
        summary.to_excel(writer, sheet_name='Total NFS-e', index=False)


def main():
    if len(sys.argv) < 2:
        logging.warning("Uso: python nfse_report.py <caminho_da_pasta>")
        return

    folder_path = sys.argv[1]
    data = process_folder(folder_path)

    if data:
        output_file = os.path.join(folder_path, 'notas_servico.xlsx')
        save_to_excel(data, output_file)
        logging.info(f'\nArquivo salvo em: {output_file}')
    else:
        logging.warning("Nenhuma nota de serviço encontrada.")


if __name__ == '__main__':
    main()
