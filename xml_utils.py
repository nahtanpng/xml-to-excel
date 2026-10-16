"""Funções compartilhadas entre os relatórios de NF-e e de NFS-e."""

import pandas as pd

NFE_NS = 'http://www.portalfiscal.inf.br/nfe'
NFSE_NS = 'http://www.sped.fazenda.gov.br/nfse'

ERROR_COLUMNS = ['Arquivo', 'Nota', 'Série', 'Erro']


def detect_layout(root):
    """Identifica o leiaute do XML pela raiz.

    Retorna 'nfe' (NF-e/NFC-e), 'nfse' (NFS-e padrão nacional), 'evento' (evento de NFS-e,
    ex.: cancelamento — mesmo namespace da NFS-e, raiz <evento> em vez de <NFSe>) ou None.
    """
    tag = root.tag
    if tag.startswith(f'{{{NFE_NS}}}'):
        return 'nfe'
    if tag == f'{{{NFSE_NS}}}evento':
        return 'evento'
    if tag.startswith(f'{{{NFSE_NS}}}'):
        return 'nfse'
    return None


def parse_text(element, path, namespaces):
    if element is None:
        return None
    found = element.find(path, namespaces)
    return found.text if found is not None else None


def to_float(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def build_summary(df, group_col, value_cols):
    """Soma value_cols por group_col e acrescenta uma linha de total geral."""
    if df.empty:
        return pd.DataFrame(columns=[group_col, *value_cols])

    summary = df.groupby(group_col)[value_cols].sum().reset_index()
    total_row = pd.DataFrame([{group_col: 'Total Geral', **df[value_cols].sum().to_dict()}])
    return pd.concat([summary, total_row], ignore_index=True)


def build_error(filename, note, series, exc):
    """Linha da aba "Erros". Nota/série vêm vazias quando o XML nem chegou a ser lido."""
    return {
        'Arquivo': filename,
        'Nota': note,
        'Série': series,
        # Muitas exceções do ElementTree e de atributo têm mensagem vazia ou crua; o nome
        # do tipo ajuda a entender o que aconteceu.
        'Erro': f'{type(exc).__name__}: {exc}',
    }


def write_errors_sheet(writer, errors):
    """Grava a aba "Erros" (só quando algum arquivo falhou)."""
    if not errors:
        return

    pd.DataFrame(errors, columns=ERROR_COLUMNS).to_excel(writer, sheet_name='Erros', index=False)
    worksheet = writer.sheets['Erros']
    worksheet.set_column(ERROR_COLUMNS.index('Arquivo'), ERROR_COLUMNS.index('Arquivo'), 50)
    worksheet.set_column(ERROR_COLUMNS.index('Erro'), ERROR_COLUMNS.index('Erro'), 80)
