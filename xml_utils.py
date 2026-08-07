"""Funções compartilhadas entre os relatórios de NF-e e de NFS-e."""

import pandas as pd

NFE_NS = 'http://www.portalfiscal.inf.br/nfe'
NFSE_NS = 'http://www.sped.fazenda.gov.br/nfse'


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
