# 📄 XML NFe to Excel

## 🇧🇷 Relatório de notas (NF-e 55 e NFC-e 65) — `billing_report.py`

O `billing_report.py` lê os XMLs de uma pasta e gera um Excel (`notas_fiscais.xlsx`) com **4 abas**:

| Aba | Conteúdo |
|-----|----------|
| `NF-e 55` | Uma linha por nota do modelo 55 |
| `NFC-e 65` | Uma linha por nota do modelo 65 |
| `Total NF-e 55` | Soma do `Valor` por `Tipo` (com total geral) do modelo 55 |
| `Total NFC-e 65` | Soma do `Valor` por `Tipo` (com total geral) do modelo 65 |

O modelo é detectado pela tag `<mod>` do XML. Notas NFC-e sem `<dest>` (consumidor não
identificado) são suportadas — os campos de destinatário ficam vazios.

### ▶️ Uso rápido no Windows (`run.bat`)

Basta **arrastar a pasta com os XMLs por cima do `run.bat`** (ou dar dois cliques, que ele usa
a própria pasta do `.bat`). Ele instala as dependências automaticamente e gera o
`notas_fiscais.xlsx` dentro da pasta processada. Requer apenas ter o **Python** instalado.

### Uso pela linha de comando

```bash
python billing_report.py <pasta/com/os/xmls>
```

---

## 🇧🇷 Relatório de notas de serviço (NFS-e) — `nfse_report.py`

O `nfse_report.py` lê os XMLs de **NFS-e do padrão nacional** (namespace
`http://www.sped.fazenda.gov.br/nfse`) de uma pasta e gera um Excel (`notas_servico.xlsx`)
com **2 abas**:

| Aba | Conteúdo |
|-----|----------|
| `NFS-e` | Uma linha por nota, ordenada por data de emissão |
| `Total NFS-e` | Soma das colunas de valor por mês de competência (com total geral) |

Colunas da aba de detalhe:

| Coluna | Origem no XML |
|--------|---------------|
| `NF` | `infNFSe/nNFSe` |
| `STATUS` | `Normal` ou `Cancelada` (ver seção de cancelamento abaixo) |
| `PRESTADOR` | `infNFSe/emit/xNome` (`infDPS/prest` só traz o CNPJ, sem nome) |
| `TOMADOR` | `infDPS/toma/xNome` (nota sem tomador identificado sai em branco) |
| `DESCRIÇÃO` | `infDPS/serv/cServ/xDescServ` (descrição livre digitada na nota) |
| `SERVIÇO` | `infNFSe/xTribMun` (enquadramento tributário municipal do serviço) |
| `VALOR` | `infDPS/valores/vServPrest/vServ` |
| `ISS` | `infNFSe/valores/vISSQN` |
| `RET` | `infDPS/valores/trib/tribMun/tpRetISSQN` (`2` = Retido, senão Não retido) |
| `INSS` | `infDPS/valores/trib/tribFed/vRetCP` |
| `CSLL` | `.../tribFed/vRetCSLL` |
| `PIS` | `.../tribFed/piscofins/vPis` |
| `COFINS` | `.../tribFed/piscofins/vCofins` |
| `CSRF` | fórmula no Excel: `=SOMA(CSLL:COFINS)` da própria linha |
| `IRR` | `.../tribFed/vRetIRRF` |
| `VALOR LIQUIDO` | calculado: `VALOR − INSS − CSRF − IRR` |

Notas sem retenção federal não trazem o grupo `<tribFed>` — nesse caso as colunas de
retenção saem `0.00` e o valor líquido é igual ao bruto. O leiaute atual não traz mais um flag
de retenção dentro de `<piscofins>` (`vPis`/`vCofins` vêm sempre crus); como os três tributos do
CSRF sempre vêm juntos quando há retenção, `PIS` e `COFINS` só entram com valor quando
`vRetCSLL` (o único campo com "Ret" no nome, ou seja, valor retido de verdade) vier preenchido.
Se `vRetCSLL = 0`, as três colunas saem zeradas.

#### ✅ Conferência do CSRF (4,65%)

`CSRF` é a contribuição social retida na fonte — a soma de **CSLL 1% + PIS 0,65% + COFINS 3%**,
que deve dar **4,65%** do faturamento da nota. Na planilha, a célula de `CSRF` é uma fórmula
(`=SOMA(CSLL:COFINS)` da própria linha), não um valor fixo — dá para conferir a conta direto
no Excel. O script confere esse total contra `VALOR × 4,65%`, com tolerância de um centavo:

- **Bate** → a nota sai normal, com `CSLL`, `PIS` e `COFINS` separados em suas colunas.
- **Não bate** → a nota **é importada assim mesmo**, mas `CSLL`, `PIS` e `COFINS` saem `0.00`
  e a célula de `CSRF` aparece em **vermelho** na planilha, porque aquele valor não pode ser
  separado nas três alíquotas. O script também emite um `WARNING` no log com o valor
  encontrado e o esperado. Cabe conferência manual da nota.
- **CSRF igual a zero** (nota sem retenção de contribuições) **não é divergência** e não sai
  em vermelho.

A alíquota e a tolerância ficam em `CSRF_RATE` e `CSRF_TOLERANCE`, no topo do `nfse_report.py`.
Na aba de totais, `CSRF` é a coluna confiável para o total retido: nas notas divergentes ela é
a única preenchida, então `CSLL + PIS + COFINS` pode não fechar com o `CSRF` somado.

#### 🚫 Notas canceladas

A Receita gera um XML separado, de **evento**, quando uma nota é cancelada (raiz `<evento>`,
não `<NFSe>`). O script identifica esse evento pelo texto de `xDesc` (contém "cancelamento") e
liga ele à nota pelo número (`infEvento/nDFSe` do evento = `nNFSe` da nota) — não precisa de
nenhum nome de arquivo especial, nem que os dois XMLs estejam na mesma ordem de leitura.

Quando encontra o evento de cancelamento de uma nota:
- `STATUS` vira `Cancelada` e a **linha inteira sai destacada em vermelho** na planilha.
- Todas as colunas de valor (`VALOR`, `ISS`, `INSS`, `CSLL`, `PIS`, `COFINS`, `CSRF`, `IRR`,
  `VALOR LIQUIDO`) saem zeradas — a nota cancelada não entra na aba `Total NFS-e`.
- As colunas descritivas (`PRESTADOR`, `TOMADOR`, `DESCRIÇÃO`, `SERVIÇO`, `RET`) continuam
  preenchidas, só para identificar qual nota foi cancelada.

O próprio arquivo de evento não vira uma linha na planilha — ele só serve de gatilho para
marcar a nota correspondente.

### ▶️ Uso rápido no Windows (`run_servicos.bat`)

Igual ao `run.bat`: **arraste a pasta com os XMLs por cima do `run_servicos.bat`**.

### Uso pela linha de comando

```bash
python nfse_report.py <pasta/com/os/xmls>
```

### 📂 Pastas com XMLs misturados

Os dois relatórios detectam o leiaute pelo namespace da raiz e **ignoram o que não é do seu tipo**.
Dá para apontar `run.bat` e `run_servicos.bat` para a mesma pasta: um gera o `notas_fiscais.xlsx`
só com as NF-e/NFC-e, o outro o `notas_servico.xlsx` só com as NFS-e.

---



A Python project that reads **XML files from Electronic Invoices (NFe)**, extracts product information, and generates an organized **Excel (.xlsx)** table.

## 📊 Features

- 📂 Reads all `.xml` files from a directory.
- 📄 Extracts the following information from each NFe:
    - Emission Date
    - Product
    - Quantity
    - Unit Value
    - Total Value
- 📊 Creates a consolidated Excel file with a separate sheet for each NFe.
- ✅ Displays the **Emission Date** only in the first row of each sheet.

## 📋 Requirements

- Python 3.10 or higher
- Required libraries:

```bash
pip install pandas openpyxl
```

## 📁 Project Structure

```
xml_to_excel/
├── main.py          # Main script
├── requirements.txt # Dependencies
└── README.md        # Project documentation
```

## 🚀 How to Run

1. **Clone the repository**

```bash
git clone https://github.com/your-username/xml-to-excel.git
cd xml-to-excel
```

2. **Install dependencies**

```bash
pip install -r requirements.txt
```

3. **Run the script**

```bash
python main.py <your/xml/files/directory>
```

4. **Output:**

- The `output.xlsx` file will be generated in the specified directory, with a sheet for each processed XML file.

## 📊 Example Output

| Date                | Product      | Quantity | Unit Value | Total Value |
|---------------------|--------------|----------|------------|-------------|
| 2024-02-07T14:30:00 | Product A    | 2        | 50.00      | 100.00      |
|                     | Product B    | 1        | 30.00      | 30.00       |
|                     | Product C    | 5        | 10.00      | 50.00       |

## 🛠️ Customization

If you need to adjust the extracted fields, edit the `find` calls inside the loop that processes products in the XML.

## 📌 Notes

- Ensure the XML files follow the NFe standard and contain the correct namespaces.
- Each sheet name in the Excel file is based on the XML file name (limited to 31 characters).

## 📄 License

This project is licensed under the [MIT License](LICENSE).

