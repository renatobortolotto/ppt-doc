# ppt-doc

Gera gráficos (PNG) a partir do Excel e atualiza um PowerPoint com esses gráficos, mantendo o layout manual (posição/tamanho/crop) já ajustado no PPT.

Além disso, inclui um fluxo de **extração de dados de um XLSX para JSON** (baseado em um `specs.json`) para alimentar um prompt de LLM (ex.: Gemini) com apenas os valores relevantes.

## Arquivos principais

- `test-ppt.ipynb`: gera os PNGs dos gráficos a partir do `testing.xlsx`.
- `update_ppt.py`: atualiza o PowerPoint trocando as imagens **em lugar** (sem distorcer e sem mexer na geometria).
- `extract_xlsx_to_json.py`: CLI para extrair ranges do Excel e imprimir/salvar JSON.
- `specs.example.json`: exemplo de specs para extração.
- `main-framework.py`: exemplo de “chassi” de API corporativa (recebe XLSX, extrai JSON, chama Gemini e retorna JSON).
- `utils/`: utilitários reutilizáveis (extração XLSX e parsing robusto de JSON retornado pelo LLM).
- `tests/`: testes unitários (unittest) do que é usado no fluxo do `main-framework.py`.

## Fluxo completo (XLSX → specs.json → JSON → Gemini → JSON)

O fluxo que você descreveu (e que está modelado no `main-framework.py`) é:

1) **API recebe um arquivo `.xlsx`**
- No ambiente corporativo, o framework entrega `file.content` como `bytes`.

2) **Carrega um `config/specs.json`**
- O `config/specs.json` define quais ranges (A1) ler do Excel.
- Cada spec tem:
  - `id`: identificador (use camelCase se quiser bater com o seu schema)
  - `sheet`: nome da aba (opcional se você usar `DEFAULT_SHEET`)
  - `labels_range`: range A1 dos labels (ex.: `L3:M3`)
  - `values_range`: range A1 dos valores (ex.: `L18:M18`)

Exemplo (formato recomendado) em `config/specs.json`:

```json
[
  {
    "id": "lucroTrimestre",
    "sheet": "DRE Saida",
    "labels_range": "C3:K3",
    "values_range": "C18:K18"
  },
  {
    "id": "lucro9M",
    "sheet": "DRE Saida",
    "labels_range": "L3:M3",
    "values_range": "L18:M18"
  }
]
```

3) **Extrai o XLSX → dict JSON-friendly**
- Implementado em `utils/xlsx_extract.py`.
- O `main-framework.py` usa `extract_xlsx_bytes_to_dict(...)` (entrada em bytes).
- Por padrão, a extração pode incluir metadados (`sheet` e `ranges`) e pode normalizar as chaves para minúsculas.

Formato típico produzido (com `lowercase_fields=True`):

```json
{
  "lucroTrimestre": {
    "labels": ["3T23", "4T23"],
    "values": [285.0, 302.0],
    "sheet": "DRE Saida",
    "ranges": {"labels": "C3:D3", "values": "C18:D18"}
  },
  "lucro9M": {
    "labels": ["9M24", "9M25"],
    "values": [1180.0, 1400.0],
    "sheet": "DRE Saida",
    "ranges": {"labels": "L3:M3", "values": "L18:M18"}
  }
}
```

4) **Monta o prompt final**
- O JSON extraído é serializado e concatenado nas instruções do prompt.
- Resultado: o modelo recebe apenas dados relevantes + regras de output.

5) **Chama o Gemini e faz parse do JSON de saída**
- O modelo é instruído a retornar apenas JSON.
- Na prática, pode vir com fences (` ```json `) ou lixo em volta; por isso `utils/json_utils.py` tem `coerce_json()`.

6) **A API retorna JSON**
- Se o parse falhar, retorna um JSON de erro com `model_response` para debug.

## Rodando a extração (CLI)

Extrair via `config/specs.json`:

```bash
cd /home/renato/ppt-doc
./.venv/bin/python extract_xlsx_to_json.py --xlsx testing.xlsx --specs-json config/specs.json --include-meta --lowercase-fields
```

Salvar em arquivo:

```bash
./.venv/bin/python extract_xlsx_to_json.py --xlsx testing.xlsx --specs-json config/specs.json --out saida.json --include-meta --lowercase-fields
```

## Rodando o exemplo de API (main-framework)

O `main-framework.py` é um exemplo para o seu chassi corporativo. Ele usa:

- `SPECS_JSON_PATH`: caminho do specs (default: `config/specs.json`)
- `DEFAULT_SHEET`: aba default (opcional)
- `project_id` e `location`: usados para instanciar `Client(vertexai=True, ...)`

Obs.: as dependências `genai_framework` e `google.genai` não estão disponíveis neste workspace; no seu ambiente corporativo elas devem existir.

## Testes unitários

Os testes cobrem apenas o que é usado no fluxo do `main-framework.py` (os utilitários em `utils/`).

Rodar:

```bash
cd /home/renato/ppt-doc
./.venv/bin/python -m unittest discover -s tests -v
```

## Como funciona o `update_ppt.py`

O script faz duas formas de substituição:

1) **Substitui imagens já inseridas no PPT**
- Ele procura shapes do tipo PICTURE.
- Para cada imagem, lê o **Alt Text** (campo `descr` interno do PPT).
- Se existir um arquivo no diretório de imagens com o mesmo nome do Alt Text (ex.: `01_lucro_trimestres.png`), ele troca apenas o “embed” da imagem.
- Resultado: mantém **posição, tamanho e crop** que você configurou manualmente no PowerPoint.

2) **(Opcional) Substitui placeholders de texto por imagens**
- Se você passar `--allow-placeholder-text`, o script também substitui TextBox cujo texto (trim) seja exatamente um nome de arquivo existente no diretório de imagens (ex.: `02_lucro_9m.png`).

### PPT de entrada e saída

Você escolhe explicitamente qual PPT atualizar via `--pptx`.

Por padrão, o script gera um arquivo ao lado, com sufixo `.updated.pptx`.
Se quiser sobrescrever o arquivo de entrada, use `--in-place`.

## Como rodar (fluxo recomendado)

1) Gere/atualize os PNGs no notebook
- Abra `test-ppt.ipynb`
- Ajuste `EXCEL_FILE` se necessário
- Rode as células que geram:
  - `01_lucro_trimestres.png`
  - `02_lucro_9m.png`
  - `03_roe_trimestres.png`
  - `04_roe_9m.png`

2) Atualize o PPT

```bash
cd /home/renato/ppt-doc
python update_ppt.py --pptx /caminho/para/seu.pptx --images-dir /caminho/para/as/imagens
```

Exemplo sobrescrevendo o próprio arquivo:

```bash
python update_ppt.py --pptx /caminho/para/seu.pptx --images-dir /home/renato/ppt-doc --in-place
```

## Logs e avisos (warnings)

Ao rodar `python update_ppt.py`, você vai ver:

- `INFO` com caminhos e quantas substituições foram feitas.
- `VERIF` com a lista de Alt Texts das imagens encontradas no PPT gerado.
- Um **warning** se:
  - existir Alt Text (ou placeholder de texto habilitado) que pareça arquivo de imagem (`.png/.jpg/.jpeg`) mas o arquivo não estiver no diretório de imagens.

Isso serve para “testar e garantir” que o PPT foi atualizado por completo.
