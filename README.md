# ppt-doc

Gera gráficos (PNG) a partir do Excel e atualiza um PowerPoint com esses gráficos, mantendo o layout manual (posição/tamanho/crop) já ajustado no PPT.

Além disso, inclui um fluxo de **extração de dados de um XLSX para JSON** (baseado em um `specs.json`) para alimentar um prompt de LLM (ex.: Gemini) com apenas os valores relevantes.

## Arquivos principais

- `test-ppt.ipynb`: gera os PNGs dos gráficos a partir do `testing.xlsx`.
- `update_ppt.py`: atualiza o PowerPoint trocando as imagens **em lugar** (sem distorcer e sem mexer na geometria).
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

## Preencher título/subtítulo no PowerPoint (texto)

Para preencher o título e subtítulo do slide a partir do JSON retornado pelo endpoint `analyze_file`, use o `update_ppt.py` com `--text-json`.

### Como preparar o template PPT

No PowerPoint, no shape do título/subtítulo, coloque um placeholder de texto como:

- `{{slide1_title}}`
- `{{slide1_subtitle}}`

Recomendação: deixe cada placeholder como um token contínuo (evite quebrar o token com estilos diferentes), para preservar a formatação.

Alternativa: em vez de token no texto, você pode setar o **Alt Text** do shape como `slide1_title` / `slide1_subtitle`.

### Como rodar

Se você salvou a resposta do LLM em um arquivo JSON (por exemplo `llm_response.json`), rode:

```bash
python update_ppt.py --pptx /caminho/para/seu.pptx --images-dir /caminho/para/as/imagens --text-json llm_response.json
```

O `--text-json` aceita tanto o formato direto do modelo:

```json
{"titles": {"slide1_title": "..."}, "subtitles": {"slide1_subtitle": "..."}}
```

quanto o wrapper do endpoint:

```json
{"response": {"titles": {"slide1_title": "..."}, "subtitles": {"slide1_subtitle": "..."}}}
```

## Job fixo (só passa o XLSX)

Se essa aplicação vai ser “fixa” e você não quer passar um monte de argumentos no job, use o runner [run_fixed_job.py](run_fixed_job.py).

Você edita **uma vez só**:

- [config/job_config.json](config/job_config.json): template do PPT, saída, images dir
- [config/text_fields.json](config/text_fields.json): mapeamento `TOKEN -> célula A1` (e `default_sheet`)

### Job 100% automático (XLSX → API/LLM → PPT)

Se você quer que o job receba o XLSX e já dispare automaticamente:

1) upload do XLSX para a FastAPI do ambiente corporativo (`analyze_file`)
2) receba o JSON `{ "response": ... }`
3) preencha os textos no PPT

Configure em [config/job_config.json](config/job_config.json):

- `api_url`: URL completa do endpoint (ex.: `https://SEU_HOST/analyze_file`)
- `api_file_field`: nome do campo do upload (default: `file`)
- `api_headers`: headers (ex.: `{"Authorization": "Bearer ..."}`)
- `llm_response_json`: onde salvar a resposta (default: `llm_response.latest.json`)

### Campos vindos do JSON do LLM

Se alguns placeholders devem vir do JSON retornado pela LLM (em vez do Excel), configure:

- Em [config/job_config.json](config/job_config.json): `llm_response_json` (caminho do arquivo JSON do LLM/endpoint)
- Em [config/text_fields.json](config/text_fields.json): `llm_fields` (lista de chaves que devem ser preenchidas a partir desse JSON)

O runner carrega o JSON, extrai um mapping `chave -> texto` (suporta `{response:{...}}`), filtra só as chaves em `llm_fields` e aplica no PPT.

Regra de merge: se a mesma chave existir no XLSX e no LLM, o valor do LLM vence.

Depois, no job você passa apenas o XLSX:

```bash
cd /home/renato/ppt-doc
./.venv/bin/python run_fixed_job.py --xlsx /caminho/para/arquivo.xlsx
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
