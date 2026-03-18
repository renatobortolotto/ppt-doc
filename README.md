# ppt-doc

Gera graficos (PNG) a partir de um Excel e atualiza um PowerPoint mantendo o layout manual do template.

Na arquitetura atual, este projeto nao chama mais a LLM diretamente. Ele recebe:

- um `xlsx`
- um JSON ja gerado pela LLM

e devolve o `.pptx` final.

## Arquivos principais

- `main-framework.py`: endpoint corporativo que recebe `xlsx + json` e retorna o PPT em base64
- `presentation_builder.py`: servico reutilizavel para montar o PPT por path ou por bytes
- `run_fixed_job.py`: runner local/fixo para gerar o PPT a partir de `xlsx + json`
- `update_ppt.py`: atualiza o template PPT trocando imagens e preenchendo textos
- `config/job_config.json`: template, saida, diretorio de imagens e JSON padrao da LLM
- `config/text_fields.json`: mapeamento `TOKEN -> celula A1` e chaves que devem vir do JSON da LLM
- `utils/`: geradores de graficos e extracao de campos do Excel

Nos campos de texto vindos do Excel, voce tambem pode usar `div` para dividir o valor antes de converter para string, `round` para definir quantas casas decimais o texto deve ter, `is_porc` para preservar a exibicao percentual do Excel (por exemplo `9,9%` em vez de `0.099`) e `is_pp` para forcar a leitura numerica pura de celulas com formato customizado em `p.p.`, deixando `round` e `VAR_` funcionarem em cima do valor bruto. Exemplo:

```json
{
  "fields": {
    "ROE_EXIBICAO": {"sheet": "DRE Saida", "cell": "K20", "is_porc": true},
    "VARIACAO_PP_VALOR": {"sheet": "DRE Saida", "cell": "K21", "is_pp": true, "round": 1},
    "CARTEIRA_EM_MILHARES": {"sheet": "Premissas", "cell": "B3", "div": 1000, "round": 1}
  }
}
```

## Fluxo atual

1) outro servico recebe o `xlsx` e devolve um JSON no formato de `llm_response.latest.json`
2) `ppt-doc` recebe o `xlsx` e esse JSON
3) gera/atualiza os PNGs dos graficos
4) mistura campos vindos do Excel com campos vindos da LLM
5) atualiza o template PPT e devolve o `.pptx`

## Endpoint corporativo

O arquivo `main-framework.py` expoe a rota `compose_presentation`.

Contrato:

- `xlsx_file`: arquivo Excel
- `llm_response_file`: arquivo JSON com a resposta da LLM

Helpers disponiveis:

- `compose_presentation(xlsx_file, llm_response_file)`
- `compose_presentation_files(xlsx_file, llm_response_file)`
- `compose_presentation_from_inputs(xlsx_bytes, llm_response_bytes)`

## Runner local

Se voce quer rodar localmente sem depender do runtime corporativo, use `run_fixed_job.py`.

Exemplo usando um JSON explicito:

```bash
cd /home/renato/projetos/double-projects/ppt-doc
MPLCONFIGDIR=/tmp/matplotlib ./.venv/bin/python run_fixed_job.py --xlsx testing.xlsx --llm-json llm_response.latest.json
```

No Windows/PowerShell, use o Python da virtualenv assim:

```powershell
cd C:\caminho\para\ppt-doc
.\.venv\Scripts\python.exe run_fixed_job.py --xlsx testing.xlsx --llm-json llm_response.latest.json
```

Se a virtualenv ja estiver ativada, `python run_fixed_job.py ...` tambem funciona.

Exemplo usando o JSON configurado em `config/job_config.json`:

```bash
cd /home/renato/projetos/double-projects/ppt-doc
MPLCONFIGDIR=/tmp/matplotlib ./.venv/bin/python run_fixed_job.py --xlsx /caminho/para/arquivo.xlsx
```

Para testar apenas os graficos de alguns slides, use `--only-slides`. Ele aceita lista e intervalo, como `3,4,7,8` ou `1-8`:

```bash
cd /home/renato/projetos/double-projects/ppt-doc
MPLCONFIGDIR=/tmp/matplotlib ./.venv/bin/python run_fixed_job.py --xlsx testing.xlsx --only-slides 1-8
```

Nesse modo, a geracao de imagens fica isolada em uma pasta temporaria para evitar reaproveitar PNG antigo de outros slides. Os textos do PPT continuam sendo atualizados normalmente.

Se aparecer erro como `BadZipFile: File is not a zip file`, o arquivo informado em `--xlsx` nao e um `.xlsx` valido. Isso costuma acontecer quando o arquivo:

- foi renomeado para `.xlsx`, mas originalmente era `.xls`, `.csv` ou `.html`
- foi baixado de algum portal e salvo como pagina web
- esta corrompido
- e um arquivo temporario/incompleto

Com a config atual:

- `pptx_template`: `teste-design.gerado.updated.pptx`
- `pptx_output`: `main_testing.pptx`
- `images_dir`: `.`
- `llm_response_json`: `llm_response.latest.json`

## Texto no PowerPoint

Para preencher titulo e subtitulo a partir do JSON, use placeholders como:

- `{{slide1_title}}`
- `{{slide1_subtitle}}`

Alternativamente, voce pode setar o Alt Text do shape como `slide1_title` ou `slide1_subtitle`.

O `update_ppt.py` aceita tanto:

```json
{"titles": {"slide1_title": "..."}, "subtitles": {"slide1_subtitle": "..."}}
```

quanto:

```json
{"response": {"titles": {"slide1_title": "..."}, "subtitles": {"slide1_subtitle": "..."}}}
```

Exemplo:

```bash
python update_ppt.py --pptx /caminho/para/seu.pptx --images-dir /caminho/para/as/imagens --text-json llm_response.latest.json
```

## Testes

```bash
cd /home/renato/projetos/double-projects/ppt-doc
MPLCONFIGDIR=/tmp/matplotlib ./.venv/bin/python -m unittest discover -s tests -v
```

## Observacoes

- O `main-framework.py` e o contrato corporativo esperam sempre dois arquivos: `xlsx + json`
- O `run_fixed_job.py` usa `--llm-json` quando voce quer apontar um JSON especifico, ou `llm_response_json` da config quando o arquivo padrao ja existe
- O notebook `test-ppt.ipynb` pode continuar sendo usado para exploracao manual, mas o fluxo principal ja gera os graficos pelo builder
