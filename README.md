# ppt-doc

Gera gráficos (PNG) a partir do Excel e atualiza um PowerPoint com esses gráficos, mantendo o layout manual (posição/tamanho/crop) já ajustado no PPT.

## Arquivos principais

- `test-ppt.ipynb`: gera os PNGs dos gráficos a partir do `testing.xlsx`.
- `update_ppt.py`: atualiza o PowerPoint trocando as imagens **em lugar** (sem distorcer e sem mexer na geometria).

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
