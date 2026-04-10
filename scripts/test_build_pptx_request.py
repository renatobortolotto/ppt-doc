from __future__ import annotations

import argparse
import base64
import json
from pathlib import Path

import requests


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Testa o endpoint /build-pptx enviando XLSX + JSON da LLM via requests."
    )
    parser.add_argument(
        "--url",
        default="http://localhost:8080/api/build-pptx",
        help="URL completa do endpoint build-pptx.",
    )
    parser.add_argument(
        "--xlsx",
        required=True,
        help="Caminho para o arquivo XLSX de entrada.",
    )
    parser.add_argument(
        "--llm-json",
        required=True,
        help="Caminho para o JSON com a resposta da LLM.",
    )
    parser.add_argument(
        "--token",
        help="JWT opcional para enviar no header Authorization.",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=120,
        help="Timeout da requisicao em segundos.",
    )
    parser.add_argument(
        "--save-response-json",
        help="Caminho opcional para salvar o JSON bruto da resposta.",
    )
    parser.add_argument(
        "--save-pptx",
        help="Caminho opcional para salvar o PPTX retornado pelo endpoint.",
    )
    return parser.parse_args()


def default_output_path(source_path: Path, suffix: str) -> Path:
    return source_path.with_name(f"{source_path.stem}{suffix}")


def main() -> int:
    args = parse_args()
    xlsx_path = Path(args.xlsx).expanduser().resolve()
    llm_json_path = Path(args.llm_json).expanduser().resolve()

    if not xlsx_path.is_file():
        raise FileNotFoundError(f"XLSX nao encontrado: {xlsx_path}")
    if not llm_json_path.is_file():
        raise FileNotFoundError(f"JSON da LLM nao encontrado: {llm_json_path}")

    headers: dict[str, str] = {}
    if args.token:
        headers["Authorization"] = f"Bearer {args.token}"

    with xlsx_path.open("rb") as xlsx_file, llm_json_path.open("rb") as llm_json_file:
        files = {
            "xlsx_file": (
                xlsx_path.name,
                xlsx_file,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ),
            "llm_response_file": (
                llm_json_path.name,
                llm_json_file,
                "application/json",
            ),
        }
        response = requests.post(
            args.url,
            files=files,
            headers=headers,
            timeout=args.timeout,
        )

    print(f"HTTP {response.status_code}")

    try:
        payload = response.json()
    except ValueError:
        print(response.text)
        response.raise_for_status()
        return 0

    response_json_path = (
        Path(args.save_response_json).expanduser().resolve()
        if args.save_response_json
        else default_output_path(xlsx_path, ".build-pptx.response.json")
    )
    response_json_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"Resposta salva em: {response_json_path}")

    if not response.ok:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        response.raise_for_status()

    pptx_base64 = payload.get("pptxBase64")
    if pptx_base64:
        pptx_output_path = (
            Path(args.save_pptx).expanduser().resolve()
            if args.save_pptx
            else default_output_path(xlsx_path, ".build-pptx.result.pptx")
        )
        pptx_output_path.write_bytes(base64.b64decode(pptx_base64))
        print(f"PPTX salvo em: {pptx_output_path}")

    summary = payload.get("summary")
    if summary:
        print("Resumo:")
        print(json.dumps(summary, ensure_ascii=False, indent=2))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
