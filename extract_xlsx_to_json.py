from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import List

from utils.xlsx_extract import ExtractSpec, extract_xlsx_to_dict, parse_specs_args, parse_specs_json


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Extrai ranges de um XLSX e gera JSON.\n\n"
            "Você define uma lista de specs (id + ranges de labels/values).\n"
            "Exemplo: ROE_9M:L3:M3:L20:M20"
        )
    )
    parser.add_argument("--xlsx", required=True, help="Caminho do arquivo .xlsx")
    parser.add_argument(
        "--sheet",
        default=None,
        help="Aba default (se você não passar sheet em cada spec).",
    )
    parser.add_argument(
        "--spec",
        action="append",
        default=[],
        help=(
            "Um spec no formato ID:LABELS_RANGE:VALUES_RANGE ou ID:SHEET:LABELS_RANGE:VALUES_RANGE. "
            "Pode repetir --spec várias vezes."
        ),
    )
    parser.add_argument(
        "--specs-json",
        default=None,
        help=(
            "Arquivo JSON com uma lista de specs. Cada item: "
            "{\"id\":\"ROE_9M\",\"sheet\":\"DRE Saida\",\"labels_range\":\"L3:M3\",\"values_range\":\"L20:M20\"}."
        ),
    )
    parser.add_argument(
        "--out",
        default=None,
        help="Arquivo de saída .json (se omitido, imprime no stdout).",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Falha se algum valor em Values não for numérico.",
    )
    parser.add_argument(
        "--include-meta",
        action="store_true",
        help="Inclui Sheet e Ranges no output.",
    )
    parser.add_argument(
        "--lowercase-fields",
        action="store_true",
        help="Usa labels/values/sheet/ranges em vez de Labels/Values/Sheet/Ranges.",
    )

    args = parser.parse_args()

    specs: List[ExtractSpec] = []
    if args.specs_json:
        specs.extend(parse_specs_json(Path(args.specs_json)))
    if args.spec:
        specs.extend(parse_specs_args(args.spec, args.sheet))

    if not specs:
        raise SystemExit("Você precisa informar ao menos um --spec ou --specs-json")

    payload = extract_xlsx_to_dict(
        args.xlsx,
        specs,
        default_sheet=args.sheet,
        strict_numbers=bool(args.strict),
        include_meta=bool(args.include_meta),
        lowercase_fields=bool(args.lowercase_fields),
    )

    rendered = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.out:
        Path(args.out).write_text(rendered + "\n", encoding="utf-8")
    else:
        print(rendered)


if __name__ == "__main__":
    main()
