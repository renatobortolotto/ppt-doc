from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any, Dict, Mapping, Sequence

from update_ppt import _flatten_text_payload, update_presentation
from utils.slide11_charts import generate_slide11_charts
from utils.slide12_charts import generate_slide12_charts
from utils.slide13_charts import generate_slide13_charts
from utils.slide14_charts import generate_slide14_charts
from utils.slide15_charts import generate_slide15_charts
from utils.slide18_charts import generate_slide18_charts
from utils.slide1_charts import generate_slide1_charts
from utils.slide20_charts import generate_slide20_charts
from utils.slide2_charts import generate_slide2_charts
from utils.slide3_charts import generate_slide3_charts
from utils.slide8_charts import generate_slide8_charts
from utils.slide9_charts import generate_slide9_charts
from utils.slide_pizza_charts import generate_pizza_charts
from utils.xlsx_extract import _load_workbook as _load_validated_workbook
from utils.xlsx_text_fields import extract_xlsx_to_text_mapping, parse_text_fields_json


@dataclass(frozen=True)
class BuildPresentationResult:
    output_path: Path
    replaced_pictures: int
    replaced_placeholders: int
    replaced_text: int
    generated_chart_count: int
    applied_text_keys: tuple[str, ...]


def resolve_path(repo_root: Path, path_value: str) -> Path:
    path = Path(path_value).expanduser()
    if path.is_absolute():
        return path
    return (repo_root / path).resolve()


def load_job_config(repo_root: Path) -> Dict[str, Any]:
    cfg_path = repo_root / "config" / "job_config.json"
    if not cfg_path.exists():
        raise FileNotFoundError(
            f"Config nao encontrada: {cfg_path}. Edite uma vez e rode novamente."
        )
    raw = json.loads(cfg_path.read_text(encoding="utf-8"))
    if not isinstance(raw, dict):
        raise ValueError("job_config.json deve ser um objeto")
    return raw


def load_llm_payload_from_path(repo_root: Path, cfg: Mapping[str, Any]) -> object | None:
    llm_path = cfg.get("llm_response_json")
    if not llm_path:
        return None

    path = resolve_path(repo_root, str(llm_path))
    if not path.exists():
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def load_llm_mapping_from_payload(payload: object | None) -> Dict[str, str]:
    if payload is None:
        return {}
    if isinstance(payload, dict) and "response" in payload and isinstance(payload["response"], dict):
        payload = payload["response"]
    return _flatten_text_payload(payload)


def _filter_llm_mapping(text_fields_config: Path, llm_mapping: Dict[str, str]) -> Dict[str, str]:
    raw_text_cfg = json.loads(text_fields_config.read_text(encoding="utf-8"))
    llm_fields: list[str] = []
    if isinstance(raw_text_cfg, dict):
        raw_llm_fields = raw_text_cfg.get("llm_fields") or raw_text_cfg.get("from_llm")
        if isinstance(raw_llm_fields, list):
            llm_fields = [str(value) for value in raw_llm_fields]

    if not llm_fields:
        return llm_mapping

    allowed = set(llm_fields)
    return {key: value for key, value in llm_mapping.items() if key in allowed}


def build_text_mapping(
    *,
    repo_root: Path,
    cfg: Mapping[str, Any],
    xlsx_path: Path,
    llm_payload: object | None,
) -> Dict[str, str]:
    text_fields_config = resolve_path(
        repo_root,
        str(cfg.get("text_fields_config", "config/text_fields.json")),
    )
    default_sheet, specs = parse_text_fields_json(text_fields_config)
    text_mapping = extract_xlsx_to_text_mapping(
        xlsx_path,
        specs,
        default_sheet=default_sheet,
    )

    llm_mapping = load_llm_mapping_from_payload(llm_payload)
    llm_mapping = _filter_llm_mapping(text_fields_config, llm_mapping)
    text_mapping.update(llm_mapping)
    return text_mapping


def _chart_generators() -> Sequence:
    return (
        generate_slide1_charts,
        generate_slide2_charts,
        generate_slide3_charts,
        generate_pizza_charts,
        generate_slide8_charts,
        generate_slide9_charts,
        generate_slide11_charts,
        generate_slide12_charts,
        generate_slide13_charts,
        generate_slide14_charts,
        generate_slide15_charts,
        generate_slide18_charts,
        generate_slide20_charts,
    )


def generate_chart_assets(*, xlsx_path: Path, images_dir: Path) -> int:
    generated = 0
    for generator in _chart_generators():
        generated += len(generator(xlsx_path=xlsx_path, output_dir=images_dir))
    return generated


def _validate_xlsx_path(xlsx_path: Path) -> None:
    try:
        wb = _load_validated_workbook(filename=xlsx_path, data_only=True)
    except ValueError as exc:
        raise ValueError(
            f"Arquivo Excel invalido: {xlsx_path}. "
            "Verifique se o arquivo e um .xlsx real do Excel, nao um .xls/.csv renomeado, HTML baixado da web, ou arquivo corrompido."
        ) from exc

    close = getattr(wb, "close", None)
    if callable(close):
        close()


def build_presentation(
    *,
    repo_root: Path,
    cfg: Mapping[str, Any],
    xlsx_path: Path,
    llm_payload: object | None = None,
    output_path: Path | None = None,
    images_dir: Path | None = None,
    skip_charts: bool = False,
) -> BuildPresentationResult:
    if not xlsx_path.exists():
        raise FileNotFoundError(f"XLSX nao encontrado: {xlsx_path}")

    _validate_xlsx_path(xlsx_path)

    effective_output_path = output_path or resolve_path(repo_root, str(cfg.get("pptx_output")))
    effective_images_dir = images_dir or resolve_path(repo_root, str(cfg.get("images_dir", ".")))
    effective_images_dir.mkdir(parents=True, exist_ok=True)

    pptx_template = resolve_path(repo_root, str(cfg.get("pptx_template")))
    allow_placeholder_text = bool(cfg.get("allow_placeholder_text", False))
    effective_llm_payload = llm_payload if llm_payload is not None else load_llm_payload_from_path(repo_root, cfg)

    generated_chart_count = 0
    if not skip_charts:
        generated_chart_count = generate_chart_assets(
            xlsx_path=xlsx_path,
            images_dir=effective_images_dir,
        )

    text_mapping = build_text_mapping(
        repo_root=repo_root,
        cfg=cfg,
        xlsx_path=xlsx_path,
        llm_payload=effective_llm_payload,
    )

    (
        replaced_pictures,
        replaced_placeholders,
        replaced_text,
        _replaced_files,
        _missing_files,
        applied_text_keys,
    ) = update_presentation(
        pptx_path=pptx_template,
        output_path=effective_output_path,
        images_dir=effective_images_dir,
        allow_placeholder_text=allow_placeholder_text,
        text_json=None,
        text_payload=text_mapping,
    )

    return BuildPresentationResult(
        output_path=effective_output_path,
        replaced_pictures=replaced_pictures,
        replaced_placeholders=replaced_placeholders,
        replaced_text=replaced_text,
        generated_chart_count=generated_chart_count,
        applied_text_keys=tuple(applied_text_keys),
    )


def build_presentation_from_bytes(
    *,
    repo_root: Path,
    cfg: Mapping[str, Any],
    xlsx_bytes: bytes,
    llm_payload: object | None = None,
    skip_charts: bool = False,
) -> tuple[bytes, BuildPresentationResult]:
    if not xlsx_bytes:
        raise ValueError("XLSX vazio")

    with TemporaryDirectory(prefix="ppt-doc-build-") as tmp_dir:
        tmp_root = Path(tmp_dir)
        xlsx_path = tmp_root / "input.xlsx"
        xlsx_path.write_bytes(xlsx_bytes)

        images_dir = tmp_root / "images"
        output_filename = str(cfg.get("api_output_filename") or cfg.get("pptx_output") or "presentation.updated.pptx")
        output_path = tmp_root / Path(output_filename).name

        result = build_presentation(
            repo_root=repo_root,
            cfg=cfg,
            xlsx_path=xlsx_path,
            llm_payload=llm_payload,
            output_path=output_path,
            images_dir=images_dir,
            skip_charts=skip_charts,
        )
        logical_result = BuildPresentationResult(
            output_path=Path(Path(output_filename).name),
            replaced_pictures=result.replaced_pictures,
            replaced_placeholders=result.replaced_placeholders,
            replaced_text=result.replaced_text,
            generated_chart_count=result.generated_chart_count,
            applied_text_keys=result.applied_text_keys,
        )
        return output_path.read_bytes(), logical_result
