from __future__ import annotations

import json
import logging
import shutil
from dataclasses import dataclass
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any, Callable, Dict, Mapping, Sequence

from update_ppt import _flatten_text_payload, update_presentation
from utils.slide11_charts import generate_slide11_charts
from utils.slide12_charts import generate_slide12_charts
from utils.slide13_charts import generate_slide13_charts
from utils.slide14_charts import generate_slide14_charts
from utils.slide15_charts import generate_slide15_charts
from utils.slide16_charts import generate_slide16_charts
from utils.slide18_charts import generate_slide18_charts
from utils.slide10_charts import generate_slide10_charts
from utils.slide20_charts import generate_slide20_charts
from utils.slide7_charts import generate_slide7_charts
from utils.slide3_charts import generate_slide3_charts
from utils.slide8_charts import generate_slide8_charts
from utils.slide9_charts import generate_slide9_charts
from utils.slide4_charts import generate_slide4_charts
from utils.xlsx_extract import _load_workbook as _load_validated_workbook
from utils.xlsx_text_fields import (
    TextFieldExtractionResult,
    TextFieldFailure,
    extract_xlsx_to_text_mapping_tolerant,
    parse_text_fields_json,
)


ChartGeneratorFn = Callable[..., list[Path]]


@dataclass(frozen=True)
class ChartGeneratorSpec:
    key: str
    label: str
    generator: ChartGeneratorFn
    output_files: tuple[str, ...]


@dataclass(frozen=True)
class ChartGenerationFailure:
    generator_key: str
    label: str
    output_files: tuple[str, ...]
    error: str


@dataclass(frozen=True)
class ChartGenerationResult:
    generated_files: tuple[Path, ...]
    failures: tuple[ChartGenerationFailure, ...]


@dataclass(frozen=True)
class BuildPresentationResult:
    output_path: Path
    replaced_pictures: int
    replaced_placeholders: int
    replaced_text: int
    generated_chart_count: int
    chart_failures: tuple[ChartGenerationFailure, ...]
    text_field_failures: tuple[TextFieldFailure, ...]
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


def _load_text_fields_config(
    *,
    repo_root: Path,
    cfg: Mapping[str, Any],
) -> tuple[Path, str | None, Sequence[object]]:
    text_fields_config = resolve_path(
        repo_root,
        str(cfg.get("text_fields_config", "config/text_fields.json")),
    )
    default_sheet, specs = parse_text_fields_json(text_fields_config)
    return text_fields_config, default_sheet, specs


def build_text_mapping(
    *,
    repo_root: Path,
    cfg: Mapping[str, Any],
    xlsx_path: Path,
    llm_payload: object | None,
) -> Dict[str, str]:
    return build_text_mapping_with_failures(
        repo_root=repo_root,
        cfg=cfg,
        xlsx_path=xlsx_path,
        llm_payload=llm_payload,
    ).mapping


def build_text_mapping_with_failures(
    *,
    repo_root: Path,
    cfg: Mapping[str, Any],
    xlsx_path: Path,
    llm_payload: object | None,
) -> TextFieldExtractionResult:
    text_fields_config, default_sheet, specs = _load_text_fields_config(
        repo_root=repo_root,
        cfg=cfg,
    )
    extraction_result = extract_xlsx_to_text_mapping_tolerant(
        xlsx_path,
        specs,
        default_sheet=default_sheet,
    )
    text_mapping = dict(extraction_result.mapping)

    llm_mapping = load_llm_mapping_from_payload(llm_payload)
    llm_mapping = _filter_llm_mapping(text_fields_config, llm_mapping)
    text_mapping.update(llm_mapping)
    return TextFieldExtractionResult(
        mapping=text_mapping,
        failures=extraction_result.failures,
    )


def _chart_generators() -> Sequence[ChartGeneratorSpec]:
    return (
        ChartGeneratorSpec(
            key="slide7",
            label="slide 7",
            generator=generate_slide7_charts,
            output_files=(
                "01_lucro_trimestres.png",
                "02_lucro_9m.png",
                "03_roe_trimestres.png",
                "04_roe_9m.png",
            ),
        ),
        ChartGeneratorSpec(
            key="slide10",
            label="slide 10",
            generator=generate_slide10_charts,
            output_files=(
                "05_qualidade_varejo_veiculos.png",
                "06_qualidade_total.png",
                "07_qualidade_atacado.png",
            ),
        ),
        ChartGeneratorSpec(
            key="slide3",
            label="slide 3",
            generator=generate_slide3_charts,
            output_files=(
                "08_emprestimos_empilhado.png",
                "09_seguros_cartoes_total.png",
            ),
        ),
        ChartGeneratorSpec(
            key="slide4",
            label="slide 4",
            generator=generate_slide4_charts,
            output_files=(
                "10_pizza_carteira.png",
                "11_pizza_trimestres.png",
                "12_pizza_9m.png",
            ),
        ),
        ChartGeneratorSpec(
            key="slide8",
            label="slide 8",
            generator=generate_slide8_charts,
            output_files=(
                "13_slide8_trimestres.png",
                "14_slide8_9m.png",
                "15_margem_financeira_bruta_total_trimestres.png",
                "16_margem_financeira_bruta_total_9m.png",
                "17_servicos_corretagem_trimestres.png",
                "18_servicos_corretagem_9m.png",
            ),
        ),
        ChartGeneratorSpec(
            key="slide9",
            label="slide 9",
            generator=generate_slide9_charts,
            output_files=(
                "09_custo_credito_trimestres.png",
                "09_custo_credito_9m.png",
                "09_indice_cobertura.png",
                "09_custo_variacao_custo_credito.png",
                "09_custo_variacao_custo_credito_9m.png",
            ),
        ),
        ChartGeneratorSpec(
            key="slide11",
            label="slide 11",
            generator=generate_slide11_charts,
            output_files=(
                "11_despesas_pessoal_adm_trimestres.png",
                "11_despesas_pessoal_adm_9m.png",
                "11_indice_eficiencia.png",
            ),
        ),
        ChartGeneratorSpec(
            key="slide12",
            label="slide 12",
            generator=generate_slide12_charts,
            output_files=("12_slide12_composicao.png",),
        ),
        ChartGeneratorSpec(
            key="slide13",
            label="slide 13",
            generator=generate_slide13_charts,
            output_files=(
                "13_varejo_produtos_entrada.png",
                "13_varejo_relacional.png",
                "13_atacado.png",
            ),
        ),
        ChartGeneratorSpec(
            key="slide14",
            label="slide 14",
            generator=generate_slide14_charts,
            output_files=(
                "14_veiculos_empilhado_trimestres.png",
                "14_veiculos_empilhado_anos.png",
                "14_veiculos_percentual_trimestres.png",
                "14_veiculos_percentual_anos.png",
            ),
        ),
        ChartGeneratorSpec(
            key="slide15",
            label="slide 15",
            generator=generate_slide15_charts,
            output_files=("15_captacoes_trimestres.png",),
        ),
        ChartGeneratorSpec(
            key="slide16",
            label="slide 16",
            generator=generate_slide16_charts,
            output_files=("16_indice_basileia_trimestres.png",),
        ),
        ChartGeneratorSpec(
            key="slide18",
            label="slide 18",
            generator=generate_slide18_charts,
            output_files=(
                "33_veiculos_empilhado.png",
                "34_premios_seguros_trimestres.png",
                "35_premios_seguros_9m.png",
            ),
        ),
        ChartGeneratorSpec(
            key="slide20",
            label="slide 20",
            generator=generate_slide20_charts,
            output_files=(
                "36_cib_empilhado.png",
                "37_carteira_atacado_comparativo.png",
            ),
        ),
    )


def _chart_generator_slide_number(spec: ChartGeneratorSpec) -> int | None:
    suffix = spec.key.removeprefix("slide")
    if spec.key.startswith("slide") and suffix.isdigit():
        return int(suffix)
    return None


def _select_chart_generators(
    only_slides: Sequence[int] | None,
) -> tuple[tuple[ChartGeneratorSpec, ...], tuple[int, ...]]:
    specs = tuple(_chart_generators())
    if not only_slides:
        return specs, ()

    requested = tuple(dict.fromkeys(int(slide) for slide in only_slides))
    available_slides = {
        slide_number
        for spec in specs
        for slide_number in [_chart_generator_slide_number(spec)]
        if slide_number is not None
    }
    selected = tuple(
        spec for spec in specs if _chart_generator_slide_number(spec) in set(requested)
    )
    ignored = tuple(slide for slide in requested if slide not in available_slides)
    if not selected:
        raise ValueError(
            "Nenhum dos slides informados possui gerador de graficos configurado. "
            f"Disponiveis: {', '.join(str(slide) for slide in sorted(available_slides))}"
        )
    return selected, ignored


def generate_chart_assets(
    *,
    xlsx_path: Path,
    images_dir: Path,
    only_slides: Sequence[int] | None = None,
) -> ChartGenerationResult:
    generated_files: list[Path] = []
    failures: list[ChartGenerationFailure] = []
    selected_specs, ignored_slides = _select_chart_generators(only_slides)

    if ignored_slides:
        logging.warning(
            "Slides sem geradores de graficos configurados foram ignorados: %s",
            ", ".join(str(slide) for slide in ignored_slides),
        )

    for spec in selected_specs:
        try:
            generated_files.extend(spec.generator(xlsx_path=xlsx_path, output_dir=images_dir))
        except Exception as exc:
            failures.append(
                ChartGenerationFailure(
                    generator_key=spec.key,
                    label=spec.label,
                    output_files=spec.output_files,
                    error=str(exc),
                )
            )
            logging.exception(
                "Falha ao gerar graficos de %s. Arquivos esperados: %s. "
                "O job vai continuar; se ja existirem PNGs anteriores, eles poderao ser reaproveitados no PPT.",
                spec.label,
                ", ".join(spec.output_files),
            )

    return ChartGenerationResult(
        generated_files=tuple(generated_files),
        failures=tuple(failures),
    )


def _persist_generated_chart_files(
    *,
    generated_files: Sequence[Path],
    target_dir: Path,
) -> tuple[Path, ...]:
    target_dir.mkdir(parents=True, exist_ok=True)

    persisted_files: list[Path] = []
    for generated_file in generated_files:
        source = Path(generated_file)
        destination = target_dir / source.name

        try:
            if source.resolve() == destination.resolve():
                persisted_files.append(destination)
                continue
        except FileNotFoundError:
            logging.warning(
                "Grafico gerado nao encontrado para copiar para images_dir: %s",
                source,
            )
            continue

        shutil.copy2(source, destination)
        persisted_files.append(destination)

    return tuple(persisted_files)


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
    only_slides: Sequence[int] | None = None,
) -> BuildPresentationResult:
    if not xlsx_path.exists():
        raise FileNotFoundError(f"XLSX nao encontrado: {xlsx_path}")

    _validate_xlsx_path(xlsx_path)

    normalized_only_slides = tuple(dict.fromkeys(int(slide) for slide in only_slides)) if only_slides else None
    if normalized_only_slides and skip_charts:
        raise ValueError("only_slides nao pode ser usado junto com skip_charts")

    effective_output_path = output_path or resolve_path(repo_root, str(cfg.get("pptx_output")))
    effective_images_dir = images_dir or resolve_path(repo_root, str(cfg.get("images_dir", ".")))

    pptx_template = resolve_path(repo_root, str(cfg.get("pptx_template")))
    allow_placeholder_text = bool(cfg.get("allow_placeholder_text", False))
    effective_llm_payload = llm_payload if llm_payload is not None else load_llm_payload_from_path(repo_root, cfg)
    _text_fields_config_path, _default_sheet, text_specs = _load_text_fields_config(
        repo_root=repo_root,
        cfg=cfg,
    )
    pp_field_ids = tuple(spec.id for spec in text_specs if getattr(spec, "is_pp", False))

    with TemporaryDirectory(prefix="ppt-doc-images-") as temp_images_dir:
        if normalized_only_slides and images_dir is None:
            active_images_dir = Path(temp_images_dir)
        else:
            active_images_dir = effective_images_dir
            active_images_dir.mkdir(parents=True, exist_ok=True)

        generated_chart_count = 0
        chart_failures: tuple[ChartGenerationFailure, ...] = ()
        text_field_failures: tuple[TextFieldFailure, ...] = ()
        if not skip_charts:
            chart_generation = generate_chart_assets(
                xlsx_path=xlsx_path,
                images_dir=active_images_dir,
                only_slides=normalized_only_slides,
            )
            generated_chart_count = len(chart_generation.generated_files)
            chart_failures = chart_generation.failures
            if active_images_dir != effective_images_dir:
                _persist_generated_chart_files(
                    generated_files=chart_generation.generated_files,
                    target_dir=effective_images_dir,
                )

        text_mapping_result = build_text_mapping_with_failures(
            repo_root=repo_root,
            cfg=cfg,
            xlsx_path=xlsx_path,
            llm_payload=effective_llm_payload,
        )
        text_mapping = text_mapping_result.mapping
        text_field_failures = text_mapping_result.failures

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
            images_dir=active_images_dir,
            allow_placeholder_text=allow_placeholder_text,
            text_json=None,
            text_payload=text_mapping,
            pp_field_ids=pp_field_ids,
        )

    return BuildPresentationResult(
        output_path=effective_output_path,
        replaced_pictures=replaced_pictures,
        replaced_placeholders=replaced_placeholders,
        replaced_text=replaced_text,
        generated_chart_count=generated_chart_count,
        chart_failures=chart_failures,
        text_field_failures=text_field_failures,
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
            chart_failures=result.chart_failures,
            text_field_failures=result.text_field_failures,
            applied_text_keys=result.applied_text_keys,
        )
        return output_path.read_bytes(), logical_result
