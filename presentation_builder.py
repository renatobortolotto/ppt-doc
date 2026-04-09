from __future__ import annotations

import json
import logging
import shutil
import unicodedata
import zipfile
from dataclasses import dataclass
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any, Callable, Dict, Mapping, Sequence
from xml.etree import ElementTree as ET

from update_ppt import _flatten_text_payload, update_presentation
from utils.slides.slide11_charts import generate_slide11_charts
from utils.slides.slide12_charts import generate_slide12_charts
from utils.slides.slide13_charts import generate_slide13_charts
from utils.slides.slide14_charts import generate_slide14_charts
from utils.slides.slide15_charts import generate_slide15_charts
from utils.slides.slide16_charts import generate_slide16_charts
from utils.slides.slide20_charts import generate_slide20_charts
from utils.slides.slide21_charts import generate_slide21_charts
from utils.slides.slide22_charts import generate_slide22_charts
from utils.slides.slide19_charts import generate_slide19_charts
from utils.slides.slide8_charts import generate_slide8_charts
from utils.slides.slide9_charts import generate_slide9_charts
from utils.slides.slide4_charts import generate_slide4_charts

try:
    from utils.slides.slide18_charts import generate_slide18_charts
except ModuleNotFoundError:
    def generate_slide18_charts(*, xlsx_path: Path, output_dir: Path) -> list[Path]:
        raise ModuleNotFoundError("utils.slides.slide18_charts nao esta disponivel no workspace atual")
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


@dataclass(frozen=True)
class XlsxPrivacyCheckResult:
    status: str
    label: str | None
    source: str | None


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
            key="slide19",
            label="slide 19",
            generator=generate_slide19_charts,
            output_files=(
                "19_veiculos_empilhado.png",
                "19_seguros_cartoes_trimestres.png",
                "19_seguros_cartoes_anos.png",
            ),
        ),
        ChartGeneratorSpec(
            key="slide20",
            label="slide 20",
            generator=generate_slide20_charts,
            output_files=(
                "20_emprestimos_empilhado.png",
                "20_seguros_cartoes_trimestres.png",
            ),
        ),
        ChartGeneratorSpec(
            key="slide21",
            label="slide 21",
            generator=generate_slide21_charts,
            output_files=(
                "21_carteira_atacado_empilhado.png",
                "21_carteira_atacado_comparativo.png",
            ),
        ),
        ChartGeneratorSpec(
            key="slide22",
            label="slide 22",
            generator=generate_slide22_charts,
            output_files=(
                "22_qualidade_4966_bloco1.png",
                "22_qualidade_4966_bloco2.png",
                "22_qualidade_4966_bloco3.png",
                "22_qualidade_4966_linha1.png",
                "22_qualidade_4966_linha2.png",
                "22_qualidade_4966_linha3.png",
                "22_carteira_reestruturada_barras.png",
                "22_carteira_reestruturada_linha.png",
                "22_cobertura_reestruturada_linha.png",
                "22_npl_barras.png",
                "22_npl_linha.png",
            ),
        ),
    )


def _chart_generator_slide_number(spec: ChartGeneratorSpec) -> int | None:
    suffix = spec.key.removeprefix("slide")
    if spec.key.startswith("slide") and suffix.isdigit():
        return int(suffix)
    return None


def _normalize_privacy_text(value: object) -> str:
    text = str(value or "").strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return " ".join(text.split())


def _xlsx_privacy_public_values(cfg: Mapping[str, Any]) -> set[str]:
    raw_values = cfg.get("xlsx_privacy_public_values")
    if isinstance(raw_values, list):
        values = {_normalize_privacy_text(value) for value in raw_values if str(value).strip()}
        if values:
            return values
    return {"publico", "public", "publica"}


def _read_xlsx_metadata_text_properties(xlsx_path: Path) -> dict[str, str]:
    properties: dict[str, str] = {}
    with zipfile.ZipFile(xlsx_path) as zf:
        if "docProps/custom.xml" in zf.namelist():
            custom_root = ET.fromstring(zf.read("docProps/custom.xml"))
            custom_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/custom-properties}"
            for prop in custom_root.findall(f"{custom_ns}property"):
                name = (prop.attrib.get("name") or "").strip()
                if not name:
                    continue
                text = "".join(prop.itertext()).strip()
                if text:
                    properties[f"custom:{name}"] = text

        if "docProps/core.xml" in zf.namelist():
            core_root = ET.fromstring(zf.read("docProps/core.xml"))
            for child in core_root:
                local_name = child.tag.rsplit("}", 1)[-1]
                if local_name not in {
                    "category",
                    "contentStatus",
                    "description",
                    "keywords",
                    "subject",
                    "title",
                }:
                    continue
                text = "".join(child.itertext()).strip()
                if text:
                    properties[f"core:{local_name}"] = text

    return properties


def inspect_xlsx_privacy(
    xlsx_path: Path,
    cfg: Mapping[str, Any],
) -> XlsxPrivacyCheckResult:
    public_values = _xlsx_privacy_public_values(cfg)
    properties = _read_xlsx_metadata_text_properties(xlsx_path)

    strong_name_tokens = (
        "classification",
        "classificacao",
        "privacy",
        "privacidade",
        "sensitivity",
        "sensibilidade",
    )
    fallback_name_tokens = ("label", "rotulo")
    non_public_tokens = (
        "confidencial",
        "confidential",
        "internal",
        "interno",
        "private",
        "privado",
        "restricted",
        "restrito",
        "sigiloso",
    )

    strong_candidates: list[tuple[str, str, bool]] = []
    fallback_candidates: list[tuple[str, str, bool]] = []
    for source, label in properties.items():
        source_name = source.split(":", 1)[-1]
        normalized_source = _normalize_privacy_text(source_name)
        normalized_label = _normalize_privacy_text(label)
        if not normalized_label or not any(ch.isalpha() for ch in normalized_label):
            continue

        if normalized_source.startswith("msip_label_") and normalized_source.endswith("_name"):
            strong_candidates.append((source, label, True))
            continue
        if any(token in normalized_source for token in strong_name_tokens):
            strong_candidates.append((source, label, True))
            continue
        if any(token in normalized_source for token in fallback_name_tokens):
            fallback_candidates.append((source, label, False))
            continue
        if source.startswith("core:"):
            fallback_candidates.append((source, label, False))

    for source, label, strong in (*strong_candidates, *fallback_candidates):
        normalized_label = _normalize_privacy_text(label)
        if normalized_label in public_values:
            return XlsxPrivacyCheckResult(status="public", label=label, source=source)
        if any(token in normalized_label for token in non_public_tokens):
            return XlsxPrivacyCheckResult(status="non_public", label=label, source=source)
        if strong:
            return XlsxPrivacyCheckResult(status="non_public", label=label, source=source)

    return XlsxPrivacyCheckResult(status="unknown", label=None, source=None)


def _validate_xlsx_privacy(
    xlsx_path: Path,
    cfg: Mapping[str, Any],
) -> None:
    mode = str(cfg.get("xlsx_privacy_check_mode", "error")).strip().lower()
    if mode in {"", "off", "disabled", "none"}:
        return

    try:
        result = inspect_xlsx_privacy(xlsx_path, cfg)
    except zipfile.BadZipFile:
        # O parse estrutural do XLSX e a mensagem principal já são tratados em _validate_xlsx_path.
        return
    except ET.ParseError as exc:
        logging.warning(
            "Nao foi possivel interpretar os metadados de privacidade do XLSX %s: %s",
            xlsx_path,
            exc,
        )
        return

    if result.status == "public":
        logging.info(
            "Privacidade do XLSX validada como publico via %s=%r",
            result.source,
            result.label,
        )
        return

    if result.status == "non_public":
        message = (
            "Arquivo Excel com privacidade diferente de publico detectada: "
            f"{result.label!r} ({result.source}). "
            "Envie um XLSX classificado como Publico."
        )
        if mode == "warn":
            logging.warning(message)
            return
        raise ValueError(message)

    logging.info(
        "Nenhum metadado reconhecido de privacidade/classificacao foi encontrado no XLSX %s; "
        "seguindo em frente.",
        xlsx_path,
    )


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
            logging.info(
                "Iniciando geracao de graficos do %s (%s)",
                spec.label,
                spec.key,
            )
            generated = spec.generator(xlsx_path=xlsx_path, output_dir=images_dir)
            generated_files.extend(generated)
            logging.info(
                "Geracao de graficos concluida com sucesso para %s (%s): arquivos=%s",
                spec.label,
                spec.key,
                ", ".join(path.name for path in generated) if generated else "<nenhum-arquivo>",
            )
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
    _validate_xlsx_privacy(xlsx_path, cfg)

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
            xlsx_path=xlsx_path,
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
