import json
import os
from pathlib import Path
from typing import Any, Dict

from genai_framework.decorators import file_input_route #framework corporativo
from genai_framework.models import FileInput #framework corporativo
from google.genai import Client, types

from utils.json_utils import coerce_json
from utils.xlsx_extract import extract_xlsx_bytes_to_dict, parse_specs_json


def _resolve_default_specs_path() -> str:
    """Resolve config/specs.json regardless of where this file lives.

    Works when this module is in repo root or under src/routes/.
    """

    here = Path(__file__).resolve()
    for parent in [here.parent, *here.parents]:
        candidate = parent / "config" / "specs.json"
        if candidate.exists():
            return str(candidate)
    return str(Path("config") / "specs.json")
    
@file_input_route("analyze_file")
def analyze_file(file: FileInput):
    client = Client(vertexai = True, project=os.getenv('project_id'), location=os.getenv('location'))

    INSTRUCTIONS = """
          Você é um especialista em criar frases de apresentações profissionais de relacionamento para investidores.
          Você entende todas as nomenclaturas financeiras e de negócios.
          Voce sabe como demonstrar resultados financeiros de forma clara e objetiva.
          Voce deve apenas responder em JSON Valido conforme esquema fornecido.
          NUNCA responda com texto que nao seja JSON valido.

          Esquema JSON que voce deve receber:
            {
                "lucroTrimestre": {
                    "labels": [
                    "3T23", "4T23", "1T24", "2T24", "3T24", "4T24", "1T25", "2T25", "3T25"
                    ],
                    "values": [
                    285.0, 302.0, 321.0, 363.0, 496.0, 542.0, 480.0, 459.0, 461.0
                    ],
                    "sheet": "DRE Saida",
                    "ranges": {
                    "labels": "C3:K3",
                    "values": "C18:K18"
                    }
                },
                "lucro9M": {
                    "labels": [
                    "9M24", "9M25"
                    ],
                    "values": [
                    1180.0, 1400.0
                    ],
                    "sheet": "DRE Saida",
                    "ranges": {
                    "labels": "L3:M3",
                    "values": "L18:M18"
                    }
                }
                ... possivelmente mais entradas ...
            }

        Esquema JSON que voce deve responder:
            {
                titles: {
                "slide1_title": "No 3T25, tivemos um lucro de R$ 461 milhões, representando um crescimento de 61% em relação ao 3T24.",
                "slide2_title": "No acumulado dos 9 meses de 2025, nosso lucro atingiu R$ 1,4 bilhão, um aumento de 19% em comparação com os 9 meses de 2024."
                ... mais slides conforme necessario ...
                },
                subtitles: {
                "slide1_subtitle": "Esse desempenho reflete nossa estratégia focada em eficiência operacional e expansão de mercado.",
                "slide2_subtitle": "Esse resultado reforça nosso compromisso com a criação de valor sustentável para nossos acionistas."
                ... mais slides conforme necessario ...
                }
             
            } 

        Instruções para criação das frases:
        SLIDE 1 - LUCRO LÍQUIDO TRIMESTRAL
            Para slide1_title:
                - A frase deve seguir o exemplo a seguir "Lucro Líquido cresceu X% frente ao Y, com sólido avanço de Z% no acumulado de 9MXX." 
                - X seria: a porcentagem de ganho do trimestre mais atual versus o trimestre imediatamente anterior, no exemplo atual seriam 3T25 (mais atual), 2T25 (anterior).  
                    - Os valores estão no campo de "values" e os labels no campo "labels" dentro do objeto "lucroTrimestre". Valor esperado no exemplo: 0,4%
                - Y seria: o trimestre imediatamente anterior ao mais atual, no exemplo atual seria 2T25.
                    - Os valores estão no campo "values" e os labels no campo "labels" do objeto "lucro9M". Valor esperado no exemplo:  18,6%
                - Z seria: a porcentagem de ganho do acumulado de 9 meses do ano mais atual versus o acumulado de 9 meses do ano anterior, no exemplo atual seriam 9M25 (mais atual), 9M24 (anterior).

            Para slide1_subtitle:
                - A frase deve seguir o exemplo a seguir "Resultados refletem os avanços na execução da nossa estratégia: fortalecer e sustentar o core business, diversificar receitas e fortalecer abordagem relacional com nossos clientes pessoas fisicas e jurídicas."
                - A frase deve ser formal e profissional, destacando a importância do resultado financeiro apresentado.  
                - Use termos que transmitam confiança e competência na gestão financeira da empresa.
            
            
        Regras importantes:
          - Sempre responda com JSON valido.
          - Todos os campos devem ser em camelCase.
          - Nao inclua nenhum texto fora do JSON.
          - Se algum valor necessário para o calculo estiver faltando, responda com null no campo correspondente.
          - Use ponto (,) como separador decimal.
          - Todos os campos devem retornar como strings.

          """
    
    try:
        specs_path = os.getenv("SPECS_JSON_PATH") or _resolve_default_specs_path()
        default_sheet = os.getenv("DEFAULT_SHEET")

        specs = parse_specs_json(specs_path)
        extracted: Dict[str, Any] = extract_xlsx_bytes_to_dict(
            file.content,
            specs,
            default_sheet=default_sheet,
            include_meta=True,
            lowercase_fields=True,
        )
        texto = json.dumps(extracted, ensure_ascii=False)
    except Exception as exc:
        return {
            "error": "Falha ao extrair dados do XLSX usando specs.json.",
            "details": str(exc),
        }
    
    prompt = f"{INSTRUCTIONS}\n\nDados extraídos:\n{texto}"

    response = client.models.generate_content(
        model="gemini-2.5-flash-lite",
        contents=prompt,
        config=types.GenerateContentConfig(temperature=0.2, max_output_tokens=65530)
    )

    try:
        txt = coerce_json(response.text)
        return {"response": txt}
    except json.JSONDecodeError as e:
        return {
            "error": "Failed to parse JSON from model response.",
            "details": str(e),
            "model_response": response.text,
        }

        
