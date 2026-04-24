from pathlib import Path
from urllib.parse import unquote

import requests


def filename_from_content_disposition(header_value):
    if not header_value:
        return None

    for part in header_value.split(";"):
        part = part.strip()
        if part.lower().startswith("filename*="):
            value = part.split("=", 1)[1].strip().strip('"')
            if "''" in value:
                value = value.split("''", 1)[1]
            return Path(unquote(value)).name

    for part in header_value.split(";"):
        part = part.strip()
        if part.lower().startswith("filename="):
            return Path(part.split("=", 1)[1].strip().strip('"')).name

    return None


xlsx_path = Path(r"users\testing\excel_vfinal.xlsx")
llm_response_path = Path(r"users\testing\llm_response.json")
output_dir = Path(r"users\testing")

with xlsx_path.open("rb") as xlsx_file, llm_response_path.open("rb") as llm_response_file:
    response = requests.post(
        "http://localhost:8000/api/ri-report-generate",
        files={
            "xlsx_file": (
                xlsx_path.name,
                xlsx_file,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            "llm_response_file": (
                llm_response_path.name,
                llm_response_file,
                "application/json"
                )
        },
        verify=False,
        stream=True,
        timeout=60,
    )

    print("Status Code:", response.status_code)
    response.raise_for_status()

    output_filename = filename_from_content_disposition(
        response.headers.get("Content-Disposition")
    ) or "output.pptx"
    output_path = output_dir / output_filename

    with output_path.open("wb") as output_file:
        for chunk in response.iter_content(chunk_size=1024*1024):
            if chunk:
                output_file.write(chunk)

    print("PPTX salvo em:", output_path)
