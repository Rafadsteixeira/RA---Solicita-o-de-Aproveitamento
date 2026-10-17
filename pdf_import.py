"""Extração local de fichas de aluno a partir de relatórios PDF."""

from __future__ import annotations

import re
import unicodedata
from pathlib import Path


class PDFStudentImportError(Exception):
    """Erro compreensível ao ler ou interpretar um relatório de alunos."""


FIELD_ALIASES = {
    "id_aluno": (
        "id do aluno", "id aluno", "matrícula do aluno", "matricula do aluno",
        "matrícula", "matricula", "número de matrícula",
        "numero de matricula", "nº de matrícula", "no de matricula",
        "registro acadêmico", "registro academico", "ra",
        "código do aluno", "codigo do aluno", "identificação do aluno",
        "identificacao do aluno", "registro do aluno", "código", "codigo",
        "número do aluno",
        "numero do aluno", "id",
    ),
    "nome_aluno": (
        "nome completo do aluno", "nome do aluno", "nome do estudante", "nome completo",
        "aluno", "nome",
    ),
    "rg": (
        "rg do aluno", "registro geral", "documento de identidade", "identidade", "rg",
    ),
    "cpf": ("cpf do aluno", "cpf do estudante", "cpf", "cadastro de pessoa física"),
    "data_nascimento": (
        "data de nascimento", "data nascimento", "nascimento", "dt. nascimento",
        "data de nasc.", "data nasc.", "dt nascimento", "dt. nasc.", "dt nasc.", "nasc.",
    ),
    "data_matricula": (
        "data de matrícula", "data de matricula", "data matrícula", "data matricula",
        "data da matrícula", "data da matricula", "matriculado em", "data de ingresso",
        "data ingresso", "dt. matrícula", "dt. matricula",
    ),
    "telefone": ("telefone celular", "telefone", "celular", "fone", "whatsapp"),
    "email": ("e-mail", "email", "correio eletrônico", "correio eletronico"),
    "curso_origem": ("curso de origem", "curso matriculado", "curso atual", "curso"),
    "ano_cursado": ("ano cursado", "ano de ingresso", "ano de matrícula", "ano de matricula"),
}


def _normalize(value: str) -> str:
    value = unicodedata.normalize("NFKD", value.casefold())
    value = "".join(char for char in value if not unicodedata.combining(char))
    return re.sub(r"[^a-z0-9]+", " ", value).strip()


_ALIAS_TO_FIELD = {
    _normalize(alias): field
    for field, aliases in FIELD_ALIASES.items()
    for alias in aliases
}
_LABELS = sorted(
    (re.escape(alias).replace(r"\ ", r"\s+") for aliases in FIELD_ALIASES.values() for alias in aliases),
    key=len,
    reverse=True,
)
_LABEL_WITH_SEPARATOR = re.compile(
    r"(?<![\w])(" + "|".join(_LABELS) + r")(?:\s*[:\uFF1A=]\s*|\s+)",
    re.IGNORECASE,
)


def _clean(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip(" \t\r\n:;|,-")


def _clean_date(value: str) -> str:
    match = re.search(r"\b(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})\b", value)
    if match:
        day, month, year = match.groups()
        if len(year) == 2:
            year = "20" + year
        return f"{int(day):02d}/{int(month):02d}/{year}"
    iso_match = re.search(r"\b(\d{4})-(\d{1,2})-(\d{1,2})\b", value)
    if iso_match:
        year, month, day = iso_match.groups()
        return f"{int(day):02d}/{int(month):02d}/{year}"
    return _clean(value)


def _put(record: dict[str, str], field: str, value: str):
    value = _clean(value)
    if not value:
        return
    if field in ("data_nascimento", "data_matricula"):
        value = _clean_date(value)
    if not record.get(field):
        record[field] = value


def _extract_senac_table_rows(lines: list[str]) -> list[dict[str, str]]:
    """Extrai as linhas da tabela do Relatório de Dados e Status do Aluno."""
    date_pattern = re.compile(r"(?<!\d)(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})(?!\d)")
    records = []
    for line in lines:
        dates = list(date_pattern.finditer(line))
        if len(dates) < 2:
            continue

        # In this report the table order is ID, student name, RG, CPF,
        # birth date, other details, and enrollment date. Empty columns
        # therefore do not shift the fields we need.
        prefix = line[:dates[0].start()]
        tokens = list(re.finditer(r"\S+", prefix))
        numeric_tokens = [
            (index, token, re.sub(r"\D", "", token.group()))
            for index, token in enumerate(tokens)
        ]
        id_item = next(
            (item for item in numeric_tokens if 6 <= len(item[2]) <= 12),
            None,
        )
        cpf_items = [item for item in numeric_tokens if len(item[2]) == 11]
        if id_item is None or not cpf_items:
            continue
        cpf_item = cpf_items[-1]
        if cpf_item[0] <= id_item[0]:
            continue

        rg_item = next(
            (
                item for item in reversed(numeric_tokens)
                if id_item[0] < item[0] < cpf_item[0] and len(item[2]) >= 6
            ),
            None,
        )
        name_end = rg_item[1].start() if rg_item else cpf_item[1].start()
        name = _clean(prefix[id_item[1].end():name_end])
        rg = _clean(prefix[rg_item[1].start():cpf_item[1].start()]) if rg_item else ""
        record = {
            "id_aluno": id_item[1].group(),
            "nome_aluno": name,
            "rg": rg,
            "cpf": cpf_item[1].group(),
            "data_nascimento": _clean_date(dates[0].group()),
            "data_matricula": _clean_date(dates[-1].group()),
        }
        if name:
            records.append(record)
    return records


def _extract_page(text: str) -> list[dict[str, str]]:
    lines = text.splitlines()
    labeled_record: dict[str, str] = {}
    for line_index, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue

        matches = list(_LABEL_WITH_SEPARATOR.finditer(line))
        if len(matches) > 1:
            segments = [
                line[match.end():matches[index + 1].start() if index + 1 < len(matches) else len(line)]
                for index, match in enumerate(matches)
            ]
            # A row containing only column headings has labels but no values.
            if all(not _clean(segment) for segment in segments):
                matches = []
        for match_index, match in enumerate(matches):
            field = _ALIAS_TO_FIELD.get(_normalize(match.group(1)))
            if not field:
                continue
            end = matches[match_index + 1].start() if match_index + 1 < len(matches) else len(line)
            _put(labeled_record, field, line[match.end():end])

        label_only = _normalize(line.rstrip(":\uff1a- "))
        field = _ALIAS_TO_FIELD.get(label_only)
        if field and not labeled_record.get(field):
            for following in lines[line_index + 1:]:
                candidate = following.strip()
                if not candidate:
                    continue
                if _normalize(candidate.rstrip(":\uff1a- ")) in _ALIAS_TO_FIELD:
                    break
                _put(labeled_record, field, candidate)
                break

    senac_records = _extract_senac_table_rows(lines)
    if senac_records:
        return senac_records

    # A table may contain multiple students on the same page.
    cells_by_line = [re.split(r"\t+|\s{2,}", line.strip()) for line in lines if line.strip()]
    table_records: list[dict[str, str]] = []
    for header_index, cells in enumerate(cells_by_line[:-1]):
        columns = [_ALIAS_TO_FIELD.get(_normalize(cell.rstrip(":\uff1a"))) for cell in cells]
        if sum(field is not None for field in columns) < 2:
            continue
        if not any(field in ("id_aluno", "nome_aluno", "cpf") for field in columns):
            continue

        for value_row in cells_by_line[header_index + 1:]:
            row_columns = [_ALIAS_TO_FIELD.get(_normalize(cell.rstrip(":\uff1a"))) for cell in value_row]
            if sum(field is not None for field in row_columns) >= 2 and any(
                field in ("id_aluno", "nome_aluno", "cpf") for field in row_columns
            ):
                break
            if len(value_row) != len(columns):
                continue
            candidate: dict[str, str] = {}
            for field, value in zip(columns, value_row):
                if field:
                    _put(candidate, field, value)
            if any(candidate.get(field) for field in ("id_aluno", "nome_aluno", "cpf")):
                table_records.append(candidate)
        if table_records:
            break

    if not table_records:
        return [labeled_record] if labeled_record else []
    if len(table_records) == 1:
        for field, value in labeled_record.items():
            if value and not table_records[0].get(field):
                table_records[0][field] = value
    elif labeled_record.get("cpf") or labeled_record.get("id_aluno"):
        wanted_cpf = re.sub(r"\D", "", labeled_record.get("cpf", ""))
        wanted_id = labeled_record.get("id_aluno", "").strip().casefold()
        for candidate in table_records:
            same_cpf = wanted_cpf and re.sub(r"\D", "", candidate.get("cpf", "")) == wanted_cpf
            same_id = wanted_id and candidate.get("id_aluno", "").strip().casefold() == wanted_id
            if same_cpf or same_id:
                for field, value in labeled_record.items():
                    if value and not candidate.get(field):
                        candidate[field] = value
                break
    return table_records


def extract_students_data(pdf_path: str | Path) -> list[dict[str, str]]:
    """Lê todas as páginas e retorna as fichas reconhecidas no PDF."""
    try:
        from pypdf import PdfReader
    except ImportError as exc:
        raise PDFStudentImportError(
            "A biblioteca pypdf não está instalada. Feche o sistema e execute run_sistema.bat "
            "para instalar as dependências."
        ) from exc

    try:
        reader = PdfReader(str(Path(pdf_path)))
        if reader.is_encrypted:
            try:
                unlocked = reader.decrypt("")
            except Exception:
                unlocked = 0
            if not unlocked:
                raise PDFStudentImportError(
                    "Este PDF está protegido por senha. Salve uma cópia sem senha e importe novamente."
                )
        page_texts = []
        for page in reader.pages:
            try:
                text = page.extract_text(extraction_mode="layout")
            except (TypeError, NotImplementedError):
                text = page.extract_text()
            except Exception:
                text = page.extract_text()
            page_texts.append(text or "")
    except PDFStudentImportError:
        raise
    except Exception as exc:
        if "cryptography" in str(exc).casefold() or "aes algorithm" in str(exc).casefold():
            raise PDFStudentImportError(
                "Este PDF usa criptografia AES e falta a biblioteca cryptography neste interpretador. "
                "Feche o sistema e execute run_sistema.bat para instalar as dependências. "
                "Se estiver usando o IDE, execute: py -m pip install -r requirements.txt"
            ) from exc
        raise PDFStudentImportError(f"Não consegui ler o PDF selecionado:\n{exc}") from exc

    if not any(text.strip() for text in page_texts):
        raise PDFStudentImportError(
            "Não encontrei texto selecionável. PDFs digitalizados como imagem precisam passar por OCR."
        )

    records = [record for text in page_texts for record in _extract_page(text)]
    if not records:
        raise PDFStudentImportError(
            "O PDF foi lido, mas não encontrei os rótulos de ID, nome, RG, CPF ou datas."
        )

    # A multi-page report can repeat one student's details; merge those pages,
    # while keeping separate IDs/CPFs as separate pre-registrations.
    merged: dict[str, dict[str, str]] = {}
    for index, record in enumerate(records):
        cpf = re.sub(r"\D", "", record.get("cpf", ""))
        student_id = record.get("id_aluno", "").strip().casefold()
        name = _normalize(record.get("nome_aluno", ""))
        birth_date = _normalize(record.get("data_nascimento", ""))
        identity = cpf or student_id
        if not identity and name and birth_date:
            identity = f"{name}|{birth_date}"
        if not identity:
            identity = f"record-{index}"
        current = merged.setdefault(identity, {})
        for field, value in record.items():
            if value and not current.get(field):
                current[field] = value
    return list(merged.values())
