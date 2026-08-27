#!/usr/bin/env python3
"""
clean_requests.py — приведение "сырых" заявок к единому виду.

Использование:
    python3 clean_requests.py входной_файл.xlsx [--outdir DIR]

На вход: .xlsx (или .csv) с колонками: Имя, Телефон, Дата заявки, Источник
На выход:
    clean_<имя_файла>.xlsx    — чистые уникальные записи
    problems_<имя_файла>.xlsx — проблемные строки (с причиной)

Правила очистки:
    Телефон -> +7XXXXXXXXXX (10 цифр после +7).
    Дата -> ГГГГ-ММ-ДД.
    Имя -> Слова с заглавной буквы, лишние пробелы убраны.
    Дубли -> по телефону, оставляем первую запись.
    Проблемные строки -> отдельный список.
"""

import argparse
import re
import sys
from pathlib import Path

import pandas as pd


RUSSIAN_MONTHS = {
    "январ": 1,
    "феврал": 2,
    "март": 3,
    "апрел": 4,
    "май": 5,
    "июн": 6,
    "июл": 7,
    "август": 8,
    "сентябр": 9,
    "октябр": 10,
    "ноябр": 11,
    "декабр": 12,
}


def normalize_phone(raw):
    """Возвращает (phone_or_None, is_problem: bool)."""

    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return None, True

    # Excel может прочитать номер как 79991234567.0
    if isinstance(raw, float) and raw.is_integer():
        raw = int(raw)

    s = str(raw).strip()
    digits = re.sub(r"\D", "", s)

    if not digits:
        return None, True

    # 11 цифр, начинается с 7 или 8
    if len(digits) == 11 and digits[0] in ("7", "8"):
        digits = "7" + digits[1:]
        return f"+{digits}", False

    # 10 цифр без кода страны
    if len(digits) == 10:
        return f"+7{digits}", False

    return None, True


def normalize_date(raw):
    """Возвращает (date_str_or_None, is_problem: bool)."""

    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return None, True

    # Если pandas уже распознала дату
    if isinstance(raw, pd.Timestamp):
        if pd.isna(raw):
            return None, True
        return raw.strftime("%Y-%m-%d"), False

    s = str(raw).strip().lower()

    if not s or s in ("nan", "nat"):
        return None, True

    # Заглушки вроде "хх.хх.2026" или "xx.xx.2026"
    if re.search(r"[хx]{2,}", s):
        return None, True

    # Приводим разделители к точке
    s_norm = s.replace("/", ".").replace("-", ".").replace(" ", ".")
    s_norm = re.sub(r"\.+", ".", s_norm).strip(".")

    # Формат yyyy.mm.dd
    m = re.fullmatch(r"(\d{4})\.(\d{1,2})\.(\d{1,2})", s_norm)
    if m:
        y, mo, d = map(int, m.groups())
        return _try_build_date(y, mo, d)

    # Формат dd.mm.yyyy / dd.mm.yy
    m = re.fullmatch(r"(\d{1,2})\.(\d{1,2})\.(\d{2,4})", s_norm)
    if m:
        d, mo, y = m.groups()

        d = int(d)
        mo = int(mo)
        y = int(y)

        if y < 100:
            y += 2000

        # Если второй компонент больше 12, возможно формат mm.dd.yyyy
        if mo > 12 and d <= 12:
            d, mo = mo, d

        return _try_build_date(y, mo, d)

    # Формат "13 марта 2026"
    m = re.fullmatch(r"(\d{1,2})\s+([а-яё]+)\s*(\d{4})?", s)
    if m:
        d, month_word, y = m.groups()

        month_num = None

        for stem, num in RUSSIAN_MONTHS.items():
            if month_word.startswith(stem):
                month_num = num
                break

        if month_num is None or y is None:
            return None, True

        return _try_build_date(
            int(y),
            month_num,
            int(d),
        )

    return None, True


def _try_build_date(y, mo, d):
    """Проверяет существование даты и возвращает YYYY-MM-DD."""

    try:
        return (
            pd.Timestamp(
                year=y,
                month=mo,
                day=d,
            ).strftime("%Y-%m-%d"),
            False,
        )

    except (ValueError, TypeError):
        return None, True


def normalize_name(raw):
    """Возвращает (name_or_None, is_problem: bool)."""

    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return None, True

    s = re.sub(r"\s+", " ", str(raw).strip())

    if not s:
        return None, True

    fixed = []

    for word in s.split():
        fixed.append(
            word[0].upper() + word[1:].lower()
        )

    return " ".join(fixed), False


def _write_with_autofit(df: pd.DataFrame, path: Path):
    """Сохраняет DataFrame в xlsx с шириной колонок по содержимому."""

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(
            writer,
            index=False,
            sheet_name="Sheet1",
        )

        ws = writer.sheets["Sheet1"]

        for i, col in enumerate(df.columns, start=1):
            max_len = max(
                [len(str(col))]
                + [len(str(v)) for v in df[col].tolist()]
            )

            ws.column_dimensions[
                ws.cell(
                    row=1,
                    column=i,
                ).column_letter
            ].width = max_len + 4


def process(input_path: Path, outdir: Path):
    """Читает файл, очищает данные и сохраняет результаты."""

    if not input_path.exists():
        sys.exit(
            f"Ошибка: файл не найден: {input_path}"
        )

    suffix = input_path.suffix.lower()

    try:
        if suffix == ".csv":
            df = pd.read_csv(input_path)

        elif suffix == ".xlsx":
            df = pd.read_excel(input_path)

        else:
            sys.exit(
                "Ошибка: поддерживаются только файлы .xlsx и .csv"
            )

    except Exception as exc:
        sys.exit(
            f"Ошибка при чтении файла "
            f"'{input_path}': {exc}"
        )

    required_cols = [
        "Имя",
        "Телефон",
        "Дата заявки",
        "Источник",
    ]

    missing_cols = [
        col
        for col in required_cols
        if col not in df.columns
    ]

    if missing_cols:
        sys.exit(
            f"Ошибка: отсутствуют колонки: {missing_cols}. "
            f"Найдены колонки: {list(df.columns)}"
        )

    rows = []
    problems = []

    for idx, row in df.iterrows():

        name, name_problem = normalize_name(
            row["Имя"]
        )

        phone, phone_problem = normalize_phone(
            row["Телефон"]
        )

        date, date_problem = normalize_date(
            row["Дата заявки"]
        )

        raw_source = row["Источник"]

        source = (
            None
            if pd.isna(raw_source)
            else str(raw_source).strip()
        )

        reasons = []

        if name_problem:
            reasons.append("нет имени")

        if phone_problem:
            reasons.append("нет телефона")

        if date_problem:
            reasons.append("битая дата")

        record = {
            "Имя": name,
            "Телефон": (
                phone
                if phone
                else "нет телефона"
            ),
            "Дата заявки": (
                date
                if date
                else "проблема с датой"
            ),
            "Источник": source,
        }

        if reasons:
            problems.append({
                **record,
                "Исходная строка": idx + 2,
                "Причина": "; ".join(reasons),
                "Исходное имя": row["Имя"],
                "Исходный телефон": row["Телефон"],
                "Исходная дата": row["Дата заявки"],
            })

        else:
            rows.append(record)

    clean_columns = [
        "Имя",
        "Телефон",
        "Дата заявки",
        "Источник",
    ]

    clean_df = pd.DataFrame(
        rows,
        columns=clean_columns,
    )

    # Дедупликация только среди полностью чистых записей
    before = len(clean_df)

    clean_df = (
        clean_df
        .drop_duplicates(
            subset=["Телефон"],
            keep="first",
        )
        .reset_index(drop=True)
    )

    duplicates_removed = before - len(clean_df)

    problem_columns = [
        "Имя",
        "Телефон",
        "Дата заявки",
        "Источник",
        "Исходная строка",
        "Причина",
        "Исходное имя",
        "Исходный телефон",
        "Исходная дата",
    ]

    problems_df = pd.DataFrame(
        problems,
        columns=problem_columns,
    )

    outdir.mkdir(
        parents=True,
        exist_ok=True,
    )

    stem = input_path.stem

    clean_path = (
        outdir / f"clean_{stem}.xlsx"
    )

    problems_path = (
        outdir / f"problems_{stem}.xlsx"
    )

    _write_with_autofit(
        clean_df,
        clean_path,
    )

    _write_with_autofit(
        problems_df,
        problems_path,
    )

    print(f"Исходных строк: {len(df)}")
    print(
        f"Чистых уникальных записей: "
        f"{len(clean_df)}"
    )
    print(
        f"Убрано дублей (по телефону): "
        f"{duplicates_removed}"
    )
    print(
        f"Проблемных строк: "
        f"{len(problems_df)}"
    )
    print(
        f"Готово:\n"
        f"  {clean_path}\n"
        f"  {problems_path}"
    )


def main():

    parser = argparse.ArgumentParser(
        description="Очистка файла заявок"
    )

    parser.add_argument(
        "input",
        help="Путь к входному файлу .xlsx или .csv",
    )

    parser.add_argument(
        "--outdir",
        default=".",
        help=(
            "Папка для результатов "
            "(по умолчанию текущая)"
        ),
    )

    args = parser.parse_args()

    process(
        Path(args.input),
        Path(args.outdir),
    )


if __name__ == "__main__":
    main()
