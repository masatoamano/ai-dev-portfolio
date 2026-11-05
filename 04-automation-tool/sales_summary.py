"""CSV / Excel sales summary. No network access is used."""
from __future__ import annotations

import argparse
import csv
import io
import re
import sys
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, localcontext
from pathlib import Path
from zipfile import BadZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException

HEADERS = ("取引日", "商品名", "数量", "単価")
# Two decimal places per operand, four for exact products. Bounds also keep
# numeric cells within Excel's 15 significant digit precision.
MAX_VALUE = Decimal("999999.99")
MAX_TOTAL = Decimal("99999999999.9999")
NUMBER = re.compile(r"[0-9]+(?:\.[0-9]{1,2})?\Z")
MONEY_FORMAT = '#,##0.00##'


class DataError(ValueError):
    """An actionable input or output validation failure."""


@dataclass(frozen=True)
class Sale:
    day: date
    product: str
    quantity: Decimal
    price: Decimal

    @property
    def amount(self) -> Decimal:
        with localcontext() as ctx:
            ctx.prec = 40
            return self.quantity * self.price


def parse_day(value: object) -> date:
    if isinstance(value, datetime):
        if value.time() != datetime.min.time():
            raise ValueError("時刻を含めず日付だけを入力してください")
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    if not re.fullmatch(r"[0-9]{4}[-/][0-9]{2}[-/][0-9]{2}", text):
        raise ValueError("YYYY-MM-DD または YYYY/MM/DD で入力してください")
    try:
        return datetime.strptime(text, "%Y-%m-%d" if "-" in text else "%Y/%m/%d").date()
    except ValueError:
        raise ValueError("実在する日付を入力してください") from None


def parse_number(value: object) -> Decimal:
    text = str(value).strip()
    if isinstance(value, bool) or not NUMBER.fullmatch(text):
        raise ValueError("0以上の数値（小数2桁まで、桁区切り・通貨記号なし）を入力してください")
    try:
        number = Decimal(text)
    except InvalidOperation:
        raise ValueError("数値を入力してください") from None
    if not number.is_finite() or number > MAX_VALUE:
        raise ValueError(f"0～{MAX_VALUE}の範囲で入力してください")
    return number


def read_rows(path: Path, sheet: str | None = None, encoding: str = "auto"):
    """Return (location, row) pairs; CSV locations are physical starting lines."""
    if path.suffix.lower() == ".csv":
        if sheet:
            raise DataError("--sheet はExcel入力でのみ指定できます")
        raw = path.read_bytes()
        choices = ("utf-8-sig", "cp932") if encoding == "auto" else (encoding,)
        for codec in choices:
            try:
                decoded = raw.decode(codec, errors="strict")
                break
            except UnicodeDecodeError:
                continue
        else:
            raise DataError("CSVの文字コードを読み取れません。UTF-8 または Shift_JIS/CP932で保存してください")
        reader = csv.reader(io.StringIO(decoded, newline=""), strict=True)
        rows = []
        while True:
            line = reader.line_num + 1
            try:
                row = next(reader)
            except StopIteration:
                break
            except csv.Error as exc:
                raise DataError(f"CSV {line}行目: CSV形式が不正です（{exc}）") from exc
            rows.append((f"CSV {line}行目", row))
        return rows
    if path.suffix.lower() != ".xlsx":
        raise DataError("入力は .csv または .xlsx を指定してください")
    # data_only=False is deliberate: formulas must not use stale cached values.
    book = load_workbook(path, read_only=True, data_only=False)
    try:
        if sheet:
            if sheet not in book.sheetnames:
                raise DataError(f"シート「{sheet}」がありません。候補: {', '.join(book.sheetnames)}")
            ws = book[sheet]
        elif len(book.sheetnames) == 1:
            ws = book.worksheets[0]
        else:
            raise DataError(f"複数シートがあります。--sheet で対象を指定してください。候補: {', '.join(book.sheetnames)}")
        rows = []
        for i, cells in enumerate(ws.iter_rows(), 1):
            for cell in cells:
                if cell.data_type in ("f", "e"):
                    raise DataError(f"シート「{ws.title}」{i}行目 {cell.coordinate}: 数式・Excelエラーは使えません。値に置き換えてください")
            rows.append((f"シート「{ws.title}」{i}行目", [c.value for c in cells]))
        return rows
    finally:
        book.close()


def validate(rows) -> list[Sale]:
    if not rows:
        raise DataError("入力が空です。1行目に見出しを入力してください")
    location, header = rows[0]
    # Excel formatting can extend the used range beyond the four columns.
    header = list(header)
    while header and header[-1] is None:
        header.pop()
    if tuple(str(x).strip() for x in header) != HEADERS:
        raise DataError(f"{location}: 見出しは順に「{','.join(HEADERS)}」の4列にしてください")
    errors, sales = [], []
    for location, original in rows[1:]:
        row = list(original)
        while len(row) > 4 and row[-1] is None:
            row.pop()
        if len(row) != 4:
            errors.append(f"{location}: 列数が{len(row)}です。4列にしてください（空行も削除してください）")
            continue
        parsed = []
        for field, value in zip(HEADERS, row):
            try:
                if value is None or not str(value).strip():
                    raise ValueError("空欄です。値を入力してください")
                if field == "取引日":
                    result = parse_day(value)
                elif field == "商品名":
                    result = str(value).strip()
                    if len(result) > 100 or any(ord(c) < 32 for c in result):
                        raise ValueError("100文字以内で、改行・制御文字を含めず入力してください")
                else:
                    result = parse_number(value)
                parsed.append(result)
            except ValueError as exc:
                errors.append(f"{location} [{field}]: {exc}")
        if len(parsed) == 4:
            sales.append(Sale(*parsed))
    if errors:
        raise DataError("入力エラーのため集計を中止しました。修正後に再実行してください。\n" + "\n".join(errors))
    if not sales:
        raise DataError("明細がありません。2行目以降に売上データを入力してください")
    return sales


def aggregate(sales: list[Sale]):
    monthly, products = defaultdict(Decimal), defaultdict(Decimal)
    with localcontext() as ctx:
        ctx.prec = 40
        total = Decimal(0)
        for sale in sales:
            monthly[sale.day.strftime("%Y-%m")] += sale.amount
            products[sale.product] += sale.amount
            total += sale.amount
            if total > MAX_TOTAL:
                raise DataError(f"全体売上が出力上限{MAX_TOTAL}を超えます。入力を期間などで分割してください")
    return dict(sorted(monthly.items())), dict(sorted(products.items())), total


def write_report(sales: list[Sale], output: Path) -> Decimal:
    if output.suffix.lower() != ".xlsx":
        raise DataError("出力ファイルには .xlsx を指定してください")
    monthly, products, total = aggregate(sales)
    book = Workbook()
    overview = book.active
    overview.title = "全体売上"
    overview.append(["項目", "値"])
    for row in [("全体売上", total), ("取引件数", len(sales)),
                ("開始日", min(s.day for s in sales)), ("終了日", max(s.day for s in sales))]:
        overview.append(row)
    overview["B2"].number_format = MONEY_FORMAT
    for name, header, values in [("月別売上", "月", monthly), ("商品別売上", "商品名", products)]:
        ws = book.create_sheet(name)
        ws.append([header, "売上金額"])
        for key, amount in values.items():
            ws.append([key, amount])
            ws.cell(ws.max_row, 1).data_type = "s"
            ws.cell(ws.max_row, 2).number_format = MONEY_FORMAT
        ws.auto_filter.ref = ws.dimensions
    detail = book.create_sheet("売上明細")
    detail.append([*HEADERS, "売上金額"])
    for sale in sales:
        detail.append([sale.day, sale.product, sale.quantity, sale.price, sale.amount])
        detail.cell(detail.max_row, 2).data_type = "s"  # formula-like product names stay literal
        for col in (3, 4, 5):
            detail.cell(detail.max_row, col).number_format = MONEY_FORMAT
    detail.auto_filter.ref = detail.dimensions
    for ws in book:
        ws.freeze_panes = "A2"
        ws.sheet_view.showGridLines = False
        ws.row_dimensions[1].height = 25
        for cell in ws[1]:
            cell.font = Font(name="Yu Gothic", bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="203864")
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.font = Font(name="Yu Gothic", size=11)
                cell.alignment = Alignment(vertical="center")
                if cell.row % 2 == 0:
                    cell.fill = PatternFill("solid", fgColor="EDF2F8")
                if isinstance(cell.value, date):
                    cell.number_format = "yyyy-mm-dd"
        for col in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col)].width = 32 if col == 2 and ws == detail else 24
    # Serialize before exclusive creation; existing files (including input) are never overwritten.
    buffer = io.BytesIO()
    book.save(buffer)
    book.close()
    try:
        with output.open("xb") as stream:
            try:
                stream.write(buffer.getvalue())
            except BaseException:
                stream.close()
                output.unlink(missing_ok=True)
                raise
    except FileExistsError:
        raise DataError(f"出力先は既に存在します: {output}。別のファイル名を指定してください") from None
    return total


def run(source: Path, output: Path, sheet: str | None = None, encoding: str = "auto"):
    if source.resolve() == output.resolve():
        raise DataError("入力と出力には別のファイルを指定してください")
    sales = validate(read_rows(source, sheet, encoding))
    total = write_report(sales, output)
    return len(sales), total


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="CSV・Excelの売上を検証し、月別・商品別・全体売上をExcelに出力します")
    parser.add_argument("input", type=Path, help="入力 .csv / .xlsx")
    parser.add_argument("-o", "--output", type=Path, help="出力 .xlsx（既存ファイルへの上書き不可）")
    parser.add_argument("--sheet", help="Excelの対象シート名。複数シートでは必須")
    parser.add_argument("--encoding", choices=["auto", "utf-8-sig", "cp932", "shift_jis"], default="auto", help="CSV文字コード（既定: UTF-8→CP932の順で判定）")
    args = parser.parse_args(argv)
    output = args.output or args.input.with_name(args.input.stem + "_集計.xlsx")
    try:
        count, total = run(args.input, output, args.sheet, args.encoding)
    except (BadZipFile, InvalidFileException, KeyError, SyntaxError, EOFError) as exc:
        print(f"エラー: 入力Excelが破損しているか対応形式ではありません。Excelで開いて.xlsxとして保存し直してください（{exc}）", file=sys.stderr)
        return 1
    except (DataError, OSError, ValueError) as exc:
        print(f"エラー: {exc}", file=sys.stderr)
        return 1
    print(f"完了: {count}件 / 全体売上 {total:,.4f}\n出力: {output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
