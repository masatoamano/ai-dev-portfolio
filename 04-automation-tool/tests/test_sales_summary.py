import contextlib
import io
import tempfile
import unittest
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook
from sales_summary import DataError, HEADERS, Sale, main, parse_number, read_rows, run, validate, write_report

ROOT = Path(__file__).resolve().parents[1]


class SalesTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.output = self.folder / 'report.xlsx'

    def csv(self, text, encoding='utf-8'):
        path = self.folder / 'input.csv'
        path.write_text(text, encoding=encoding)
        return path

    def xlsx(self, rows, extra_sheet=False):
        path = self.folder / 'input.xlsx'
        wb = Workbook()
        wb.active.title = '売上'
        for row in rows:
            wb.active.append(row)
        if extra_sheet:
            wb.create_sheet('別シート')
        wb.save(path)
        wb.close()
        return path

    def test_sample_aggregate_and_export(self):
        self.assertEqual(run(ROOT / 'samples/sales.csv', self.output), (7, Decimal('26400.6250')))
        wb = load_workbook(self.output, data_only=True)
        self.addCleanup(wb.close)
        self.assertEqual(wb.sheetnames, ['全体売上', '月別売上', '商品別売上', '売上明細'])
        self.assertEqual(Decimal(str(wb['全体売上']['B2'].value)), Decimal('26400.625'))
        self.assertEqual(list(wb['月別売上'].values), [
            ('月', '売上金額'), ('2026-01', 9800), ('2026-02', 7500), ('2026-03', 9100.625)])
        self.assertEqual(dict(list(wb['商品別売上'].values)[1:]), {
            '架空ブレンドコーヒー': 8000, '架空焼き菓子セット': 10800,
            '架空マグカップ': 7500, '架空量り売り茶葉': 100.625})
        self.assertEqual(sum(Decimal(str(r[4])) for r in list(wb['売上明細'].values)[1:]), Decimal('26400.625'))
        self.assertEqual(wb['売上明細']['A2'].value, datetime(2026, 1, 5))
        self.assertEqual(wb['売上明細'].freeze_panes, 'A2')
        self.assertTrue(wb['売上明細'].auto_filter.ref)
        self.assertEqual(wb['全体売上']['B2'].number_format, '#,##0.00##')

    def test_csv_encodings(self):
        for encoding in ['utf-8', 'utf-8-sig', 'cp932', 'shift_jis']:
            with self.subTest(encoding=encoding):
                source = self.csv('取引日,商品名,数量,単価\n2026/01/01,架空商品,0.1,0.2\n', encoding)
                self.assertEqual(validate(read_rows(source))[0].amount, Decimal('0.02'))

    def test_xlsx_dates_and_numbers(self):
        source = self.xlsx([HEADERS, [date(2026, 2, 1), '架空商品', 1.25, 80.5]])
        self.assertEqual(run(source, self.output), (1, Decimal('100.625')))

    def test_all_validation_errors_and_no_output(self):
        with self.assertRaises(DataError) as caught:
            run(ROOT / 'samples/invalid_sales.csv', self.output)
        for text in ['2行目 [取引日]', '3行目 [商品名]', '3行目 [数量]', '3行目 [単価]',
                     '4行目 [数量]', '4行目 [単価]']:
            self.assertIn(text, str(caught.exception))
        self.assertFalse(self.output.exists())

    def test_bad_numbers(self):
        for value in ['-1', 'NaN', 'Infinity', '1e3', '1,000', '￥100', True, '0.001', '1000000', '１２', '+1']:
            with self.subTest(value=value), self.assertRaises(ValueError):
                parse_number(value)
        self.assertEqual(parse_number('0'), Decimal(0))
        self.assertEqual(parse_number('999999.99'), Decimal('999999.99'))

    def test_bad_structure(self):
        for text in ['', ',,,\n', '取引日,商品名,数量,単価\n', '取引日,商品名,数量,単価\n\n',
                     '取引日,商品名,数量,単価\n2026-01-01,A,1,1,extra\n', '商品名,取引日,数量,単価\n']:
            with self.subTest(text=text), self.assertRaises(DataError):
                validate(read_rows(self.csv(text)))

    def test_csv_physical_lines_and_malformed_quotes(self):
        source = self.csv('取引日,商品名,数量,単価\n2026-01-01,"A\nB",1,2\n2026-02-30,C,1,2\n')
        with self.assertRaises(DataError) as caught:
            validate(read_rows(source))
        self.assertIn('CSV 2行目 [商品名]', str(caught.exception))
        self.assertIn('CSV 4行目 [取引日]', str(caught.exception))
        with self.assertRaisesRegex(DataError, 'CSV 2行目'):
            read_rows(self.csv('取引日,商品名,数量,単価\n"unclosed'))

    def test_date_rules(self):
        for day in ['2026-02-29', '2026-1-01', '2026-01/01', '01/02/2026', 46000, datetime(2026, 1, 1, 12)]:
            with self.subTest(day=day), self.assertRaises(DataError):
                validate([('1行目', HEADERS), ('2行目', [day, '商品', 1, 2])])
        self.assertEqual(validate([('1', HEADERS), ('2', ['2024-02-29', '商品', 0, 2])])[0].amount, 0)

    def test_multiple_sheets_require_selection(self):
        source = self.xlsx([HEADERS, ['2026-01-01', '商品', 1, 2]], extra_sheet=True)
        with self.assertRaisesRegex(DataError, '--sheet'):
            read_rows(source)
        self.assertEqual(len(validate(read_rows(source, '売上'))), 1)
        with self.assertRaisesRegex(DataError, 'ありません'):
            read_rows(source, '不存在')

    def test_excel_formula_and_error_cells(self):
        for value in ['=1+1', '#DIV/0!']:
            with self.subTest(value=value):
                source = self.xlsx([HEADERS, ['2026-01-01', '商品', 1, value]])
                with self.assertRaisesRegex(DataError, '2行目 D2'):
                    run(source, self.output)
                self.assertFalse(self.output.exists())

    def test_formula_like_product_is_written_as_text(self):
        source = self.csv('取引日,商品名,数量,単価\n2026-01-01,=1+1,1,2\n')
        run(source, self.output)
        wb = load_workbook(self.output, data_only=False)
        self.addCleanup(wb.close)
        for sheet, coordinate in [('商品別売上', 'A2'), ('売上明細', 'B2')]:
            self.assertEqual(wb[sheet][coordinate].value, '=1+1')
            self.assertEqual(wb[sheet][coordinate].data_type, 's')

    def test_existing_output_and_same_input_preserved(self):
        source = self.xlsx([HEADERS, ['2026-01-01', '商品', 1, 2]])
        original = source.read_bytes()
        with self.assertRaises(DataError):
            run(source, source)
        self.assertEqual(source.read_bytes(), original)
        self.output.write_bytes(b'existing file')
        with self.assertRaises(DataError):
            run(source, self.output)
        self.assertEqual(self.output.read_bytes(), b'existing file')
        self.assertEqual(source.read_bytes(), original)

    def test_excel_precision_bounds(self):
        sales = [Sale(date(2026, 1, 1), 'A', Decimal('999999.99'), Decimal('999999.99'))]
        with self.assertRaisesRegex(DataError, '出力上限'):
            write_report(sales, self.output)
        self.assertFalse(self.output.exists())
        sales = [Sale(date(2026, 1, 1), 'A', Decimal('0.01'), Decimal('0.01'))]
        self.assertEqual(write_report(sales, self.output), Decimal('0.0001'))
        wb = load_workbook(self.output, data_only=True)
        self.addCleanup(wb.close)
        self.assertEqual(Decimal(str(wb['全体売上']['B2'].value)), Decimal('0.0001'))

    def test_cli_success_failure_and_missing_file(self):
        with contextlib.redirect_stdout(io.StringIO()) as stdout:
            self.assertEqual(main([str(ROOT / 'samples/sales.csv'), '-o', str(self.output)]), 0)
        self.assertIn('26,400.6250', stdout.getvalue())
        for source in [ROOT / 'samples/invalid_sales.csv', self.folder / 'missing.csv']:
            with contextlib.redirect_stderr(io.StringIO()) as stderr:
                self.assertEqual(main([str(source), '-o', str(self.output)]), 1)
            self.assertIn('エラー', stderr.getvalue())

    def test_corrupt_xlsx_encoding_and_options(self):
        source = self.folder / 'broken.xlsx'
        source.write_bytes(b'not a zip')
        with contextlib.redirect_stderr(io.StringIO()) as stderr:
            self.assertEqual(main([str(source), '-o', str(self.output)]), 1)
        self.assertNotIn('Traceback', stderr.getvalue())
        source = self.folder / 'bad.csv'
        source.write_bytes(b'\xff\xff')
        with self.assertRaises(DataError):
            read_rows(source, encoding='utf-8-sig')
        with self.assertRaises(DataError):
            read_rows(source, sheet='売上')
        with self.assertRaises(DataError):
            read_rows(self.folder / 'input.xls')
        with self.assertRaises(DataError):
            write_report([], self.folder / 'report.csv')

    def test_excel_trailing_formatted_columns(self):
        source = self.xlsx([list(HEADERS) + [None], ['2026-01-01', '商品', 1, 2, None]])
        self.assertEqual(len(validate(read_rows(source))), 1)

    def test_write_permission_error_leaves_no_output(self):
        sales = [Sale(date(2026, 1, 1), '商品', Decimal(1), Decimal(2))]
        with patch.object(Path, 'open', side_effect=PermissionError('書き込み不可')):
            with self.assertRaises(PermissionError):
                write_report(sales, self.output)
        self.assertFalse(self.output.exists())


if __name__ == '__main__':
    unittest.main()
