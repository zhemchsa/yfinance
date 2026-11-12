import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from src import main


class WorkbookTickerTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.workbook_path = Path(self.temp_dir.name) / 'tickers.xlsx'
        workbook = Workbook()
        workbook.active.title = 'Data Import'
        workbook.save(self.workbook_path)
        workbook.close()

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_reads_only_the_first_30_ticker_slots(self):
        workbook = load_workbook(self.workbook_path)
        sheet = workbook['Data Import']
        for index in range(31):
            sheet.cell(row=23, column=2 + (index * 2)).value = f'T{index + 1}'
        workbook.save(self.workbook_path)
        workbook.close()

        with patch.object(main, 'WORKBOOK_PATH', self.workbook_path):
            self.assertEqual(
                main.get_workbook_tickers(),
                [f'T{index}' for index in range(1, 31)],
            )

    def test_refreshes_30_tickers_and_writes_the_last_output_block(self):
        tickers = [f'T{index}' for index in range(1, 32)]
        fake_stock = type('FakeStock', (), {'info': {'currentPrice': 123.45}})()

        with (
            patch.object(main, 'WORKBOOK_PATH', self.workbook_path),
            patch.object(main.yf, 'Ticker', return_value=fake_stock) as ticker_factory,
        ):
            main.refresh_workbook_data(tickers)

        self.assertEqual(ticker_factory.call_count, 30)

        workbook = load_workbook(self.workbook_path, data_only=False)
        sheet = workbook['Data Import']
        self.assertEqual(sheet.cell(row=23, column=60).value, 'T30')
        self.assertEqual(sheet.cell(row=558, column=1).value, 'T30')
        self.assertEqual(sheet.cell(row=558, column=3).value, 123.45)
        self.assertEqual(sheet.cell(row=558, column=5).value, '=IF(D558="",C558,D558)')
        workbook.close()

    def test_writes_weighted_sector_pe_and_stock_comparison(self):
        companies = SimpleNamespace(
            iterrows=lambda: iter([
                ('A', {'market weight': 0.25}),
                ('B', {'market weight': 0.75}),
                ('C', {'market weight': 0.50}),
            ])
        )
        stocks = {
            'GE': {'sector': 'Industrials', 'trailingPE': 39},
            'A': {'trailingPE': 10},
            'B': {'trailingPE': 30},
            'C': {'trailingPE': -5},
        }

        with (
            patch.object(main, 'WORKBOOK_PATH', self.workbook_path),
            patch.object(
                main.yf,
                'Ticker',
                side_effect=lambda symbol: SimpleNamespace(info=stocks[symbol]),
            ),
            patch.object(
                main.yf,
                'Sector',
                return_value=SimpleNamespace(top_companies=companies),
            ),
        ):
            main.refresh_workbook_data(['GE'])

        workbook = load_workbook(self.workbook_path, data_only=False)
        sheet = workbook['Data Import']
        self.assertEqual(sheet['G6'].value, 'Sector')
        self.assertEqual(sheet['H6'].value, 'Sector P/E (Top 50)')
        self.assertEqual(sheet['I6'].value, 'P/E vs Sector (%)')
        self.assertEqual(sheet['C17'].value, 39)
        self.assertEqual(sheet['G17'].value, 'Industrials')
        self.assertAlmostEqual(sheet['H17'].value, 20)
        self.assertEqual(
            sheet['I17'].value,
            '=IF(AND(ISNUMBER(E17),E17>0,ISNUMBER(H17),H17>0),E17/H17-1,"")',
        )
        workbook.close()


if __name__ == '__main__':
    unittest.main()
