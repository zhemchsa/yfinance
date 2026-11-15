from contextlib import redirect_stdout
from io import StringIO
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

import pandas as pd
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
        history = pd.DataFrame({'Close': range(1, 201)})
        fake_stock = SimpleNamespace(
            info={'currentPrice': 123.45},
            history=Mock(return_value=history),
        )
        output = StringIO()

        with (
            redirect_stdout(output),
            patch.object(main, 'WORKBOOK_PATH', self.workbook_path),
            patch.object(main.yf, 'Ticker', return_value=fake_stock) as ticker_factory,
        ):
            main.refresh_workbook_data(tickers)

        self.assertEqual(ticker_factory.call_count, 30)
        self.assertIn('Technical Buy Analysis:', output.getvalue())
        self.assertIn('50-day MA', output.getvalue())
        self.assertIn('200-day MA', output.getvalue())
        self.assertIn('Buy price', output.getvalue())
        self.assertIn('Analyst target', output.getvalue())
        self.assertIn('Recommendation', output.getvalue())
        self.assertIn('T30', output.getvalue())
        self.assertIn('$123.45', output.getvalue())
        self.assertIn('Insufficient data', output.getvalue())

        workbook = load_workbook(self.workbook_path, data_only=False)
        sheet = workbook['Data Import']
        self.assertEqual(sheet.cell(row=23, column=60).value, 'T30')
        self.assertEqual(sheet['J23'].value, 'T5')
        self.assertEqual(sheet['L23'].value, 'T6')
        self.assertEqual(sheet.cell(row=558, column=1).value, 'T30')
        self.assertEqual(sheet.cell(row=558, column=3).value, 123.45)
        self.assertEqual(sheet.cell(row=558, column=5).value, '=IF(D558="",C558,D558)')
        self.assertEqual(sheet['J6'].value, '50-day MA')
        self.assertEqual(sheet['K6'].value, '200-day MA')
        self.assertEqual(sheet['L6'].value, 'Buy price')
        self.assertEqual(sheet['M6'].value, 'Buy recommendation')
        self.assertAlmostEqual(sheet['J568'].value, 175.5)
        self.assertAlmostEqual(sheet['K568'].value, 100.5)
        self.assertEqual(sheet['M568'].value, 'Insufficient data')
        workbook.close()

    def test_screen_table_shows_sector_pe_after_stock_pe(self):
        output = StringIO()
        stock = SimpleNamespace(
            info={
                'trailingPE': 12,
                'priceToBook': 1,
                'pegRatio': 0.5,
                'sector': 'Technology',
            }
        )

        with (
            redirect_stdout(output),
            patch.object(main.yf, 'Ticker', return_value=stock),
            patch.object(main, 'get_sector_average_pe', return_value=28.5) as sector_pe_lookup,
        ):
            main.screen_value_candidates(['TEST'])

        lines = output.getvalue().splitlines()
        self.assertLess(lines[0].index('P/E'), lines[0].index('Sector P/E'))
        self.assertLess(lines[0].index('Sector P/E'), lines[0].index('P/B'))
        self.assertEqual(
            lines[2].split(),
            ['TEST', '12.00', '28.50x', '1.00', '0.50', 'MATCH'],
        )
        sector_pe_lookup.assert_called_once_with('Technology')

    def test_screen_table_shows_sector_pe_even_when_no_ticker_matches(self):
        output = StringIO()
        stock = SimpleNamespace(
            info={
                'trailingPE': 20,
                'priceToBook': 2,
                'pegRatio': 1.5,
                'sector': 'Technology',
            }
        )

        with (
            redirect_stdout(output),
            patch.object(main.yf, 'Ticker', return_value=stock),
            patch.object(main, 'get_sector_average_pe', return_value=28.5),
        ):
            main.screen_value_candidates(['TEST'])

        lines = output.getvalue().splitlines()
        self.assertEqual(
            lines[2].split(),
            ['TEST', '20.00', '28.50x', '2.00', '1.50', 'NO', 'MATCH'],
        )
        self.assertEqual(lines[3], 'No tickers matched the screen.')

    def test_moving_averages_require_enough_valid_closes(self):
        self.assertEqual(
            main.calculate_moving_averages(pd.DataFrame({'Close': range(49)})),
            (None, None),
        )
        fifty_day, two_hundred_day = main.calculate_moving_averages(
            pd.DataFrame({'Close': range(1, 51)})
        )
        self.assertEqual(fifty_day, 25.5)
        self.assertIsNone(two_hundred_day)

    def test_writes_weighted_sector_pe_and_stock_comparison(self):
        companies = SimpleNamespace(
            iterrows=lambda: iter([
                ('A', {'market weight': 0.25}),
                ('B', {'market weight': 0.75}),
                ('C', {'market weight': 0.50}),
            ])
        )
        stocks = {
            'GE': {
                'sector': 'Industrials',
                'currentPrice': 150,
                'trailingPE': 39,
                'targetMeanPrice': 120,
            },
            'A': {'trailingPE': 10},
            'B': {'trailingPE': 30},
            'C': {'trailingPE': -5},
        }

        with (
            patch.object(main, 'WORKBOOK_PATH', self.workbook_path),
            patch.object(
                main.yf,
                'Ticker',
                side_effect=lambda symbol: SimpleNamespace(
                    info=stocks[symbol],
                    history=Mock(return_value=pd.DataFrame({'Close': [100] * 200})),
                ),
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
        self.assertEqual(sheet['N6'].value, 'Analyst target')
        self.assertEqual(sheet['I6'].value, 'P/E vs Sector (%)')
        self.assertEqual(sheet['C17'].value, 39)
        self.assertEqual(sheet['G17'].value, 'Industrials')
        self.assertAlmostEqual(sheet['H17'].value, 20)
        self.assertEqual(
            sheet['I17'].value,
            '=IF(AND(ISNUMBER(E17),E17>0,ISNUMBER(H17),H17>0),E17/H17-1,"")',
        )
        self.assertEqual(sheet['J17'].value, 100)
        self.assertEqual(sheet['K17'].value, 100)
        self.assertEqual(sheet['L17'].value, 100)
        self.assertEqual(sheet['M17'].value, 'Wait for buy price')
        self.assertEqual(sheet['N17'].value, 120)
        workbook.close()

    def test_buy_recommendation_uses_valuation_and_trend_filters(self):
        buy_price, recommendation = main.get_buy_price_and_recommendation(
            current_price=90,
            analyst_target=120,
            fifty_day_average=95,
            two_hundred_day_average=85,
        )
        self.assertEqual(buy_price, 95)
        self.assertEqual(recommendation, 'Consider buy')

        _, recommendation = main.get_buy_price_and_recommendation(
            current_price=80,
            analyst_target=120,
            fifty_day_average=95,
            two_hundred_day_average=85,
        )
        self.assertEqual(recommendation, 'Wait - below 200-day MA')

        buy_price, recommendation = main.get_buy_price_and_recommendation(
            current_price=91,
            analyst_target=90,
            fifty_day_average=95,
            two_hundred_day_average=85,
        )
        self.assertEqual(buy_price, 90)
        self.assertEqual(recommendation, 'Wait for buy price')

        self.assertEqual(
            main.get_buy_price_and_recommendation(90, None, 95, 85),
            (None, 'Insufficient data'),
        )


if __name__ == '__main__':
    unittest.main()
