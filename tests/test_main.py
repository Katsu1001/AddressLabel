"""
テストコード
POSSIM ハガキ宛名印刷自動化システムのユニットテスト
"""

import unittest
import pandas as pd
import sys
from pathlib import Path

# srcディレクトリをパスに追加
sys.path.insert(0, str(Path(__file__).parent.parent / 'src'))

from data_cleaner import DataCleaner


class TestDataCleaner(unittest.TestCase):
    """DataCleanerクラスのテスト"""

    def setUp(self):
        """テスト前の準備"""
        # サンプルデータを作成
        self.sample_data = pd.DataFrame({
            '優先順位': [1, 2, 3, 1],
            '氏名': ['木邑敏章', '正司宣彦', '後藤由起子', 'John Doe'],
            '郵便番号': ['2300025', '618-0011', 'ABC123', '1234567'],
            '住所（標準化）': [
                '神奈川県横浜市鶴見区市場大和町8-4',
                '京都府三島郡島本町広瀬３丁目６－７',
                '5425 Buchanan Street Burnaby BC Canada',
                '東京都渋谷区1-1-1'
            ],
            '都道府県': ['神奈川県', '京都府', 'カナダ', '東京都']
        })

    def test_normalize_postal_code(self):
        """郵便番号正規化のテスト"""
        cleaner = DataCleaner(self.sample_data)
        cleaner._normalize_postal_code()

        # 正しい郵便番号形式に変換されているか確認
        self.assertEqual(cleaner.df.loc[0, '郵便番号'], '230-0025')
        self.assertEqual(cleaner.df.loc[1, '郵便番号'], '618-0011')

        # 不正な郵便番号はNaNになっているか確認
        self.assertTrue(pd.isna(cleaner.df.loc[2, '郵便番号']))

        # 7桁の数字のみの郵便番号も正しく変換されているか確認
        self.assertEqual(cleaner.df.loc[3, '郵便番号'], '123-4567')

    def test_normalize_address(self):
        """住所標準化のテスト"""
        cleaner = DataCleaner(self.sample_data)
        cleaner._normalize_address()

        # 住所が正しく整形されているか確認
        self.assertEqual(
            cleaner.df.loc[0, '住所_整形済み'],
            '神奈川県横浜市鶴見区市場大和町8-4'
        )
        self.assertEqual(
            cleaner.df.loc[1, '住所_整形済み'],
            '京都府三島郡島本町広瀬３丁目６－７'
        )

    def test_format_name(self):
        """氏名整形のテスト"""
        cleaner = DataCleaner(self.sample_data)
        cleaner._format_name()

        # 「様」が正しく付与されているか確認
        self.assertEqual(cleaner.df.loc[0, '氏名_整形済み'], '木邑敏章 様')
        self.assertEqual(cleaner.df.loc[1, '氏名_整形済み'], '正司宣彦 様')

    def test_detect_foreign_address(self):
        """国外住所検出のテスト"""
        cleaner = DataCleaner(self.sample_data)
        cleaner._normalize_address()
        cleaner._detect_foreign_address()

        # 国外住所が正しく検出されているか確認
        self.assertFalse(cleaner.df.loc[0, '国外住所フラグ'])
        self.assertFalse(cleaner.df.loc[1, '国外住所フラグ'])
        self.assertTrue(cleaner.df.loc[2, '国外住所フラグ'])

    def test_detect_duplicate_address(self):
        """重複住所検出のテスト"""
        # 重複住所を含むサンプルデータを作成
        duplicate_data = pd.DataFrame({
            '優先順位': [1, 2, 3],
            '氏名': ['田中太郎', '田中花子', '山田太郎'],
            '郵便番号': ['1234567', '1234567', '7654321'],
            '住所（標準化）': [
                '東京都渋谷区1-1-1',
                '東京都渋谷区1-1-1',
                '東京都新宿区2-2-2'
            ],
            '都道府県': ['東京都', '東京都', '東京都']
        })

        cleaner = DataCleaner(duplicate_data)
        cleaner._normalize_address()
        cleaner._detect_duplicate_address()

        # 重複住所が正しく検出されているか確認
        self.assertTrue(cleaner.df.loc[0, '重複住所フラグ'])
        self.assertTrue(cleaner.df.loc[1, '重複住所フラグ'])
        self.assertFalse(cleaner.df.loc[2, '重複住所フラグ'])

    def test_clean_full_process(self):
        """全体のクレンジングプロセスのテスト"""
        cleaner = DataCleaner(self.sample_data)
        df_cleaned = cleaner.clean()

        # 必要なカラムが作成されているか確認
        self.assertIn('郵便番号', df_cleaned.columns)
        self.assertIn('住所_整形済み', df_cleaned.columns)
        self.assertIn('氏名_整形済み', df_cleaned.columns)
        self.assertIn('国外住所フラグ', df_cleaned.columns)
        self.assertIn('重複住所フラグ', df_cleaned.columns)

        # エラーリストが取得できるか確認
        errors = cleaner.get_errors()
        self.assertIsInstance(errors, list)


class TestAddressSplitting(unittest.TestCase):
    """住所分割機能のテスト"""

    def test_split_long_address(self):
        """長い住所が正しく分割されるかテスト"""
        from pdf_generator import PDFGenerator

        # ダミーのデータフレームを作成
        dummy_df = pd.DataFrame({
            '郵便番号': ['123-4567'],
            '住所_整形済み': ['東京都千代田区霞が関1-2-3 中央合同庁舎第5号館10階'],
            '氏名_整形済み': ['山田太郎 様'],
            '国外住所フラグ': [False]
        })

        pdf_gen = PDFGenerator(dummy_df, 'dummy.pdf')

        # 住所を分割
        lines = pdf_gen._split_address('東京都千代田区霞が関1-2-3 中央合同庁舎第5号館10階', max_length=20)

        # 複数行に分割されているか確認
        self.assertGreater(len(lines), 1)

        # 各行が最大文字数以下であることを確認
        for line in lines:
            self.assertLessEqual(len(line), 25)


if __name__ == '__main__':
    unittest.main()
