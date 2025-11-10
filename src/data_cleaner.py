"""
データクレンジングモジュール
営業リストデータをクレンジングする
"""

import pandas as pd
import re


class DataCleaner:
    """
    営業リストデータをクレンジングするクラス
    """

    def __init__(self, df: pd.DataFrame):
        """
        コンストラクタ

        Args:
            df (pd.DataFrame): 生データのデータフレーム
        """
        self.df = df.copy()
        self.errors = []  # エラーリスト

    def clean(self) -> pd.DataFrame:
        """
        データをクレンジングする

        Returns:
            pd.DataFrame: クレンジング済みのデータフレーム
        """
        # 1. 郵便番号の正規化
        self._normalize_postal_code()

        # 2. 住所の標準化
        self._normalize_address()

        # 3. 氏名の整形
        self._format_name()

        # 4. 国外住所の検出
        self._detect_foreign_address()

        # 5. 重複住所の検出
        self._detect_duplicate_address()

        return self.df

    def _normalize_postal_code(self):
        """郵便番号を XXX-XXXX 形式に正規化"""
        def format_postal_code(code):
            # NaNや空文字列の場合はNoneを返す
            if pd.isna(code) or str(code).strip() == '' or str(code).lower() == 'nan':
                return None

            # 数字のみを抽出
            digits = re.sub(r'\D', '', str(code))

            # 7桁でない場合はNaNを返す
            if len(digits) != 7:
                return None

            # XXX-XXXX形式に変換
            return f"{digits[:3]}-{digits[3:]}"

        self.df['郵便番号'] = self.df['郵便番号'].apply(format_postal_code)

        # 郵便番号欠損の件数をカウント
        missing_count = self.df['郵便番号'].isna().sum()
        if missing_count > 0:
            self.errors.append(f"郵便番号欠損: {missing_count}件")

    def _normalize_address(self):
        """住所を標準化"""
        def clean_address(row):
            address = str(row['住所（標準化）'])
            prefecture = str(row['都道府県'])

            # NaNの場合は空文字列を返す
            if address == 'nan':
                return ''

            # 都道府県名の重複を削除
            # 例: 「神奈川県神奈川県横浜市」→「神奈川県横浜市」
            if prefecture != 'nan' and address.count(prefecture) > 1:
                address = address.replace(prefecture, '', 1)

            # 全角スペースを半角に統一
            address = address.replace('　', ' ')

            return address

        self.df['住所_整形済み'] = self.df.apply(clean_address, axis=1)

    def _format_name(self):
        """氏名に「様」を付与"""
        def format_name(name):
            name = str(name).strip()

            # NaNの場合は空文字列を返す
            if name == 'nan':
                return ''

            # 既に「様」が付いている場合はそのまま
            if name.endswith('様'):
                return name

            # 全角スペースを削除
            name = name.replace(' ', '').replace('　', '')

            # 「様」を付与
            return f"{name} 様"

        self.df['氏名_整形済み'] = self.df['氏名'].apply(format_name)

    def _detect_foreign_address(self):
        """国外住所を検出"""
        def is_foreign(address):
            # NaNや空文字列の場合はFalse
            if pd.isna(address) or str(address).strip() == '':
                return False

            address = str(address)

            # ローマ字（A-Z, a-z）の割合を計算
            alphabet_count = len(re.findall(r'[A-Za-z]', address))
            total_count = len(address)

            if total_count == 0:
                return False

            # ローマ字が50%以上なら国外と判定
            return (alphabet_count / total_count) > 0.5

        self.df['国外住所フラグ'] = self.df['住所_整形済み'].apply(is_foreign)

        # 国外住所の件数をカウント
        foreign_count = self.df['国外住所フラグ'].sum()
        if foreign_count > 0:
            self.errors.append(f"国外住所検出: {foreign_count}件")

    def _detect_duplicate_address(self):
        """重複住所を検出"""
        # 空の住所は除外してカウント
        non_empty_df = self.df[self.df['住所_整形済み'] != '']

        # 住所ごとに件数をカウント
        address_counts = non_empty_df.groupby('住所_整形済み').size()

        # 重複している住所（2件以上）を抽出
        duplicate_addresses = address_counts[address_counts > 1].index.tolist()

        # 重複フラグを付与
        self.df['重複住所フラグ'] = self.df['住所_整形済み'].isin(duplicate_addresses)

        # 重複住所の件数をカウント
        duplicate_count = self.df['重複住所フラグ'].sum()
        if duplicate_count > 0:
            self.errors.append(f"重複住所検出: {duplicate_count}件")

    def get_errors(self) -> list:
        """エラーリストを取得"""
        return self.errors
