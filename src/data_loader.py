"""
データ読み込みモジュール
Excelファイルから営業リストを読み込む
"""

import pandas as pd
from pathlib import Path


class DataLoader:
    """
    Excelファイルから営業リストを読み込むクラス
    """

    def __init__(self, file_path: str):
        """
        コンストラクタ

        Args:
            file_path (str): Excelファイルのパス
        """
        self.file_path = Path(file_path)

    def load(self) -> pd.DataFrame:
        """
        Excelファイルを読み込む

        Returns:
            pd.DataFrame: 営業リストのデータフレーム

        Raises:
            FileNotFoundError: ファイルが存在しない場合
            ValueError: 必須カラムが存在しない場合
        """
        # ファイルの存在確認
        if not self.file_path.exists():
            raise FileNotFoundError(f"ファイルが見つかりません: {self.file_path}")

        # Excelファイルを読み込み（シート「営業リスト」を指定）
        try:
            df = pd.read_excel(self.file_path, sheet_name='営業リスト')
        except ValueError as e:
            raise ValueError(f"シート「営業リスト」が見つかりません: {e}")

        # 必須カラムの確認
        required_columns = ['優先順位', '氏名', '郵便番号', '住所（標準化）', '都道府県']
        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            raise ValueError(f"必須カラムが存在しません: {missing_columns}")

        # 郵便番号を文字列型に変換（数値型で読み込まれる場合があるため）
        df['郵便番号'] = df['郵便番号'].astype(str)

        return df
