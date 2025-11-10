"""
CSV生成モジュール
宛名データCSVを生成する（Word差し込み印刷用）
"""

import pandas as pd


class CSVGenerator:
    """
    宛名データCSVを生成するクラス
    """

    def __init__(self, df: pd.DataFrame, output_path: str):
        """
        コンストラクタ

        Args:
            df (pd.DataFrame): クレンジング済みデータフレーム
            output_path (str): 出力CSVファイルパス
        """
        self.df = df
        self.output_path = output_path

    def generate(self):
        """CSVを生成"""
        # 必要なカラムのみを抽出
        df_output = self.df[['郵便番号', '住所_整形済み', '氏名_整形済み']].copy()

        # カラム名を変更（Word差し込み印刷用）
        df_output.columns = ['郵便番号', '住所', '氏名']

        # 欠損データを除外
        df_output = df_output.dropna(subset=['郵便番号'])

        # 空の住所を除外
        df_output = df_output[df_output['住所'] != '']

        # 国外住所を除外
        df_output = df_output[~self.df['国外住所フラグ']]

        # インデックスをリセット
        df_output = df_output.reset_index(drop=True)

        # CSVに出力（Shift_JISエンコーディング）
        # Word差し込み印刷との互換性のため
        try:
            df_output.to_csv(self.output_path, index=False, encoding='shift_jis')
            print(f"✅ CSVを生成しました: {self.output_path}")
        except UnicodeEncodeError:
            # Shift_JISで表現できない文字がある場合はUTF-8で保存
            print("⚠ 警告: Shift_JISで表現できない文字があるため、UTF-8で保存します")
            df_output.to_csv(self.output_path, index=False, encoding='utf-8-sig')
            print(f"✅ CSVを生成しました（UTF-8）: {self.output_path}")

        print(f"   出力件数: {len(df_output)}件")
