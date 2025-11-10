"""
レポート生成モジュール
データ品質レポートを生成する
"""

import pandas as pd
from datetime import datetime


class ReportGenerator:
    """
    データ品質レポートを生成するクラス
    """

    def __init__(self, df: pd.DataFrame, errors: list, output_path: str):
        """
        コンストラクタ

        Args:
            df (pd.DataFrame): クレンジング済みデータフレーム
            errors (list): エラーリスト
            output_path (str): 出力テキストファイルパス
        """
        self.df = df
        self.errors = errors
        self.output_path = output_path

    def generate(self):
        """レポートを生成"""
        report_lines = []

        # ヘッダー
        report_lines.append("=" * 70)
        report_lines.append("POSSIM ハガキ宛名印刷 データ品質レポート")
        report_lines.append("=" * 70)
        report_lines.append("")
        report_lines.append(f"生成日時: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}")
        report_lines.append("")

        # 総件数
        total_count = len(self.df)
        report_lines.append(f"【総件数】: {total_count}件")
        report_lines.append("")

        # 有効件数（郵便番号あり、住所あり、国外住所でない）
        valid_count = len(self.df[
            ~self.df['郵便番号'].isna() &
            (self.df['住所_整形済み'] != '') &
            ~self.df['国外住所フラグ']
        ])
        report_lines.append(f"【有効件数】: {valid_count}件")
        report_lines.append("")

        # 除外件数
        excluded_count = total_count - valid_count
        report_lines.append(f"【除外件数】: {excluded_count}件")
        report_lines.append("")

        # エラー内容
        if self.errors:
            report_lines.append("【エラー・警告内容】:")
            for error in self.errors:
                report_lines.append(f"  - {error}")
            report_lines.append("")

        # 優先度別の集計
        if '優先順位' in self.df.columns:
            report_lines.append("【優先度別の件数】:")
            priority_counts = self.df['優先順位'].value_counts().sort_index()
            for priority, count in priority_counts.items():
                # 有効件数も表示
                valid_priority_count = len(self.df[
                    (self.df['優先順位'] == priority) &
                    ~self.df['郵便番号'].isna() &
                    (self.df['住所_整形済み'] != '') &
                    ~self.df['国外住所フラグ']
                ])
                report_lines.append(f"  優先順位{priority}: {count}件 （有効: {valid_priority_count}件）")
            report_lines.append("")

        # データ品質サマリー
        report_lines.append("【データ品質サマリー】:")

        # 郵便番号欠損率
        postal_missing_rate = (self.df['郵便番号'].isna().sum() / total_count) * 100
        report_lines.append(f"  郵便番号欠損率: {postal_missing_rate:.1f}%")

        # 国外住所率
        foreign_rate = (self.df['国外住所フラグ'].sum() / total_count) * 100
        report_lines.append(f"  国外住所率: {foreign_rate:.1f}%")

        # 重複住所率
        duplicate_rate = (self.df['重複住所フラグ'].sum() / total_count) * 100
        report_lines.append(f"  重複住所率: {duplicate_rate:.1f}%")

        report_lines.append("")

        # 処理結果サマリー
        report_lines.append("【処理結果サマリー】:")
        report_lines.append(f"  宛名ラベル生成数: {valid_count}枚")
        report_lines.append(f"  成功率: {(valid_count / total_count) * 100:.1f}%")
        report_lines.append("")

        # フッター
        report_lines.append("=" * 70)
        report_lines.append("処理完了")
        report_lines.append("=" * 70)

        # レポートを出力
        report_text = "\n".join(report_lines)
        with open(self.output_path, 'w', encoding='utf-8') as f:
            f.write(report_text)

        print(f"✅ レポートを生成しました: {self.output_path}")
        print("")
        print(report_text)
