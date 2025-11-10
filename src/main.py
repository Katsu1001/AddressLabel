"""
メインエントリーポイント
POSSIM ハガキ宛名印刷自動化システム
"""

from pathlib import Path
import sys

from data_loader import DataLoader
from data_cleaner import DataCleaner
from pdf_generator import PDFGenerator
from csv_generator import CSVGenerator
from report_generator import ReportGenerator


def main():
    """
    メイン処理
    """
    print("=" * 70)
    print("POSSIM ハガキ宛名印刷自動化システム")
    print("Postcard Address Automation System")
    print("=" * 70)
    print("")

    # 入力ファイルパス
    input_file = Path("data/input/POSSIM_営業リスト_優先順位付き.xlsx")

    # 出力ファイルパス
    output_pdf = Path("data/output/possim_宛名ラベル.pdf")
    output_csv = Path("data/output/possim_宛名データ.csv")
    output_report = Path("data/output/possim_品質レポート.txt")

    # 出力ディレクトリを作成
    output_pdf.parent.mkdir(parents=True, exist_ok=True)

    try:
        # 1. データ読み込み
        print("【ステップ1】データ読み込み中...")
        loader = DataLoader(str(input_file))
        df = loader.load()
        print(f"✅ {len(df)}件のデータを読み込みました")
        print("")

        # 2. データクレンジング
        print("【ステップ2】データクレンジング中...")
        cleaner = DataCleaner(df)
        df_cleaned = cleaner.clean()
        errors = cleaner.get_errors()
        print(f"✅ データクレンジングが完了しました")
        print("")

        # 3. PDF生成
        print("【ステップ3】PDF生成中...")
        pdf_gen = PDFGenerator(df_cleaned, str(output_pdf))
        pdf_gen.generate()
        print("")

        # 4. CSV生成
        print("【ステップ4】CSV生成中...")
        csv_gen = CSVGenerator(df_cleaned, str(output_csv))
        csv_gen.generate()
        print("")

        # 5. レポート生成
        print("【ステップ5】レポート生成中...")
        report_gen = ReportGenerator(df_cleaned, errors, str(output_report))
        report_gen.generate()
        print("")

        print("=" * 70)
        print("✅ すべての処理が完了しました！")
        print("=" * 70)
        print("")
        print("【出力ファイル】")
        print(f"  - PDF: {output_pdf}")
        print(f"  - CSV: {output_csv}")
        print(f"  - レポート: {output_report}")
        print("")

        return 0

    except FileNotFoundError as e:
        print(f"❌ エラー: {e}")
        print("")
        print("【解決方法】")
        print(f"  1. 入力ファイルを {input_file} に配置してください")
        print(f"  2. ファイル名が正しいか確認してください")
        return 1

    except ValueError as e:
        print(f"❌ エラー: {e}")
        print("")
        print("【解決方法】")
        print(f"  1. Excelファイルのフォーマットを確認してください")
        print(f"  2. 必須カラムが存在するか確認してください")
        return 1

    except Exception as e:
        print(f"❌ 予期しないエラーが発生しました: {e}")
        print("")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
