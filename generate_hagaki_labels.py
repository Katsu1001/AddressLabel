#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
はがき宛名印刷自動化システム - AddrLabel Pro

このモジュールは、Excelファイルから顧客情報を読み込み、
日本郵便標準のはがき宛名PDFを自動生成するシステムです。

主な機能:
- Excelファイルからの顧客データ読み込み
- 日本郵便標準レイアウトでのPDF生成（100mm × 148mm、縦書き）
- 欠損データの検出とログ出力
- 進捗表示機能
- テストモード対応（10件のみで動作確認）

作成者: AddrLabel Pro Development Team
最終更新: 2025-11-06
"""

import pandas as pd
import logging
import sys
import os
import argparse
from datetime import datetime
from pathlib import Path
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont


# ==================== 定数定義セクション ====================

# ファイルパスの設定
DEFAULT_EXCEL_FILE = "POSSIM_営業リスト_優先順位付き.xlsx"
DEFAULT_SHEET_NAME = "営業リスト"
DEFAULT_OUTPUT_PDF = "output/hagaki_labels.pdf"

# Excelファイルの列名
COLUMN_NAME = "氏名"
COLUMN_POSTAL_CODE = "郵便番号"
COLUMN_ADDRESS = "住所"
COLUMN_PREFECTURE = "都道府県"

# PDFレイアウト設定（単位: mm）
# 日本郵便標準はがきサイズ: 100mm × 148mm
HAGAKI_WIDTH = 100 * mm   # 横幅（ポイント単位に変換: 1mm = 2.834645669291339ポイント）
HAGAKI_HEIGHT = 148 * mm  # 高さ（ポイント単位に変換）

# 各要素の配置座標（mmからポイント単位へ変換済み）
POSTAL_CODE_X = 10 * mm   # 郵便番号のX座標（左から10mm）
POSTAL_CODE_Y = 130 * mm  # 郵便番号のY座標（下から130mm）

ADDRESS_X = 80 * mm       # 住所のX座標（右寄せ、左から80mm）
ADDRESS_Y = 120 * mm      # 住所のY座標（下から120mm）

NAME_X = 35 * mm          # 氏名のX座標（中央寄せ、左から35mm）
NAME_Y = 50 * mm          # 氏名のY座標（下から50mm）

# フォント設定
FONT_NAME = "HeiseiMin-W3"  # 日本語フォント（明朝体）
FONT_SIZE_POSTAL = 14       # 郵便番号のフォントサイズ（ポイント）
FONT_SIZE_ADDRESS = 12      # 住所のフォントサイズ（ポイント）
FONT_SIZE_NAME = 16         # 氏名のフォントサイズ（ポイント）

# 縦書き設定
VERTICAL_CHAR_SPACING = 1.2  # 縦書き時の文字間隔（フォントサイズに対する倍率）

# 進捗表示の間隔
PROGRESS_INTERVAL = 100  # 100件ごとに進捗を表示


# ==================== ユーティリティ関数群 ====================

def setup_logging():
    """
    ログファイルの設定を行う関数

    処理内容:
    1. logs/ ディレクトリを作成（存在しない場合）
    2. ログファイル名を生成（形式: hagaki_generation_YYYYMMDD_HHMMSS.log）
    3. ロギング設定（コンソールとファイルの両方に出力）

    戻り値:
        str: 作成されたログファイルのパス
    """
    # ログディレクトリの作成
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)

    # ログファイル名の生成（タイムスタンプ付き）
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"hagaki_generation_{timestamp}.log"

    # ロギング設定
    # レベル: INFO以上（DEBUG, INFO, WARNING, ERROR, CRITICAL）
    # フォーマット: 時刻 - レベル - メッセージ
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),  # ファイルに出力
            logging.StreamHandler(sys.stdout)                  # コンソールに出力
        ]
    )

    logging.info(f"ログファイルを作成しました: {log_file}")
    return str(log_file)


def load_excel_data(file_path, sheet_name):
    """
    Excelファイルを読み込む関数

    引数:
        file_path (str): Excelファイルのパス
        sheet_name (str): 読み込むシート名

    戻り値:
        pandas.DataFrame: 読み込んだデータフレーム

    例外:
        FileNotFoundError: ファイルが存在しない場合
        ValueError: シートが存在しない場合
    """
    logging.info(f"Excelファイルを読み込んでいます: {file_path}")

    # ファイルの存在確認
    if not os.path.exists(file_path):
        error_msg = f"Excelファイルが見つかりません: {file_path}"
        logging.error(error_msg)
        logging.error(f"対処: ファイルパスを確認してください。カレントディレクトリ: {os.getcwd()}")
        raise FileNotFoundError(error_msg)

    try:
        # Excelファイルの読み込み（openpyxlエンジンを使用）
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        logging.info(f"データを読み込みました: {len(df)}行 × {len(df.columns)}列")
        logging.info(f"列名: {list(df.columns)}")
        return df

    except ValueError as e:
        error_msg = f"シート '{sheet_name}' が見つかりません"
        logging.error(f"{error_msg}: {e}")
        logging.error(f"対処: シート名を確認してください。")
        raise ValueError(error_msg) from e

    except Exception as e:
        error_msg = f"Excelファイルの読み込みに失敗しました"
        logging.error(f"{error_msg}: {e}")
        logging.error(f"対処: ファイルが破損していないか、他のプログラムで開いていないか確認してください。")
        raise


def validate_data(df):
    """
    読み込んだデータの検証を行う関数

    引数:
        df (pandas.DataFrame): 検証するデータフレーム

    例外:
        KeyError: 必要な列が存在しない場合
    """
    logging.info("データの検証を開始します")

    # 必要な列の存在確認
    required_columns = [COLUMN_NAME, COLUMN_POSTAL_CODE, COLUMN_ADDRESS, COLUMN_PREFECTURE]
    missing_columns = [col for col in required_columns if col not in df.columns]

    if missing_columns:
        error_msg = f"必要な列が存在しません: {missing_columns}"
        logging.error(error_msg)
        logging.error(f"対処: Excelファイルに以下の列が存在することを確認してください: {required_columns}")
        raise KeyError(error_msg)

    logging.info("データ検証完了: すべての必要な列が存在します")


def detect_missing_data(df):
    """
    欠損データを検出する関数

    引数:
        df (pandas.DataFrame): 検証するデータフレーム

    戻り値:
        dict: 各列の欠損データ情報
            {
                '列名': {
                    'count': 欠損数,
                    'percentage': 欠損率（%）,
                    'indices': 欠損行のインデックスリスト
                }
            }
    """
    logging.info("欠損データを検出しています")

    missing_info = {}
    total_rows = len(df)

    for column in [COLUMN_NAME, COLUMN_POSTAL_CODE, COLUMN_ADDRESS, COLUMN_PREFECTURE]:
        # 欠損データの検出（NaN、None、空文字を含む）
        missing_mask = df[column].isna() | (df[column].astype(str).str.strip() == '')
        missing_count = missing_mask.sum()
        missing_percentage = (missing_count / total_rows * 100) if total_rows > 0 else 0
        missing_indices = df[missing_mask].index.tolist()

        if missing_count > 0:
            missing_info[column] = {
                'count': missing_count,
                'percentage': missing_percentage,
                'indices': missing_indices
            }
            logging.warning(
                f"列 '{column}': {missing_count}件の欠損データ "
                f"({missing_percentage:.1f}%)"
            )

    return missing_info


def log_missing_data(missing_info):
    """
    欠損データの詳細をログに出力する関数

    引数:
        missing_info (dict): detect_missing_data() の戻り値
    """
    if not missing_info:
        logging.info("欠損データは検出されませんでした")
        return

    logging.warning("=" * 60)
    logging.warning("欠損データ検出レポート")
    logging.warning("=" * 60)

    for column, info in missing_info.items():
        logging.warning(f"\n列名: {column}")
        logging.warning(f"  欠損件数: {info['count']}件")
        logging.warning(f"  欠損率: {info['percentage']:.2f}%")

        # 欠損行のインデックスを表示（最初の10件のみ）
        indices_display = info['indices'][:10]
        if len(info['indices']) > 10:
            indices_display_str = f"{indices_display} ... (他 {len(info['indices']) - 10}件)"
        else:
            indices_display_str = str(indices_display)
        logging.warning(f"  欠損行インデックス: {indices_display_str}")

    logging.warning("=" * 60)


# ==================== PDF生成関数群 ====================

def setup_pdf_canvas(output_file):
    """
    PDF生成用のキャンバスを初期設定する関数

    引数:
        output_file (str): 出力PDFファイルのパス

    戻り値:
        reportlab.pdfgen.canvas.Canvas: 設定済みのキャンバスオブジェクト
    """
    logging.info(f"PDFキャンバスを初期化しています: {output_file}")

    # 出力ディレクトリの作成
    output_dir = Path(output_file).parent
    output_dir.mkdir(exist_ok=True)

    # キャンバスの作成（はがきサイズ: 100mm × 148mm）
    c = canvas.Canvas(output_file, pagesize=(HAGAKI_WIDTH, HAGAKI_HEIGHT))

    # 日本語フォントの登録
    # reportlabには標準で日本語フォントが含まれているため、
    # UnicodeCIDFont を使用して日本語フォントを登録
    global FONT_NAME

    try:
        pdfmetrics.registerFont(UnicodeCIDFont(FONT_NAME))
        logging.info(f"日本語フォントを登録しました: {FONT_NAME}")
    except Exception as e:
        logging.warning(f"フォント '{FONT_NAME}' の登録に失敗しました: {e}")
        logging.warning("代替フォントを使用します")
        # 代替フォント（ゴシック体）を試す
        try:
            pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
            FONT_NAME = 'HeiseiKakuGo-W5'
            logging.info(f"代替フォントを登録しました: {FONT_NAME}")
        except Exception as e2:
            logging.error(f"代替フォントの登録にも失敗しました: {e2}")
            raise

    return c


def draw_postal_code(c, postal_code, x, y):
    """
    郵便番号を描画する関数

    引数:
        c (canvas.Canvas): PDFキャンバス
        postal_code (str): 郵便番号（例: "123-4567"）
        x (float): X座標（ポイント単位）
        y (float): Y座標（ポイント単位）
    """
    # 欠損データの処理
    if pd.isna(postal_code) or str(postal_code).strip() == '':
        postal_code = "郵便番号未記入"
    else:
        postal_code = str(postal_code).strip()

    # フォント設定
    c.setFont(FONT_NAME, FONT_SIZE_POSTAL)

    # 郵便番号の描画（横書き）
    display_text = f"〒 {postal_code}"
    c.drawString(x, y, display_text)


def draw_address_vertical(c, address, prefecture, x, y):
    """
    住所を縦書きで描画する関数

    引数:
        c (canvas.Canvas): PDFキャンバス
        address (str): 住所
        prefecture (str): 都道府県
        x (float): 開始X座標（ポイント単位、右寄せ）
        y (float): 開始Y座標（ポイント単位、上から下へ）
    """
    # 欠損データの処理
    if pd.isna(prefecture) or str(prefecture).strip() == '':
        prefecture = ""
    else:
        prefecture = str(prefecture).strip()

    if pd.isna(address) or str(address).strip() == '':
        address = "住所未記入"
    else:
        address = str(address).strip()

    # 完全な住所を作成
    full_address = f"{prefecture}{address}"

    # フォント設定
    c.setFont(FONT_NAME, FONT_SIZE_ADDRESS)

    # 縦書き描画の準備
    # reportlabで縦書きを実現するために、キャンバスを回転させる
    c.saveState()  # 現在の状態を保存

    # 縦書きの文字間隔
    char_height = FONT_SIZE_ADDRESS * VERTICAL_CHAR_SPACING

    # 1文字ずつ描画（上から下へ）
    current_y = y
    for char in full_address:
        # 文字を描画
        c.drawString(x, current_y, char)
        current_y -= char_height  # 下に移動

        # ページからはみ出す場合は改行（左に移動）
        if current_y < 10 * mm:
            x -= FONT_SIZE_ADDRESS * 1.5  # 左に移動
            current_y = y  # Y座標をリセット

    c.restoreState()  # 状態を復元


def draw_name_vertical(c, name, x, y):
    """
    氏名を縦書きで描画する関数

    引数:
        c (canvas.Canvas): PDFキャンバス
        name (str): 氏名
        x (float): 開始X座標（ポイント単位、中央寄せ）
        y (float): 開始Y座標（ポイント単位、上から下へ）
    """
    # 欠損データの処理（氏名が欠損している場合は処理を継続）
    if pd.isna(name) or str(name).strip() == '':
        logging.warning("氏名が欠損しています。'氏名未記入'と表示します。")
        name = "氏名未記入"
    else:
        name = str(name).strip()

    # フォント設定
    c.setFont(FONT_NAME, FONT_SIZE_NAME)

    # 縦書き描画の準備
    c.saveState()

    # 縦書きの文字間隔
    char_height = FONT_SIZE_NAME * VERTICAL_CHAR_SPACING

    # 「様」を追加
    full_name = f"{name} 様"

    # 1文字ずつ描画（上から下へ）
    current_y = y
    for char in full_name:
        c.drawString(x, current_y, char)
        current_y -= char_height

    c.restoreState()


def generate_single_hagaki(c, row_data, index):
    """
    1件分のはがきを生成する関数

    引数:
        c (canvas.Canvas): PDFキャンバス
        row_data (pandas.Series): 1行分のデータ
        index (int): 行インデックス（エラーログ用）

    戻り値:
        bool: 成功した場合True、失敗した場合False
    """
    try:
        # データの取得
        name = row_data.get(COLUMN_NAME)
        postal_code = row_data.get(COLUMN_POSTAL_CODE)
        address = row_data.get(COLUMN_ADDRESS)
        prefecture = row_data.get(COLUMN_PREFECTURE)

        # 郵便番号の描画
        draw_postal_code(c, postal_code, POSTAL_CODE_X, POSTAL_CODE_Y)

        # 住所の描画（縦書き）
        draw_address_vertical(c, address, prefecture, ADDRESS_X, ADDRESS_Y)

        # 氏名の描画（縦書き）
        draw_name_vertical(c, name, NAME_X, NAME_Y)

        # ページを確定
        c.showPage()

        return True

    except Exception as e:
        logging.error(f"はがき生成エラー（行 {index}）: {e}")
        logging.error(f"データ: 氏名={row_data.get(COLUMN_NAME)}, "
                      f"郵便番号={row_data.get(COLUMN_POSTAL_CODE)}, "
                      f"住所={row_data.get(COLUMN_ADDRESS)}")
        return False


# ==================== メイン処理関数 ====================

def main(excel_file, output_pdf, sheet_name, test_mode=False):
    """
    はがき宛名PDF生成のメイン処理関数

    引数:
        excel_file (str): Excelファイルのパス
        output_pdf (str): 出力PDFファイルのパス
        sheet_name (str): シート名
        test_mode (bool): テストモード（10件のみ処理）

    戻り値:
        bool: 処理が成功した場合True、失敗した場合False
    """
    try:
        # ロギング設定
        log_file = setup_logging()

        logging.info("=" * 80)
        logging.info("はがき宛名PDF自動生成システム - AddrLabel Pro")
        logging.info("=" * 80)
        logging.info(f"実行モード: {'テストモード（10件）' if test_mode else '本番モード（全件）'}")
        logging.info(f"入力ファイル: {excel_file}")
        logging.info(f"出力ファイル: {output_pdf}")
        logging.info(f"シート名: {sheet_name}")
        logging.info(f"ログファイル: {log_file}")
        logging.info("=" * 80)

        # Step 1: Excelファイルの読み込み
        df = load_excel_data(excel_file, sheet_name)

        # Step 2: データの検証
        validate_data(df)

        # Step 3: 欠損データの検出
        missing_info = detect_missing_data(df)
        log_missing_data(missing_info)

        # テストモードの場合は10件のみ処理
        if test_mode:
            df = df.head(10)
            logging.info(f"テストモード: 最初の10件のみを処理します")

        total_count = len(df)
        logging.info(f"処理を開始します。総件数: {total_count}件")

        # Step 4: PDFキャンバスの初期化
        c = setup_pdf_canvas(output_pdf)

        # Step 5: 各行を処理してはがきを生成
        success_count = 0
        error_count = 0

        for index, row in df.iterrows():
            # はがきを生成
            if generate_single_hagaki(c, row, index):
                success_count += 1
            else:
                error_count += 1

            # 進捗表示（100件ごと）
            if (success_count + error_count) % PROGRESS_INTERVAL == 0:
                processed = success_count + error_count
                percentage = (processed / total_count * 100)
                logging.info(f"処理中... [{processed}/{total_count}件] ({percentage:.1f}%)")

        # Step 6: PDFファイルの保存
        logging.info("PDFファイルを保存しています...")
        c.save()

        # Step 7: 処理結果のサマリー
        logging.info("=" * 80)
        logging.info("処理完了")
        logging.info(f"成功: {success_count}件")
        logging.info(f"エラー: {error_count}件")
        logging.info(f"合計: {total_count}件")
        logging.info(f"出力ファイル: {output_pdf}")

        # ファイルサイズを確認
        file_size = os.path.getsize(output_pdf)
        file_size_mb = file_size / (1024 * 1024)
        logging.info(f"ファイルサイズ: {file_size_mb:.2f} MB")

        logging.info("=" * 80)

        # 欠損データがあった場合の警告
        if missing_info:
            print("\n⚠️  警告: 欠損データを検出しました。詳細はログを参照してください。")
            print(f"   ログファイル: {log_file}")

        print(f"\n✅ 処理完了! 出力ファイル: {output_pdf}")

        return True

    except Exception as e:
        logging.error(f"処理中にエラーが発生しました: {e}")
        import traceback
        logging.error(f"スタックトレース:\n{traceback.format_exc()}")
        logging.error(f"対処: ログファイルを確認し、エラー内容を確認してください。")
        print(f"\n❌ エラーが発生しました: {e}")
        print(f"   詳細はログファイルを確認してください。")
        return False


# ==================== エントリーポイント ====================

if __name__ == "__main__":
    # コマンドライン引数のパース
    parser = argparse.ArgumentParser(
        description="はがき宛名PDF自動生成システム - AddrLabel Pro",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  # テストモード（10件のみ処理）
  python generate_hagaki_labels.py --test

  # 本番モード（全件処理）
  python generate_hagaki_labels.py

  # カスタム入力ファイルを指定
  python generate_hagaki_labels.py --input custom_file.xlsx --sheet "シート1"

  # カスタム出力ファイルを指定
  python generate_hagaki_labels.py --output custom_output.pdf
        """
    )

    parser.add_argument(
        '--input', '-i',
        default=DEFAULT_EXCEL_FILE,
        help=f'入力Excelファイルのパス（デフォルト: {DEFAULT_EXCEL_FILE}）'
    )

    parser.add_argument(
        '--output', '-o',
        default=DEFAULT_OUTPUT_PDF,
        help=f'出力PDFファイルのパス（デフォルト: {DEFAULT_OUTPUT_PDF}）'
    )

    parser.add_argument(
        '--sheet', '-s',
        default=DEFAULT_SHEET_NAME,
        help=f'シート名（デフォルト: {DEFAULT_SHEET_NAME}）'
    )

    parser.add_argument(
        '--test', '-t',
        action='store_true',
        help='テストモード: 最初の10件のみを処理'
    )

    args = parser.parse_args()

    # メイン処理の実行
    success = main(
        excel_file=args.input,
        output_pdf=args.output,
        sheet_name=args.sheet,
        test_mode=args.test
    )

    # 終了コード
    sys.exit(0 if success else 1)
