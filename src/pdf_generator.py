"""
PDF生成モジュール
宛名ラベルPDFを生成する
"""

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
import pandas as pd


class PDFGenerator:
    """
    宛名ラベルPDFを生成するクラス
    """

    def __init__(self, df: pd.DataFrame, output_path: str):
        """
        コンストラクタ

        Args:
            df (pd.DataFrame): クレンジング済みデータフレーム
            output_path (str): 出力PDFファイルパス
        """
        self.df = df
        self.output_path = output_path

        # 日本語フォントを登録
        self._register_font()

    def _register_font(self):
        """日本語フォントを登録"""
        try:
            # まず、ビルトインのCIDフォント（HeiseiMin-W3）を試す
            pdfmetrics.registerFont(UnicodeCIDFont('HeiseiMin-W3'))
            self.font_name = 'HeiseiMin-W3'
            print("✓ 日本語フォント（HeiseiMin-W3）を登録しました")
        except:
            try:
                # 次に、HeiseiKakuGo-W5を試す
                pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
                self.font_name = 'HeiseiKakuGo-W5'
                print("✓ 日本語フォント（HeiseiKakuGo-W5）を登録しました")
            except:
                try:
                    # Mac環境: ヒラギノ明朝を登録
                    pdfmetrics.registerFont(
                        TTFont('HiraginoMincho', '/System/Library/Fonts/ヒラギノ明朝 ProN W3.ttc')
                    )
                    self.font_name = 'HiraginoMincho'
                    print("✓ 日本語フォント（ヒラギノ明朝）を登録しました")
                except:
                    try:
                        # Linux環境: IPAexMinchoを使用
                        pdfmetrics.registerFont(
                            TTFont('IPAexMincho', '/usr/share/fonts/truetype/fonts-japanese-mincho.ttf')
                        )
                        self.font_name = 'IPAexMincho'
                        print("✓ 日本語フォント（IPAexMincho）を登録しました")
                    except:
                        print("⚠ 警告: 専用の日本語フォントが見つかりません。Courier（デフォルト）を使用します。")
                        print("  日本語が正しく表示されない可能性があります。")
                        self.font_name = 'Courier'

    def generate(self):
        """PDFを生成"""
        # A4サイズのキャンバスを作成
        pdf = canvas.Canvas(self.output_path, pagesize=A4)
        width, height = A4

        # 10面レイアウトの設定
        # A4を2列×5行に分割
        label_width = width / 2
        label_height = height / 5

        # 配置位置の定義（左列5個、右列5個）
        positions = []
        for i in range(5):
            # 左列
            positions.append((0, height - label_height * (i + 1)))
        for i in range(5):
            # 右列
            positions.append((label_width, height - label_height * (i + 1)))

        # 宛名データを10枚ずつ処理
        page_count = 0
        total_generated = 0

        for index, row in self.df.iterrows():
            # 郵便番号、住所が欠損している場合はスキップ
            if pd.isna(row['郵便番号']) or row['住所_整形済み'] == '':
                continue

            # 国外住所はスキップ
            if row['国外住所フラグ']:
                continue

            # 配置位置を取得
            position_index = page_count % 10
            x, y = positions[position_index]

            # 宛名を描画
            self._draw_label(pdf, x, y, row, label_width, label_height)

            # ページカウントを増やす
            page_count += 1
            total_generated += 1

            # 10枚ごとに次のページへ
            if page_count % 10 == 0:
                pdf.showPage()

        # 最後のページを保存
        if page_count % 10 != 0:  # 最後のページが途中で終わっている場合
            pdf.showPage()

        pdf.save()
        print(f"✅ PDFを生成しました: {self.output_path}")
        print(f"   生成枚数: {total_generated}枚")

    def _draw_label(self, pdf: canvas.Canvas, x: float, y: float, row: pd.Series,
                    label_width: float, label_height: float):
        """
        1枚の宛名ラベルを描画

        Args:
            pdf: reportlabのCanvasオブジェクト
            x: X座標
            y: Y座標
            row: データフレームの1行
            label_width: ラベルの幅
            label_height: ラベルの高さ
        """
        # デバッグ用の枠線を描画（オプション）
        # pdf.rect(x, y, label_width, label_height)

        # 郵便番号
        pdf.setFont(self.font_name, 11)
        postal_text = f"〒 {row['郵便番号']}"
        pdf.drawString(x + 10, y + label_height - 30, postal_text)

        # 住所（複数行に分割）
        address_lines = self._split_address(row['住所_整形済み'], max_length=25)
        pdf.setFont(self.font_name, 10)
        for i, line in enumerate(address_lines):
            pdf.drawString(x + 10, y + label_height - 50 - (i * 14), line)

        # 氏名
        pdf.setFont(self.font_name, 12)
        name_y = y + label_height - 50 - (len(address_lines) * 14) - 20
        # 最低位置を確保
        if name_y < y + 10:
            name_y = y + 10
        pdf.drawString(x + 10, name_y, row['氏名_整形済み'])

    def _split_address(self, address: str, max_length: int = 25) -> list:
        """
        住所を複数行に分割

        Args:
            address: 住所文字列
            max_length: 1行の最大文字数

        Returns:
            list: 分割された住所のリスト
        """
        if not address or address == '':
            return ['']

        lines = []
        current_line = ""

        for char in address:
            current_line += char
            if len(current_line) >= max_length:
                lines.append(current_line)
                current_line = ""

        if current_line:
            lines.append(current_line)

        return lines if lines else ['']
