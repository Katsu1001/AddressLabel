# AddrLabel Pro - はがき宛名印刷自動化システム

## 📋 目次

- [プロジェクト概要](#プロジェクト概要)
- [主な機能](#主な機能)
- [技術スタック](#技術スタック)
- [システム要件](#システム要件)
- [環境構築手順](#環境構築手順)
- [使用方法](#使用方法)
- [出力例](#出力例)
- [トラブルシューティング](#トラブルシューティング)
- [FAQ（よくある質問）](#faqよくある質問)
- [ライセンス](#ライセンス)

---

## プロジェクト概要

**AddrLabel Pro**は、Excelファイルから顧客情報を読み込み、日本郵便標準のはがき宛名PDFを自動生成するシステムです。

### 背景
株式会社POSSIMのヘルスケアアプリ開発プロジェクトにおいて、約1,500件の営業先へのはがき発送業務を自動化するために開発されました。

### 処理性能
- **処理件数**: 1,507件（全件）
- **処理時間**: 1～2分（MacBook Pro M4想定）
- **出力形式**: 単一PDFファイル（各ページが1枚のはがきに対応）
- **はがきサイズ**: 100mm × 148mm（日本郵便標準）

---

## 主な機能

### ✅ 自動化機能
- Excelファイル（`.xlsx`）からの顧客データ自動読み込み
- 日本郵便標準レイアウトでのPDF自動生成
- 1,500件以上のデータを1～2分で処理（M4 Mac）

### ✅ データ品質管理
- 欠損データの自動検出（郵便番号、住所、氏名など）
- 欠損率の計算とレポート出力
- 欠損データがある場合も処理を継続（警告表示）

### ✅ 進捗管理
- リアルタイム進捗表示（100件ごと）
- 詳細なログ出力（タイムスタンプ付き）
- エラー発生時の詳細ログ記録

### ✅ テスト機能
- テストモード対応（10件のみで動作確認可能）
- コマンドライン引数でのカスタマイズ対応

### ✅ 日本語対応
- 日本語フォント自動設定（明朝体: HeiseiMin-W3）
- 代替フォント自動切り替え（ゴシック体: HeiseiKakuGo-W5）
- 縦書きレイアウト完全対応

---

## 技術スタック

### プログラミング言語
- **Python 3.11以上**

### 主要ライブラリ
| ライブラリ | バージョン | 用途 |
|----------|----------|------|
| pandas | 2.0.0以上 | Excelデータの読み込みと操作 |
| openpyxl | 3.1.0以上 | Excelファイル（.xlsx）の読み込み |
| reportlab | 4.0.0以上 | PDF生成、日本語縦書き対応 |

### 開発環境
- **ハードウェア**: MacBook Pro M4チップ搭載
- **パッケージマネージャ**: miniforge3（Apple Siliconネイティブ対応）
- **エディタ**: Visual Studio Code

---

## システム要件

### ハードウェア要件
- **CPU**: Apple Silicon（M1以降）推奨、Intel Macでも動作可能
- **メモリ**: 4GB以上（8GB推奨）
- **ディスク空き容量**: 500MB以上

### ソフトウェア要件
- **OS**: macOS 11.0 (Big Sur) 以降
- **Python**: 3.11以上
- **Conda**: miniforge3 または Anaconda

---

## 環境構築手順

### Step 1: miniforge3のインストール（初回のみ）

```bash
# ① Homebrewがインストールされていない場合
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

# ② miniforge3のインストール
brew install --cask miniforge

# ③ シェルの再起動
source ~/.zshrc  # zshの場合
# または
source ~/.bash_profile  # bashの場合
```

### Step 2: プロジェクトのクローン

```bash
# プロジェクトをクローン（Gitリポジトリの場合）
git clone <リポジトリURL>
cd AddressLabel

# または、ZIPファイルをダウンロードして解凍
```

### Step 3: 仮想環境の作成

```bash
# Conda仮想環境の作成（Python 3.11）
conda create -n possim_hagaki python=3.11 -y

# 仮想環境の有効化
conda activate possim_hagaki
```

### Step 4: 必要なライブラリのインストール

```bash
# pandasとopenpyxlをcondaでインストール
conda install pandas openpyxl -y

# reportlabをpipでインストール
pip install reportlab
```

### Step 5: インストールの確認

```bash
# Pythonバージョンの確認
python --version
# 出力例: Python 3.11.x

# インストール済みライブラリの確認
pip list | grep -E "pandas|openpyxl|reportlab"
# 出力例:
# pandas        2.x.x
# openpyxl      3.x.x
# reportlab     4.x.x
```

### Step 6: Excelファイルの配置

```bash
# プロジェクトルートに「POSSIM_営業リスト_優先順位付き.xlsx」を配置
# ディレクトリ構成例:
# AddressLabel/
# ├── generate_hagaki_labels.py
# ├── POSSIM_営業リスト_優先順位付き.xlsx  ← ここに配置
# ├── requirements.txt
# ├── README.md
# ├── logs/
# └── output/
```

---

## 使用方法

### 基本的な使い方

#### 1. テストモード（10件のみ処理）

初回実行時や動作確認時は、テストモードの使用を推奨します。

```bash
# テストモードで実行
python generate_hagaki_labels.py --test
```

**出力:**
- `output/hagaki_labels.pdf` - 最初の10件のはがきPDF
- `logs/hagaki_generation_YYYYMMDD_HHMMSS.log` - 処理ログ

#### 2. 本番モード（全件処理）

```bash
# 本番モードで実行（1507件すべてを処理）
python generate_hagaki_labels.py
```

**処理時間:** 約1～2分（MacBook Pro M4想定）

### コマンドライン引数

#### カスタム入力ファイルを指定

```bash
python generate_hagaki_labels.py --input "カスタムファイル.xlsx"
```

#### カスタムシート名を指定

```bash
python generate_hagaki_labels.py --sheet "シート1"
```

#### カスタム出力ファイル名を指定

```bash
python generate_hagaki_labels.py --output "output/custom_output.pdf"
```

#### すべての引数を組み合わせる

```bash
python generate_hagaki_labels.py \
  --input "営業リスト.xlsx" \
  --sheet "顧客マスタ" \
  --output "output/hagaki_2025.pdf" \
  --test
```

### ヘルプの表示

```bash
python generate_hagaki_labels.py --help
```

---

## 出力例

### ログ出力例

```
2025-11-06 10:00:00 - INFO - ログファイルを作成しました: logs/hagaki_generation_20251106_100000.log
2025-11-06 10:00:00 - INFO - ================================================================================
2025-11-06 10:00:00 - INFO - はがき宛名PDF自動生成システム - AddrLabel Pro
2025-11-06 10:00:00 - INFO - ================================================================================
2025-11-06 10:00:00 - INFO - 実行モード: 本番モード（全件）
2025-11-06 10:00:00 - INFO - 入力ファイル: POSSIM_営業リスト_優先順位付き.xlsx
2025-11-06 10:00:00 - INFO - 出力ファイル: output/hagaki_labels.pdf
2025-11-06 10:00:00 - INFO - シート名: 営業リスト
2025-11-06 10:00:01 - INFO - Excelファイルを読み込んでいます: POSSIM_営業リスト_優先順位付き.xlsx
2025-11-06 10:00:02 - INFO - データを読み込みました: 1507行 × 20列
2025-11-06 10:00:02 - INFO - 欠損データを検出しています
2025-11-06 10:00:02 - WARNING - 列 '郵便番号': 226件の欠損データ (15.0%)
2025-11-06 10:00:02 - WARNING - 列 '住所': 331件の欠損データ (22.0%)
2025-11-06 10:00:02 - INFO - 処理を開始します。総件数: 1507件
2025-11-06 10:00:10 - INFO - 処理中... [100/1507件] (6.6%)
2025-11-06 10:00:20 - INFO - 処理中... [200/1507件] (13.3%)
...
2025-11-06 10:01:30 - INFO - 処理中... [1500/1507件] (99.5%)
2025-11-06 10:01:32 - INFO - PDFファイルを保存しています...
2025-11-06 10:01:35 - INFO - ================================================================================
2025-11-06 10:01:35 - INFO - 処理完了
2025-11-06 10:01:35 - INFO - 成功: 1507件
2025-11-06 10:01:35 - INFO - エラー: 0件
2025-11-06 10:01:35 - INFO - 合計: 1507件
2025-11-06 10:01:35 - INFO - 出力ファイル: output/hagaki_labels.pdf
2025-11-06 10:01:35 - INFO - ファイルサイズ: 12.34 MB
```

### PDF出力レイアウト

```
┌───────────────────────────┐
│                           │
│  〒 123-4567              │ ← 郵便番号（左上）
│                           │
│                   東      │
│                   京      │
│                   都      │
│                   渋  様  │ ← 住所（右寄せ、縦書き）
│                   谷      │    氏名（中央、縦書き）
│                   区      │
│                   〇      │
│                   〇      │
│               山  １      │
│               田  −      │
│               太  ２      │
│               郎  −      │
│                   ３      │
│                           │
└───────────────────────────┘
100mm × 148mm（日本郵便標準サイズ）
```

---

## トラブルシューティング

### 問題1: `ModuleNotFoundError: No module named 'pandas'`

**原因**: 必要なライブラリがインストールされていない

**解決方法**:
```bash
# 仮想環境が有効化されていることを確認
conda activate possim_hagaki

# ライブラリを再インストール
conda install pandas openpyxl -y
pip install reportlab
```

---

### 問題2: `FileNotFoundError: [Errno 2] No such file or directory: 'POSSIM_営業リスト_優先順位付き.xlsx'`

**原因**: Excelファイルが見つからない

**解決方法**:
```bash
# 現在のディレクトリを確認
pwd

# ファイルが存在するか確認
ls -la *.xlsx

# ファイルをプロジェクトルートに配置
# または、--input 引数でパスを指定
python generate_hagaki_labels.py --input "/path/to/your/file.xlsx"
```

---

### 問題3: `ValueError: Worksheet 営業リスト does not exist.`

**原因**: 指定されたシート名が存在しない

**解決方法**:
```bash
# Excelファイルのシート名を確認
# または、正しいシート名を--sheet引数で指定
python generate_hagaki_labels.py --sheet "正しいシート名"
```

---

### 問題4: 日本語が文字化けする

**原因**: フォントの設定に問題がある

**解決方法**:

スクリプト内のフォント設定を確認:
```python
# generate_hagaki_labels.py の FONT_NAME 定数を確認
FONT_NAME = "HeiseiMin-W3"  # または "HeiseiKakuGo-W5"
```

reportlabの日本語フォントは以下が利用可能:
- `HeiseiMin-W3` - 明朝体
- `HeiseiKakuGo-W5` - ゴシック体

---

### 問題5: 処理が遅い（1507件で5分以上かかる）

**原因**: CPUリソースが不足している、または他のアプリケーションが実行中

**解決方法**:
1. 他のアプリケーションを閉じる
2. MacのActivity Monitorで他の重いプロセスを確認
3. テストモードで10件のみ処理して、スクリプト自体の問題か確認
   ```bash
   python generate_hagaki_labels.py --test
   ```

---

### 問題6: `PermissionError: [Errno 13] Permission denied: 'output/hagaki_labels.pdf'`

**原因**: 出力PDFファイルが他のアプリケーション（Adobe Acrobat、Preview.appなど）で開かれている

**解決方法**:
1. PDFファイルを閉じる
2. 別のファイル名で出力
   ```bash
   python generate_hagaki_labels.py --output "output/hagaki_labels_new.pdf"
   ```

---

### 問題7: Apple Siliconで`pandas`が動作しない

**原因**: Intel版のPythonまたはライブラリを使用している

**解決方法**:
```bash
# Pythonアーキテクチャの確認
python -c "import platform; print(platform.machine())"
# 出力が "arm64" であればOK（Apple Silicon用）
# 出力が "x86_64" の場合はIntel版

# miniforge3を使用してApple Silicon用の環境を再構築
conda create -n possim_hagaki python=3.11 -y
conda activate possim_hagaki
conda install pandas openpyxl -y
pip install reportlab
```

---

## FAQ（よくある質問）

### Q1: Excelファイルの列名は変更できますか?

**A**: はい、可能です。`generate_hagaki_labels.py` の定数セクションを編集してください。

```python
# ファイルの先頭部分（行30付近）
COLUMN_NAME = "氏名"          # ← 変更可能
COLUMN_POSTAL_CODE = "郵便番号"  # ← 変更可能
COLUMN_ADDRESS = "住所"        # ← 変更可能
COLUMN_PREFECTURE = "都道府県"  # ← 変更可能
```

---

### Q2: 欠損データがある場合、どのように表示されますか?

**A**: 以下のように表示されます。

- **郵便番号が欠損**: `〒 郵便番号未記入`
- **住所が欠損**: `住所未記入` と縦書きで表示
- **氏名が欠損**: `氏名未記入 様` と縦書きで表示

処理は継続され、すべてのデータがPDFに出力されます。

---

### Q3: レイアウトをカスタマイズできますか?

**A**: はい、可能です。`generate_hagaki_labels.py` の定数セクションで座標を変更してください。

```python
# PDFレイアウト設定（単位: mm）
POSTAL_CODE_X = 10 * mm   # 郵便番号のX座標
POSTAL_CODE_Y = 130 * mm  # 郵便番号のY座標

ADDRESS_X = 80 * mm       # 住所のX座標
ADDRESS_Y = 120 * mm      # 住所のY座標

NAME_X = 35 * mm          # 氏名のX座標
NAME_Y = 50 * mm          # 氏名のY座標
```

---

### Q4: CSVファイルには対応していますか?

**A**: 現在はExcel（.xlsx）形式のみ対応しています。CSVファイルを使用する場合は、以下の手順でExcelに変換してください。

1. CSVファイルをExcelで開く
2. 「名前を付けて保存」→ 「Excelブック（.xlsx）」形式で保存
3. 保存したファイルを使用

---

### Q5: WindowsやLinuxでも動作しますか?

**A**: 基本的には動作しますが、以下の点に注意してください。

**Windows**:
- miniforge3の代わりにAnacondaを使用
- パスの区切り文字が異なるため、ファイルパスをWindowsスタイルに変更

**Linux**:
- 日本語フォントが標準でインストールされていない場合があるため、フォント設定の調整が必要

---

### Q6: 複数のExcelファイルを一度に処理できますか?

**A**: 現在の実装では1ファイルのみ対応しています。複数ファイルを処理する場合は、以下のようなシェルスクリプトを作成してください。

```bash
#!/bin/bash
for file in *.xlsx; do
  python generate_hagaki_labels.py --input "$file" --output "output/${file%.xlsx}.pdf"
done
```

---

### Q7: 処理をバックグラウンドで実行できますか?

**A**: はい、可能です。

```bash
# バックグラウンドで実行
python generate_hagaki_labels.py > process.log 2>&1 &

# プロセスIDを確認
jobs

# ログをリアルタイムで確認
tail -f logs/hagaki_generation_*.log
```

---

## プロジェクト構成

```
AddressLabel/
├── generate_hagaki_labels.py  # メインスクリプト
├── requirements.txt            # 必要なライブラリリスト
├── README.md                   # このファイル
├── .gitignore                  # Git除外設定
├── logs/                       # ログファイル保存ディレクトリ
│   └── hagaki_generation_*.log
└── output/                     # PDF出力ディレクトリ
    └── hagaki_labels.pdf
```

---

## ライセンス

このプロジェクトは株式会社POSSIM内部用途で開発されました。

---

## サポート

問題が発生した場合は、以下の情報を含めて開発チームに連絡してください。

1. エラーメッセージの全文
2. ログファイル（`logs/hagaki_generation_*.log`）
3. 実行環境の情報
   ```bash
   python --version
   conda list
   ```
4. Excelファイルのサンプル（個人情報を削除したもの）

---

**開発: AddrLabel Pro Development Team**
**最終更新: 2025-11-06**
