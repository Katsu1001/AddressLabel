# AddrLabel Pro - 使用ガイド

## クイックスタート（3ステップ）

### ステップ1: 環境構築（初回のみ）

```bash
# 1. 仮想環境の作成
conda create -n possim_hagaki python=3.11 -y

# 2. 仮想環境の有効化
conda activate possim_hagaki

# 3. ライブラリのインストール
conda install pandas openpyxl -y
pip install reportlab
```

### ステップ2: サンプルファイルでテスト

```bash
# サンプルExcelファイルの作成
python create_sample_excel.py

# テストモードで実行（10件のみ）
python generate_hagaki_labels.py --input sample_営業リスト.xlsx --test
```

### ステップ3: 本番データで実行

```bash
# Excelファイルをプロジェクトルートに配置
# 例: POSSIM_営業リスト_優先順位付き.xlsx

# 本番モードで実行（全件処理）
python generate_hagaki_labels.py
```

---

## 詳細な使用方法

### 実行前のチェックリスト

- [ ] miniforge3（またはAnaconda）がインストールされている
- [ ] 仮想環境 `possim_hagaki` が作成されている
- [ ] 必要なライブラリ（pandas, openpyxl, reportlab）がインストールされている
- [ ] Excelファイルが準備されている
- [ ] Excelファイルに以下の列が存在する:
  - 氏名
  - 郵便番号
  - 都道府県
  - 住所

### Excelファイルの準備

#### 必須列

| 列名 | 説明 | 例 |
|------|------|-----|
| 氏名 | 宛先の氏名 | 山田太郎 |
| 郵便番号 | 郵便番号（ハイフン付き） | 100-0001 |
| 都道府県 | 都道府県名 | 東京都 |
| 住所 | 市区町村以降の住所 | 千代田区千代田1-1 |

#### Excelファイル形式の要件

- **ファイル形式**: `.xlsx`（Excel 2007以降）
- **シート名**: デフォルトは「営業リスト」（変更可能）
- **ヘッダー行**: 1行目に列名が必要
- **データ開始行**: 2行目から

#### サンプルExcelレイアウト

```
| 氏名      | 郵便番号   | 都道府県 | 住所                  | 会社名 | 電話番号     |
|----------|-----------|---------|----------------------|--------|-------------|
| 山田太郎  | 100-0001  | 東京都   | 千代田区千代田1-1     | (任意) | (任意)      |
| 佐藤花子  | 150-0002  | 東京都   | 渋谷区渋谷2-2-2       | (任意) | (任意)      |
```

---

## コマンドライン引数の詳細

### 基本構文

```bash
python generate_hagaki_labels.py [オプション]
```

### オプション一覧

#### `--input` / `-i`
入力Excelファイルのパスを指定

```bash
python generate_hagaki_labels.py --input "path/to/file.xlsx"
```

**デフォルト**: `POSSIM_営業リスト_優先順位付き.xlsx`

---

#### `--output` / `-o`
出力PDFファイルのパスを指定

```bash
python generate_hagaki_labels.py --output "output/custom_name.pdf"
```

**デフォルト**: `output/hagaki_labels.pdf`

---

#### `--sheet` / `-s`
読み込むシート名を指定

```bash
python generate_hagaki_labels.py --sheet "顧客マスタ"
```

**デフォルト**: `営業リスト`

---

#### `--test` / `-t`
テストモード（最初の10件のみ処理）

```bash
python generate_hagaki_labels.py --test
```

**用途**:
- 初回実行時の動作確認
- レイアウト調整時の確認
- エラーの切り分け

---

### 実行例

#### 例1: デフォルト設定で実行

```bash
python generate_hagaki_labels.py
```

この場合:
- 入力: `POSSIM_営業リスト_優先順位付き.xlsx`
- シート: `営業リスト`
- 出力: `output/hagaki_labels.pdf`
- モード: 本番モード（全件処理）

---

#### 例2: カスタムファイルでテスト実行

```bash
python generate_hagaki_labels.py \
  --input "営業リスト_2025年度.xlsx" \
  --test
```

この場合:
- 入力: `営業リスト_2025年度.xlsx`
- シート: `営業リスト`（デフォルト）
- 出力: `output/hagaki_labels.pdf`（デフォルト）
- モード: テストモード（10件のみ）

---

#### 例3: すべてカスタマイズして実行

```bash
python generate_hagaki_labels.py \
  --input "data/customer_list.xlsx" \
  --sheet "顧客マスタ" \
  --output "output/hagaki_2025_Q1.pdf"
```

この場合:
- 入力: `data/customer_list.xlsx`
- シート: `顧客マスタ`
- 出力: `output/hagaki_2025_Q1.pdf`
- モード: 本番モード（全件処理）

---

## 出力ファイルの確認

### PDFファイル

#### 保存場所
- デフォルト: `output/hagaki_labels.pdf`
- カスタム: `--output` で指定した場所

#### 確認方法

**macOS**:
```bash
open output/hagaki_labels.pdf
```

**Linux**:
```bash
xdg-open output/hagaki_labels.pdf
```

#### PDFの構成
- **ページ数**: データ件数と同じ（1件 = 1ページ = 1枚のはがき）
- **ページサイズ**: 100mm × 148mm（はがきサイズ）
- **向き**: 縦
- **レイアウト**: 日本郵便標準の縦書き形式

---

### ログファイル

#### 保存場所
`logs/hagaki_generation_YYYYMMDD_HHMMSS.log`

例: `logs/hagaki_generation_20251106_143522.log`

#### 確認方法

```bash
# 最新のログファイルを表示
ls -lt logs/ | head -n 2

# ログファイルの内容を表示
cat logs/hagaki_generation_20251106_143522.log

# リアルタイムでログを表示（処理中）
tail -f logs/hagaki_generation_20251106_143522.log
```

#### ログレベル

| レベル | 説明 | 例 |
|--------|------|-----|
| INFO | 通常の処理情報 | データを読み込みました: 1507行 × 20列 |
| WARNING | 警告（処理は継続） | 列 '郵便番号': 226件の欠損データ (15.0%) |
| ERROR | エラー（該当データをスキップ） | はがき生成エラー（行 123）: ... |

---

## エラー対応

### よくあるエラーと対処方法

#### エラー1: ModuleNotFoundError

```
ModuleNotFoundError: No module named 'pandas'
```

**原因**: ライブラリがインストールされていない

**対処**:
```bash
conda activate possim_hagaki
conda install pandas openpyxl -y
pip install reportlab
```

---

#### エラー2: FileNotFoundError

```
FileNotFoundError: [Errno 2] No such file or directory: 'POSSIM_営業リスト_優先順位付き.xlsx'
```

**原因**: Excelファイルが見つからない

**対処**:
1. ファイルがプロジェクトルートに存在するか確認
   ```bash
   ls -la *.xlsx
   ```

2. ファイルパスを明示的に指定
   ```bash
   python generate_hagaki_labels.py --input "/full/path/to/file.xlsx"
   ```

---

#### エラー3: ValueError (シートが存在しない)

```
ValueError: Worksheet 営業リスト does not exist.
```

**原因**: 指定されたシート名が存在しない

**対処**:
1. Excelファイルを開いてシート名を確認
2. 正しいシート名を指定
   ```bash
   python generate_hagaki_labels.py --sheet "正しいシート名"
   ```

---

#### エラー4: KeyError (列が存在しない)

```
KeyError: '氏名'
```

**原因**: 必要な列が存在しない

**対処**:
1. Excelファイルに以下の列が存在するか確認:
   - 氏名
   - 郵便番号
   - 都道府県
   - 住所

2. 列名が異なる場合は、スクリプトを編集
   ```python
   # generate_hagaki_labels.py の30行目付近
   COLUMN_NAME = "氏名"  # ← Excelの実際の列名に変更
   ```

---

#### エラー5: PermissionError

```
PermissionError: [Errno 13] Permission denied: 'output/hagaki_labels.pdf'
```

**原因**: PDFファイルが他のアプリで開かれている

**対処**:
1. PDFファイルを閉じる
2. または、別のファイル名で出力
   ```bash
   python generate_hagaki_labels.py --output "output/hagaki_new.pdf"
   ```

---

## パフォーマンス最適化

### 処理時間の目安

| データ件数 | MacBook Pro M4 | Intel Mac (Core i5) |
|-----------|---------------|-------------------|
| 10件 | 約5秒 | 約10秒 |
| 100件 | 約10秒 | 約20秒 |
| 1,000件 | 約1分 | 約2分 |
| 1,507件 | 約1.5分 | 約3分 |

### 処理を高速化する方法

1. **他のアプリケーションを閉じる**
   - ブラウザ、動画再生アプリなどを終了

2. **テストモードで事前確認**
   ```bash
   python generate_hagaki_labels.py --test
   ```
   - 問題がないことを確認してから本番実行

3. **バックグラウンド実行**
   ```bash
   python generate_hagaki_labels.py > process.log 2>&1 &
   ```
   - 長時間処理の場合、バックグラウンドで実行

---

## カスタマイズ

### レイアウトの調整

`generate_hagaki_labels.py` の定数セクション（60行目付近）を編集:

```python
# 郵便番号の位置を調整
POSTAL_CODE_X = 10 * mm   # 左から10mm → 変更可能
POSTAL_CODE_Y = 130 * mm  # 下から130mm → 変更可能

# 住所の位置を調整
ADDRESS_X = 80 * mm       # 左から80mm → 変更可能
ADDRESS_Y = 120 * mm      # 下から120mm → 変更可能

# 氏名の位置を調整
NAME_X = 35 * mm          # 左から35mm → 変更可能
NAME_Y = 50 * mm          # 下から50mm → 変更可能
```

### フォントサイズの調整

```python
# フォントサイズを変更
FONT_SIZE_POSTAL = 14     # 郵便番号 → 変更可能
FONT_SIZE_ADDRESS = 12    # 住所 → 変更可能
FONT_SIZE_NAME = 16       # 氏名 → 変更可能
```

### フォントの変更

```python
# フォントを変更
FONT_NAME = "HeiseiMin-W3"      # 明朝体
# FONT_NAME = "HeiseiKakuGo-W5" # ゴシック体
```

---

## ベストプラクティス

### 1. 初回実行時の推奨フロー

```bash
# ① サンプルファイルで動作確認
python create_sample_excel.py
python generate_hagaki_labels.py --input sample_営業リスト.xlsx --test

# ② 本番ファイルでテストモード実行
python generate_hagaki_labels.py --test

# ③ 問題がなければ本番実行
python generate_hagaki_labels.py
```

### 2. 大量データ処理の推奨フロー

```bash
# ① テストモードで確認
python generate_hagaki_labels.py --test

# ② バックグラウンドで本番実行
nohup python generate_hagaki_labels.py > process.log 2>&1 &

# ③ 進捗確認
tail -f logs/hagaki_generation_*.log

# ④ 完了確認
jobs  # 実行中のジョブを確認
```

### 3. データ品質チェックの推奨フロー

```bash
# ① テストモード実行でログを確認
python generate_hagaki_labels.py --test

# ② ログから欠損データを確認
cat logs/hagaki_generation_*.log | grep WARNING

# ③ 必要に応じてExcelファイルを修正

# ④ 本番実行
python generate_hagaki_labels.py
```

---

## サポート

問題が解決しない場合は、以下の情報を収集してサポートチームに連絡してください。

### 必要な情報

1. **エラーメッセージの全文**
   ```bash
   python generate_hagaki_labels.py 2>&1 | tee error.log
   ```

2. **ログファイル**
   ```bash
   cat logs/hagaki_generation_*.log
   ```

3. **環境情報**
   ```bash
   python --version
   conda list
   uname -a
   ```

4. **Excelファイルのサンプル**
   - 個人情報を削除したサンプルデータ

---

**最終更新: 2025-11-06**
