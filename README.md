# 商品カタログ生成システム

商品マスタ（Excel）と画像から、仕入先別の商品カタログを自動生成します。

## セットアップ

### 1. 依存パッケージ

```bash
pip install -r requirements.txt
```

### 2. データ配置

```
data/
└── 受発注管理台帳.xlsx    # 商品マスタ

images/
└── HAK/                   # 仕入先コード別
    ├── PRD_SNJ_HAK_0001_01.jpg
    └── ...
```

## 使い方

### カタログ生成

```bash
# 基本
python generate_catalog.py HAK

# オプション指定
python generate_catalog.py HAK \
  --excel data/受発注管理台帳.xlsx \
  --template templates/catalog_template.pptx
```

### PDF変換

```bash
libreoffice --headless --convert-to pdf --outdir output output/カタログ_*.pptx
```

## テンプレート編集

`templates/catalog_template.pptx` をPowerPointで自由に編集できます。

プレースホルダー（`{{商品名_1}}` など）の形式を維持してください。

## 仕入先コード一覧（主要）

| コード | 仕入先名 | 商品数 |
|--------|----------|--------|
| GGM | GOODGOOD | 97 |
| FJT | フジタコーポレーション | 37 |
| HAK | 北見ハッカ通商 | 16 |

詳細は商品マスタを参照。
