# DartsApp — ダーツ上達確認アプリ

体験者がダーツを2セット（各3投）投げ、グルーピング半径の比較によって上達度を数値で可視化するアプリです。

---

## スクリーンショット

| セット1入力中 | 結果表示 |
|---|---|
| ダーツ盤をクリックして着弾点を記録 | バラつき・正確さの改善率を表示 |

---

## 機能

- **ダーツ盤クリック入力** — 画面上のダーツ盤をクリックして着弾点を記録
- **ミス入力対応** — 盤外クリックでミス（300mm 固定距離）として記録
- **自動セット移行** — セット1（3投）完了後、自動でセット2へ移行
- **データ修正** — 記録済みの投擲ラベルをクリックして再入力可能
- **グルーピング計算** — 重心からの最大距離（mm）でバラつきを数値化
- **改善率表示** — `(1回目 − 2回目) / 1回目 × 100%` で上達度を算出
- **Excel 出力** — 座標・距離・比較結果をスタイル付きで `.xlsx` 保存
- **保存後自動リセット** — Excel 出力後にデータを自動リセット
- **日英切り替え** — ヘッダーのボタンで日本語 / English を即時切替

---

## 動作環境

| プラットフォーム | 配布ファイル | 必要なもの |
|---|---|---|
| macOS 11 以降 | `DartsApp_Mac.zip` | なし（Python 不要） |
| Windows 10 / 11 | `DartsApp_Windows.zip` | なし（Python 不要） |
| Python 環境 | `darts_app.py` | Python 3.10以上、openpyxl |

---

## インストールと起動

### Mac

1. `DartsApp_Mac.zip` を解凍
2. `DartsApp.app` をダブルクリック

> **初回起動時のセキュリティ警告について**
> 「開発元を確認できません」と表示された場合：
> `右クリック（control+クリック）→「開く」→「開く」`
> または システム設定 → プライバシーとセキュリティ →「このまま開く」

### Windows

1. `DartsApp_Windows.zip` を解凍
2. `DartsApp.exe` をダブルクリック

> **SmartScreen 警告について**
> 「WindowsによってPCが保護されました」と表示された場合：
> 「詳細情報」→「実行」をクリック

### Python 環境から実行（開発者向け）

```bash
pip install openpyxl
python darts_app.py
```

---

## 使い方

```
1. 体験者名を入力
       ↓
2. セット1：ダーツ盤を3回クリックして着弾点を入力
   （盤外クリック = ミス扱い）
       ↓
3. セット2へ自動移行（操作不要）
       ↓
4. セット2：同様に3回クリック
       ↓
5. 右側パネルに結果が自動表示
       ↓
6.「Excelに保存」ボタンでデータを出力
```

### 入力を修正したい場合

記録済みのラベル（`✎` マーク付き）をクリックすると修正モードに入ります。
その状態でダーツ盤をクリックすると、そのデータが上書きされます。
同じラベルをもう一度クリックするとキャンセルできます。

---

## 計算方法

### グルーピング半径（バラつき）

```
重心 = 3投の平均座標
グルーピング半径 = 重心から各投までの距離の最大値 (mm)
```

ミス投は中心から 300mm の点として計算に含まれます。

### 改善率

```
改善率 = (1回目の半径 − 2回目の半径) / 1回目の半径 × 100 (%)
```

プラスなら上達、マイナスなら悪化。

### スケール

画面上のダーツ盤外周（ダブルリング）= 実寸 170mm として換算。

---

## ビルド方法（開発者向け）

### Mac

```bash
pip install pyinstaller openpyxl
pyinstaller --windowed --name "DartsApp" --icon darts.icns -y darts_app.py
# dist/DartsApp.app が生成される
```

### Windows

```powershell
pip install pyinstaller openpyxl
pyinstaller --windowed --name "DartsApp" --icon darts.ico --onefile -y darts_app.py
# dist/DartsApp.exe が生成される
```

### GitHub Actions（Mac/Windows 同時ビルド）

`main` ブランチへ push すると自動でビルドが走ります。
`.github/workflows/build.yml` を参照してください。

---

## ファイル構成

```
darts_app/
├── darts_app.py          # アプリ本体（Python）
├── darts.icns            # アプリアイコン（Mac用）
├── darts.ico             # アプリアイコン（Windows用）
├── .github/
│   └── workflows/
│       └── build.yml     # GitHub Actions ビルド設定
└── README.md             # このファイル
```

---

## 依存ライブラリ

| ライブラリ | 用途 | 備考 |
|---|---|---|
| tkinter | GUI フレームワーク | Python 標準ライブラリ |
| openpyxl | Excel 出力 | `pip install openpyxl` |
| math / datetime | 計算・タイムスタンプ | Python 標準ライブラリ |

---

## ライセンス

MIT License
