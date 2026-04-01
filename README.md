# CutManager

PySide6 で作る Windows 向けの CSV 編集 GUI ツールです。映像制作向けの固定列 CSV を軽快に編集し、素材フォルダーや動画ファイルのドラッグアンドドロップにも対応しています。

## セットアップ

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python main.py
```

## VS Code での実行

1. VS Code でこのフォルダーを開きます。
2. Python インタープリタに `.venv\Scripts\python.exe` を選びます。
3. 初回だけターミナルで以下を実行します。

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

4. `F5` を押して `.vscode/launch.json` の `CutManager` 構成で起動します。

## Windows 向けビルド

CutManager は `PySide6` に含まれる `pyside6-deploy` を使って `.exe` 化できます。  
まずは `standalone` 形式でビルドするのが扱いやすいです。

### 手順 1. VS Code でこのフォルダーを開く

VS Code で `CutManager` のフォルダーを開き、ターミナルを PowerShell にします。

### 手順 2. 仮想環境を作る

まだ `.venv` が無い場合は、次のコマンドを実行します。

```powershell
python -m venv .venv
```

### 手順 3. 仮想環境を有効化する

```powershell
.venv\Scripts\Activate.ps1
```

有効化できると、PowerShell の行頭に `(.venv)` が付きます。

### 手順 4. 必要なライブラリを入れる

```powershell
python -m pip install -r requirements.txt
```

### 手順 5. 開発実行でアプリが起動するか確認する

ビルド前に、まず通常起動できることを確認します。

```powershell
python main.py
```

CutManager の画面が開けば次へ進みます。確認できたらアプリを閉じます。

### 手順 6. `standalone` 形式でビルドする

```powershell
.venv\Scripts\pyside6-deploy.exe main.py -f --name CutManager --mode standalone
```

この設定では GUI アプリとしてビルドされるため、`CutManager.exe` 起動時にターミナルは開きません。

この処理では以下が起こります。

- `pyside6-deploy` がビルドを開始します。
- 初回は内部で `Nuitka` の取得が走ることがあり、少し時間がかかります。
- 完了すると `deployment\` 配下に出力されます。

### 手順 7. ビルド結果を確認する

エクスプローラーか VS Code で `deployment\` フォルダーを開きます。  
その中に `CutManager.exe` を含む一式ができていれば成功です。

### 手順 8. ビルドした exe を起動する

`deployment\` 配下の `CutManager.exe` を起動します。

確認する項目:

1. アプリが起動する
2. CSV を開ける
3. CSV を保存できる
4. 素材フォルダーや動画ファイルのドラッグアンドドロップが動く
5. コピー貼り付け、アンドゥ、リドゥが動く

### 手順 9. Python が無い PC でも確認する

可能なら、Python を入れていない別の Windows PC で `deployment\` 一式をコピーして起動確認します。  
配布前の最終確認として有効です。

### 手順 10. 1 ファイル配布にしたい場合だけ `onefile` を使う

1 ファイル化したい場合は、次のコマンドでビルドします。

```powershell
.venv\Scripts\pyside6-deploy.exe main.py -f --name CutManager --mode onefile
```

補足:

- `onefile` は配布しやすい反面、起動時に展開処理が入るため遅くなりやすいです。
- 普段使い用の配布では `standalone` の方が無難です。

### 手順 11. ビルド条件を固定したい場合は設定ファイルを使う

毎回同じ条件でビルドしたい場合だけ使います。

1. ひな形を作ります。

```powershell
.venv\Scripts\pyside6-deploy.exe --init main.py
```

2. 生成された `pysidedeploy.spec` を必要に応じて編集します。
3. 次のコマンドで設定ファイルを使ってビルドします。

```powershell
.venv\Scripts\pyside6-deploy.exe -c pysidedeploy.spec -f
```

### ビルドで詰まりやすい点

- 初回ビルドが長い:
  `Nuitka` の取得や解析に時間がかかることがあります。
- ビルドに失敗する:
  Windows の C/C++ ビルド環境が不足している可能性があります。
  Visual Studio Build Tools か Visual Studio の C++ 開発環境を入れて再実行してください。
- 配布後に起動しない:
  まず `onefile` ではなく `standalone` で再ビルドして切り分けるのが安全です。

## アプリ内更新

CutManager には、メニューバーの `ヘルプ > 更新を確認` から最新版を確認し、更新ファイルをダウンロードして適用する機能があります。

### 更新機能の前提

1. GitHub Releases に最新版を登録します。
2. リリースの `tag_name` は `v0.2.0` や `0.2.0` のようなバージョン形式にします。
3. リリース asset に Windows 向けの配布物を添付します。

対応している asset 形式:

- `.zip`
  推奨です。配布版フォルダー一式を zip にして添付します。
- `.exe`
  インストーラー形式でも使えます。この場合はアプリからインストーラーを起動します。

### zip asset の置き方

zip には、更新後のアプリ本体を入れます。次のどちらかの構成にしてください。

1. zip の直下に `CutManager.exe` など実行ファイルと関連ファイルを置く
2. zip の直下に 1 つだけフォルダーを置き、その中に実行ファイルと関連ファイルを入れる

### GitHub 側の公開条件

この更新機能は GitHub の `releases/latest` API を使います。  
そのため、アプリから更新確認できるのは、基本的に外部から読める GitHub Releases です。

- リポジトリや Release が public:
  そのまま更新確認できます。
- リポジトリが private:
  認証なしでは更新確認できません。

private リポジトリのまま運用する場合は、別途認証付き更新サーバーを作るか、公開された更新 JSON / 配布 URL を用意する必要があります。

### 更新用リリースの作り方

1. `pyside6-deploy` で `standalone` ビルドを作ります。
2. 出力された配布フォルダー一式を zip にまとめます。
3. GitHub Releases で新しいタグを作ります。
4. その Release に zip か installer の `.exe` を asset として添付します。
5. 配布版の CutManager で `ヘルプ > 更新を確認` を実行します。

手元でまとめて作る場合は、リポジトリ直下の `build_release.ps1` を使うと
`standalone` ビルド、配布 zip、SHA-256 ファイルの生成まで一括で実行できます。

```powershell
.\build_release.ps1
```

依存関係も合わせて入れ直したい場合:

```powershell
.\build_release.ps1 -InstallDependencies
```

生成物:

- `dist\CutManager-<version>-windows-standalone.zip`
- `dist\CutManager-<version>-windows-standalone.sha256.txt`

## 主な機能

- UTF-8 with BOM の CSV 読み込みと保存
- 固定ヘッダーの `QTableView` + `QAbstractTableModel`
- メニューバーの `ファイル` `編集` `並べ替え` に主要操作を集約
- `ファイル` メニュー内の `最近開いたファイル` から履歴を再オープン
- 前回開いていた CSV を起動時に自動で復元
- `編集` メニューからアンドゥ / リドゥ
- `環境設定` からアンドゥ履歴数を変更
- 読み込み直後と保存時は基本的にカット番号昇順で内部順序を維持
- 各列ヘッダーのクリックで昇順 / 降順を切り替え
- 各列ヘッダー右端の漏斗ボタンから候補値チェックボックス絞り込み
- `169_170` のような兼用名は複数カットとして分割し、`区分` 列へ `兼用` を自動入力
- `カット番号順に戻す` で既定順へ復帰
- `Delete` でセル内容クリア
- `Insert` で行追加
- `Ctrl+Delete` で選択行削除
- `Ctrl+C` / `Ctrl+V` でセルのコピー / 貼り付け
- CSV ファイルのドラッグアンドドロップで開く
- 素材フォルダーのドラッグアンドドロップでカット行を自動追加
- 動画ファイルのドラッグアンドドロップでテイク情報と納品日を反映

## 列構成

1. カット番号
2. AB分け
3. 区分
4. 素材入れ回数
5. 素材入れ日
6. テイク
7. テイク番号
8. 納品日

## 素材フォルダー取り込み

- ドロップしたフォルダー直下のサブフォルダーを走査します
- 直下にサブフォルダーが無い場合は、ドロップしたフォルダー名そのものも抽出対象にします
- フォルダー名に含まれる 3 桁数字をカット番号として抽出します
- `085A` `085B` のように末尾アルファベットが付く場合は、カット番号を `085` にしつつ `AB分け` に `A` `B` を入れて別行として扱います
- `169_170` のように複数の 3 桁数字があれば、兼用として複数行を追加します
- `区分` 列は `兼用` `BANK` `欠番` のプルダウン選択です
- 追加時の素材入れ回数は `1`、素材入れ日はドロップ当日を `YYYY/MM/DD` 形式で設定します
- 既存の `カット番号 + AB分け` と重複した行は追加しません

## 動画ファイル取り込み

- 対応拡張子は `mp4` `mov` `mxf` `avi` `wmv` `m4v`
- ファイル名から 3 桁のカット番号を抽出して既存行へ反映します
- `085A` `085B` のような名前は `カット番号=085` と `AB分け=A/B` の組み合わせで対象行を判定します
- `169_170` のような兼用名は該当する複数カットへ同時反映します
- 同じカットへ複数動画が当たる場合は、更新日時が新しいファイルの内容を優先します
- 納品日はドロップ当日を `YYYY/MM/DD` 形式で設定します
- `take02` `tk03` `t04` のような表記からテイク番号を抽出します

## ファイル構成

- `main.py`: エントリーポイント
- `cutmanager/main_window.py`: メイン画面と各種操作
- `cutmanager/model.py`: テーブルデータモデル
- `cutmanager/proxy.py`: 列ごとの候補値絞り込み
- `cutmanager/filter_popup.py`: 列ヘッダーから開く並べ替え / 絞り込みポップアップ
- `cutmanager/csv_io.py`: CSV 読み書き
- `cutmanager/folder_import.py`: 素材フォルダー取り込み
- `cutmanager/video_import.py`: 動画ファイル取り込み
- `cutmanager/view.py`: テーブル操作のカスタマイズ
