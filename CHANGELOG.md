# Changelog

## Unreleased

- ダークテーマ環境でも表の非選択時テキストが背景に埋もれないよう、表と選択色をアプリのパレットに追従する描画へ変更
- 交互行の色分けと `兼用` `BANK` `欠番` の行色も、ライトテーマ / ダークテーマに応じて見やすい色へ追従するように変更
- `区分` 列で既存値のあるセルをダブルクリックしたとき、プルダウンが二重に見えることがある問題を修正

## 0.2.8 - 2026-04-14

- Fixed dark/light theme switching so table, headers, and menu bar colors refresh consistently.
- Updated alternating rows and status row colors for Shared, BANK, and Missing states.
- Fixed row-wide color updates when changing the Status column.
- Moved the drop hint to an overlay so the table no longer gains outer padding while dragging.
- Tightened cell selection and inline editor spacing for a cleaner table layout.

## 0.2.7 - 2026-04-13
- 空の CSV に動画ファイルをドロップしたとき、確認後に動画名からカットを仮登録して納品情報を反映できるように変更
- 既存 CSV で一致しない動画があるとき、未一致ファイル名を一覧表示し、必要なら未一致分だけ仮登録できるように変更
- README の構成を見直し、ユーザー向けの使い方を前半、セットアップやビルド手順を後半へ移動
- README の最新版ダウンロード手順に単体版 `windows-onefile.exe` とインストーラー版 `windows-setup.exe` の違いを追記

## 0.2.6 - 2026-04-03

- Release asset に Inno Setup の `windows-setup.exe` を追加
- アプリ内更新では onefile exe の自己上書きではなく、setup exe の起動を優先するように変更
- インストール先の既定値を `%LOCALAPPDATA%\Programs\CutManager` に設定

## 0.2.5 - 2026-04-03

- Release workflow の GitHub Release 作成/更新 step が、初回作成時でも落ちないように修正

## 0.2.4 - 2026-04-03

- `aaa_01_001_B1.mov` のような動画ファイル名から `テイク=B` `テイク番号=1` を抽出できるように変更
- `take02` `tk03` `t04` 形式は引き続き `テイク=T` として反映

## 0.2.3 - 2026-04-02

- 配布形式を onefile の単体 exe に切り替え
- README の冒頭に最新版 exe を Release からダウンロードする手順を追加
- アプリ内更新で onefile の exe をダウンロードして自己置換できるように変更

## 0.2.2 - 2026-04-02

- Nuitka 配布版でも zip 更新の自動適用を利用できるように更新実行環境の判定を修正
- zip 更新が使えない場合のメッセージを実行ファイル名に依存しない形へ調整

## 0.2.1 - 2026-04-02

- `区分` 列の値に応じて行全体の背景色を表示するように変更
- `欠番` は濃いグレー、`兼用` は薄い緑、`BANK` は薄い赤で表示

## 0.2.0 - 2026-04-02

- GitHub Releases を参照するアプリ内更新チェックを追加
- リリース asset のダウンロードと Windows 向け更新適用を追加
- GitHub Releases を使った配布手順を README に整理
- `build_release.ps1` による配布 zip と SHA-256 の生成手順を追加

## 0.1.0 - 2026-04-01

- CSV 編集 GUI の初回リリース
- 素材フォルダーと動画ファイルのドラッグアンドドロップ取り込みを実装
- 並べ替え、絞り込み、アンドゥ/リドゥ、最近使ったファイルを実装
