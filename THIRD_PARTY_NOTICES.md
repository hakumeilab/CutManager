# サードパーティ ソフトウェアの通知（THIRD PARTY NOTICES）

CutManager 本体は MIT License で配布されています（`LICENSE` を参照）。
配布物には以下のサードパーティ製ソフトウェアが含まれます。

---

## FFmpeg

CutManager は、動画サムネイルの生成に **FFmpeg**（`ffmpeg.exe`）を外部プログラムとして
呼び出して利用します。配布物には FFmpeg の実行ファイルを同梱しています。

- プロジェクト: https://ffmpeg.org/
- ライセンス: **GNU General Public License, version 3 (GPLv3)**
  （同梱している FFmpeg ビルドの構成による。ビルドに含まれるライセンス文書
  `ffmpeg-LICENSE.txt` を参照）
- 対応ソースコードの入手先: https://ffmpeg.org/download.html
  （同梱バイナリに対応するソースは、上記から入手できます）

FFmpeg は CutManager とは独立した別個のプログラムであり、`subprocess` を介して
外部プロセスとして起動されます。CutManager 本体のソースコードは FFmpeg と
リンク（結合）しておらず、引き続き MIT License の下で提供されます。

FFmpeg is a trademark of Fabrice Bellard, originator of the FFmpeg project.
