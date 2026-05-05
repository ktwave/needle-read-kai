ポケットモンスター ウルトラサン・ウルトラムーンの乱数調整向けに、<br>偽トロキャプチャ映像から針番号を自動判定しGen7 Main RNG Toolへ出力するツールです。

## 動作環境

- Windows 11
- 偽トロキャプチャ環境
- Python 3.x

Windows11での偽トロキャプチャのセットアップ例:
[https://note.com/wabecch2/n/nbd2aab50d21f?sub_rt=share_pb](https://note.com/wabecch2/n/nbd2aab50d21f?sub_rt=share_pb)

## 機能概要

- タイトル画像/QR画像をテンプレートマッチングで判定
- 直近10件の検知画像をプレビュー表示
- 出力結果をカンマ区切り1行で保持
- Gen7 Main​ RNG Tool への自動反映
- タイトルモード/QRモード 切り替え

## モード仕様

### 共通

- モードラジオボタン
- 監視開始 / 監視停止
- ステータス表示
- 出力結果テキスト
- 直近10件の検知画像
- 出力結果コピー / 出力結果クリア

### タイトルモード

- 検知インターバル(秒): デフォルト 1.0
- `監視停止時` に `停止時にGen7 Main RNG Toolへ連携` のチェックがONなら Gen7 Main RNG Tool の針リストへ自動出力

### QRモード

- `監視開始` ボタンで QR針を自動認識
- `出力` ボタンで Gen7 側の針リストへ反映

## 3DS Viewer 表示設定

- 転送モード: Light Weight Mode
- フィルター: No Filter
- 等倍調整: dot by dot x2
- 上下表示: 上下画面表示、上下比率 50%:50%

## 動作イメージ

![タイトル](resources/readme/title_preview.gif)

![QR](resources/readme/qr_preview.gif)