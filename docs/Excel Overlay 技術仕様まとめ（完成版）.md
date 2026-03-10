PlugBits Launcher for kintone
Excel Overlay 技術仕様まとめ（完成版）

Version: MVP完成版
対象: PlugBits Launcher Chrome Extension
Overlay機能: 一覧 + 詳細対応

1. 機能概要

Excel Overlay は kintone のレコード一覧および詳細画面を
スプレッドシート形式で編集できる UI レイヤーを提供する機能。

特徴

Excel風セル編集

コピー / ペースト

複数行編集

列並び替え保存

列幅保存

Grid / Form 表示切替

詳細画面の1レコード編集

サブテーブル編集

ビュー準拠表示

Pro制御

2. 起動方法

Overlayは以下の3方法で起動できる。

1 サイドパネルボタン

PlugBits Launcher SidePanel

[フィルター] [Excel Overlay] [設定]

Overlayボタンを押すと起動。

2 キーボードショートカット
Ctrl + Shift + E

動作

画面	挙動
一覧	一覧Overlay
詳細	1レコードOverlay
3 Pro制御

Overlay編集は Proのみ

非Pro

Overlay起動可能

編集不可（閲覧のみ）

3. Overlayモード

Overlayは2種類ある。

3.1 一覧Overlay

対象

kintone.app.getId()

表示

複数レコード

レイアウト

┌─────────────────────────────┐
│ Header                      │
│ Grid / Form / Save / Close  │
├─────────────────────────────┤
│ Spreadsheet Grid            │
│                             │
│ Rows                        │
│                             │
├─────────────────────────────┤
│ Footer (統計 / 件数)         │
└─────────────────────────────┘
3.2 詳細Overlay

対象

kintone.app.record.detail

表示

1レコード

レイアウト

┌───────────────┐
│ Field | Value │
│               │
│ Field | Value │
│               │
└───────────────┘

特徴

縦表示

Grid切替可能

4. 表示モード

Overlayは2モードを持つ

4.1 Grid

Excel形式

| A | B | C |
|---|---|---|
|   |   |   |

用途

大量編集

コピー

ペースト

行追加

4.2 Form

詳細型

Field | Value

用途

詳細確認

長文編集

5. ビュー対応

一覧Overlayは ビューを尊重する

表示

ビューの列のみ

切替

ビュー / 全項目
6. 編集対象フィールド

対応フィールド

フィールド	対応
Single line text	編集可
Number	編集可
Date	編集可
Dropdown	編集可
Radio	編集可
Link	編集可
Text area	編集可
Checkbox	編集可
Multi choice	編集可
Subtable	編集可
7. 編集制御

編集不可

フィールド
Lookup
System
Status
Calculated

表示のみ

8. リッチテキスト

仕様

閲覧のみ

理由

HTML

安全性

複雑編集

9. サブテーブル

サブテーブルは

モーダル編集

表示

┌───────────────┐
│ Subtable      │
│               │
│ Row Editor    │
│               │
└───────────────┘

操作

行追加

行削除

自動幅調整

10. 列並び替え

ユーザーは列順を変更できる

UI

列順

ドラッグ

保存

localStorage

キー

pbColumnLayout
11. 列幅保存

保存場所

localStorage

キー

pbColumnWidth

設定画面で管理

12. 設定画面

設定画面で

列レイアウト管理

表示

App / View

操作

ボタン	動作
並び解除	order reset
幅解除	width reset
削除	layout delete
13. フィルター

Overlayは

ビュー絞り込み

を継承する

保存後

フィルター再適用
14. 保存

保存方式

差分保存

管理

dirtyRecords

保存API

kintone REST API
15. コピー / ペースト

Excel互換

対応

Ctrl+C
Ctrl+V

複数セル

TSV
16. キーボード操作

ナビゲーション

キー	動作
Arrow	セル移動
Enter	編集確定
Tab	次セル

編集時

矢印はカーソル移動
17. ESC動作

ESC挙動

優先順位

1

サブテーブルモーダル

閉じる

2

Overlay

閉じる

18. Pro制御

Proチェック

pro-service.js

優先順

1

Developer override

2

extension installType = development

3

production entitlement
19. Developer override

設定画面

?dev=1

キー

pbDeveloperProOverride

保存

chrome.storage.local
20. 保存後再描画

保存成功後

1 dirty reset
2 source rows update
3 filter reapply
4 pagination rebuild
5 grid render
21. UI構成

主要ファイル

content.js
core.js
launcher.js
overlay.css
22. ファイル構成
manifest.json
service_worker.js
content.js
core.js
launcher.js
launcher.css
overlay.css
options.html
options.js
options.css
sidepanel.html
pins.js
permission-service.js
page-bridge.js
pro-service.js
vendor/lucide.umd.js
23. アーキテクチャ

構造

Chrome Extension
      │
SidePanel UI
      │
Content Script
      │
Overlay Engine
      │
kintone REST API
24. パフォーマンス

最適化

diff save

local layout cache

minimal DOM update

25. 今後の拡張

可能な拡張

RichText editor

将来

CSV import
Excel export
Column freeze
Undo history
26. プロジェクト状態
Overlay MVP 完成

状態

Production ready