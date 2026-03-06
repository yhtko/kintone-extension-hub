# kintone Excel Overlay – Project Summary

## 概要

kintone のレコード一覧を **Excelのように直接編集できるオーバーレイUI** を提供するブラウザ拡張機能。

目的は

* kintone の一覧編集の生産性向上
* Excel感覚のデータ入力
* 大量データの高速編集

---

# 現在の機能

## 1. Excel風グリッドUI

オーバーレイ上に以下の構造のグリッドを表示

```
Overlay
 ├ Header (列名)
 ├ Body (スクロールグリッド)
 └ Status bar
```

特徴

* 仮想スクロール（Virtual Rows）
* 高速レンダリング
* 100件表示
* 横スクロール対応
* 行番号固定
* ヘッダー固定

---

# 2. 編集機能

対応フィールド

| type             | 状態     |
| ---------------- | ------ |
| SINGLE_LINE_TEXT | 編集可能   |
| NUMBER           | 編集可能   |
| DATE             | 編集可能   |
| DROP_DOWN        | 編集可能   |
| RADIO_BUTTON     | 編集可能   |
| LOOKUPキー         | 編集可能   |
| SUBTABLE         | モーダル編集 |

---

## キーボード操作

| キー          | 動作   |
| ----------- | ---- |
| Tab         | 右移動  |
| Shift+Tab   | 左移動  |
| Enter       | 下移動  |
| Shift+Enter | 上移動  |
| Arrow Keys  | セル移動 |
| Ctrl+C      | コピー  |
| Ctrl+V      | ペースト |
| Ctrl+Z      | Undo |
| Ctrl+Y      | Redo |

特徴

* 矢印キーは **常にセル移動**
* スクロールには使わない
* Excelと同じ操作感

---

# 3. Undo / Redo

履歴スタックを持つ

```
historyStack
redoStack
```

対応

* セル変更
* ペースト
* 行追加
* 行削除

---

# 4. 差分管理

変更は即保存されず

```
dirtyRecords
```

に保持される。

UI表示

```
Unsaved: X
```

保存時

```
/k/v1/records.json
```

へまとめて送信

---

# 5. 行削除

削除は即削除ではなく

```
pending delete
```

として扱う

UI

* 赤背景
* 保存時削除

---

# 6. 権限対応

ユーザー権限を考慮

制御

| 権限   | 動作      |
| ---- | ------- |
| 閲覧のみ | 編集不可    |
| 編集可  | 編集可     |
| 削除不可 | 削除ボタン無効 |

実装

```
records/acl/evaluate
```

---

# 7. フィルター

列ヘッダーに

```
▼
```

を表示

機能

* 列フィルタ
* 並び替え

---

# 8. ステータスバー

下部に表示

```
Selected
Filtered
Pending delete
New rows
Sum
Avg
Count
```

---

# 9. サブテーブル編集

サブテーブルは

```
Edit subtable modal
```

で編集

改善

```
width: min(1100px, 90vw)
height: min(600px, 80vh)
```

---

# 10. セル種別アイコン

セル右端に小さく表示

| 種類           | 表示 |
| ------------ | -- |
| DROP_DOWN    | ▾  |
| RADIO_BUTTON | ◎  |
| LOOKUP       | ↗  |

実装

```
CSS pseudo element
::after
```

DOMを増やさないため高速

---

# 11. Copy / Paste

Excel互換

対応

* 複数セルコピー
* 複数セル貼り付け
* 範囲ペースト
* クリップボード対応

---

# 12. ページング

一覧の

```
Prev / Next
```

に対応

取得

```
limit 100
offset
```

---

# UI設計方針

目標

```
Excel感覚
+
kintone安全性
```

原則

* UIは静かに
* hover要素は最小
* DOMを増やさない
* CSS疑似要素優先
* 仮想スクロール必須

---

# パフォーマンス設計

採用技術

```
Virtual Rows
DOM reuse
Transform scroll sync
```

レンダリング

```
visible rows only
```

メリット

* 数千レコードでも軽量
* スクロール高速

---

# 現在の改善課題

## 1 Toastのクリッピング

トーストが

```
grid container overflow
```

で見切れる

対策

```
overlay root直下に配置
position: fixed
z-index: 9999
```

---

## 2 Subtable modal サイズ

現在小さい

改善

```
width: min(1100px, 90vw)
height: min(600px, 80vh)
```

---

## 3 Lookupフィールド検出

kintone lookup は

```
type: SINGLE_LINE_TEXT
```

だが

```
field.lookup
```

プロパティを持つ

判定

```
field.type === SINGLE_LINE_TEXT && field.lookup
```

---

# 将来拡張

候補

* 列幅リサイズ
* 列固定
* 列非表示保存
* セルコメント
* Excelインポート
* Excelエクスポート

---

# 技術構成

拡張

```
Chrome Extension
```

主要ファイル

```
content.js
overlay.css
page-bridge.js
```

通信

```
kintone REST API
```

---

# 現在の完成度

UI完成度

```
★★★★☆ (4/5)
```

実用性

```
高
```

kintoneの標準UIより

```
編集効率は大幅に向上
```

---

# プロジェクト状態

```
Production ready (beta)
```

現場テスト可能
