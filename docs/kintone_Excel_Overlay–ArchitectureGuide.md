#

このドキュメントは
**新しい開発者 / AI が5分でプロジェクトを理解するための設計図**です。

---

# 1. プロジェクト概要

kintone のレコード一覧を **Excel風に直接編集できる Overlay UI** を提供するブラウザ拡張。

目的

* kintone一覧の編集効率を上げる
* Excel操作に近づける
* 大量データ編集を高速化

特徴

```
Excel UX
+
kintone API
+
Virtual Grid
```

---

# 2. UI構造

Overlay構造

```
pb-overlay
 ├ overlay-header
 │   ├ Prev
 │   ├ page info
 │   ├ Next
 │   ├ Columns
 │   ├ Unsaved count
 │   ├ Add row
 │   ├ Undo
 │   ├ Redo
 │   ├ Save
 │   └ Close
 │
 ├ grid-head
 │   └ column headers
 │
 ├ grid-body
 │   ├ row headers
 │   └ cell grid
 │
 ├ status-bar
 │
 └ toast-container
```

---

# 3. Virtual Grid

大量データ対応のため

```
Virtual Rendering
```

を採用

### 描画方式

```
visible rows only
```

スクロール

```
scrollTop
↓
row index
↓
render rows
```

### 利点

* DOM削減
* 高速スクロール
* 大量レコード対応

---

# 4. データモデル

Overlay内部状態

```
state
```

主な構造

```
records
filteredRecords
dirtyRecords
pendingDelete
historyStack
redoStack
selection
editingCell
```

---

# 5. セル編集

編集は

```
input element
```

をセルに差し込む方式

```
display cell
↓
enter edit
↓
input render
↓
value commit
```

---

# 6. Undo / Redo

履歴管理

```
historyStack
redoStack
```

保存内容

```
{
  rowIndex
  fieldCode
  oldValue
  newValue
}
```

---

# 7. 保存フロー

変更は即保存しない

```
dirtyRecords
```

に保持

保存時

```
/k/v1/records.json
```

へ送信

---

# 8. 行削除

削除は

```
pendingDelete
```

として保持

UI

```
red row
```

保存時削除

---

# 9. コピー / ペースト

Excel互換

対応

```
Ctrl+C
Ctrl+V
```

複数セル対応

---

# 10. キーボード操作

```
Arrow keys → セル移動
Tab → 右
Shift+Tab → 左
Enter → 下
Shift+Enter → 上
```

スクロール用途では使用しない

---

# 11. セルタイプ

識別

| type         | 表示 |
| ------------ | -- |
| DROP_DOWN    | ▾  |
| RADIO_BUTTON | ◎  |
| LOOKUP       | ↗  |

LOOKUP判定

```
field.type === SINGLE_LINE_TEXT
AND
field.lookup
```

---

# 12. サブテーブル

サブテーブルは

```
modal editor
```

表示

サイズ

```
width: min(1100px, 90vw)
height: min(600px, 80vh)
```

---

# 13. フィルター

列ヘッダー

```
▼
```

機能

```
filter
sort
```

---

# 14. ステータスバー

表示

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

# 15. トースト

通知

```
Copied
Saved
Error
```

位置

```
overlay root
position: fixed
z-index: 9999
```

---

# 16. パフォーマンス設計

採用技術

```
Virtual DOM reuse
Transform scroll sync
CSS pseudo icons
```

DOM追加を最小化

---

# 17. 技術構成

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

# 18. 今後の改善

候補

```
column resize
column pin
column hide
excel import
excel export
```

---

# 19. 完成度

```
Production ready (beta)
```

実運用可能レベル
