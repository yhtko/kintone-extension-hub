📝 プロジェクト要件定義（MVP）

0) ゴール
•	Kintone 画面上に Excelライクな編集オーバーレイを被せ、文字列/数値/日付を編集 → 一括保存できるChrome拡張（MV3）。
•	どのアプリ・ビューでも起動可。カスタムビュー不要。Kintone API権限は既存ログインを使用。
________________________________________
1) スコープ（必須 / 非スコープ）

必須（MVP）
•	起動トリガ：Ctrl+Shift+E（Win/Mac両対応）でオーバーレイ表示／Escで閉じる
•	データ取得：
o	現在アプリID：kintone.app.getId()
o	現在の一覧条件：kintone.app.getQuery() を取得
o	フィールド定義：GET /k/v1/app/form/fields
o	レコード：GET /k/v1/records（fields: ['$id', …], totalCount:true）
•	対応フィールド型：SINGLE_LINE_TEXT / NUMBER / DATE（この3種のみ編集可）
•	UI：画面全面のオーバーレイ（body直下、Kintone DOM非依存）
o	列ヘッダ（A,B,C…）・行ヘッダ（1,2,3…）あり
o	グリッド本体は固定幅 160px/列（MVP）
o	スクロール同期（列ヘッダ/行ヘッダ）
•	編集：セル内インライン編集（input / type=date）
o	入力即時バリデーション：
	NUMBER：数値に変換可能か
	DATE：YYYY-MM-DD 形式
o	変更セルは淡黄ハイライト（dirty）
•	保存：PUT /k/v1/records へ 50件/バッチで一括送信
o	失敗時はそのバッチをエラー表示（行ID・フィールドコード）
o	成功した分はオーバーレイに反映、差分バッファをクリア
•	ショートカット：
o	Enter/Shift+Enter：下/上へ確定移動
o	Tab/Shift+Tab：右/左へ移動
o	Ctrl/⌘+S：保存
o	Esc：オーバーレイ閉じる
•	ステータス表示：右下に Sum / Avg / Count（dirtyセルの数値を対象に簡易集計）
•	未保存警告：閉じる/ESC時に未保存があれば確認ダイアログ

非スコープ（MVPではやらない）
•	ドロップダウン/チェックボックス/サブテーブル/添付/関連/計算の編集
•	クリップボード複数セル貼り付け、Undo/Redo、フィルハンドル、列固定/幅リサイズ、フィルタ行、仮想スクロール
•	同時更新の高度な競合解決（※MVPは保存失敗の通知のみ）
________________________________________
2) 技術構成
/kintone-excel-overlay
├─ manifest.json             // MV3
├─ content.js                // オーバーレイUI/編集/保存（content script）
├─ injected.js               // page context で kintone.api を橋渡し
├─ overlay.css               // Excel風スタイル
└─ icons/icon128.png
2.1 manifest.json（要件）
•	"manifest_version": 3
•	permissions: ["scripting","storage"]
•	host_permissions:
https://*.cybozu.com/*, https://*.kintone.com/*, https://*.cybozu.cn/*
•	content_scripts:
matches: https://*/k/*, https://*/k/#/* / run_at: document_idle
js: content.js, css: overlay.css
•	action: default_title: "Excelモード"
________________________________________
3) ページ連携方式
•	content script は kintone オブジェクトへ直接アクセス不可のため、injected.js を <script> でページに挿入。
•	postMessageブリッジで通信（双方向）：
o	送信：{ type: 'PB_TO_PAGE', id, action, payload }
o	応答：{ type: 'PB_FROM_PAGE', id, ok: true|false, result|error }
•	injected.js が受けた action に応じて kintone.api / kintone.app.* を実行。

必須アクション
•	getAppId → kintone.app.getId()
•	getQuery → kintone.app.getQuery()
•	getFormFields ({app}) → GET /k/v1/app/form/fields
•	getRecords ({payload}) → GET /k/v1/records
•	putRecords ({payload}) → PUT /k/v1/records
________________________________________
4) 画面仕様（UI/UX）

4.1 オーバーレイ構造
•	ルート：#pb-excel-overlay（position:fixed; inset:0; z-index:2147483647; background:rgba(255,255,255,.96)）
•	ツールバー：左にタイトル、右に「未保存: n」「保存」「閉じる」ボタン
•	グリッドラッパ：上＝列ヘッダ、左＝行ヘッダ、右下＝スクロール領域
•	ステータスバー：右下 Sum / Avg / Count
•	小トースト：右下に短時間表示

4.2 表示ルール
•	列は 対応3型のフィールドのみに限定。順序は form 定義順（MVP）
•	列ヘッダは A, B, C…（見た目用、内部はフィールドコードを保持）
•	行ヘッダは 1..N
•	セル高さ 24px、列幅 160px（MVP固定）
•	dirtyセルは背景#fff6d6、エラーは赤枠＆titleで理由表示

4.3 操作
•	クリックでセル内 <input> フォーカス
•	IME対応：compositionstart/endでEnter処理抑止（確定前Enterは移動させない）
•	ショートカットはフォーカスがオーバーレイ内のときのみ有効
________________________________________
5) データ処理

5.1 フィールド抽出
•	/k/v1/app/form/fields の properties から
type in ['SINGLE_LINE_TEXT','NUMBER','DATE'] を抽出 → {code, label, type}

5.2 レコード取得
•	GET /k/v1/records 引数：
o	app: appId
o	query: kintone.app.getQuery()（MVPはそのまま使用）
o	fields: ['$id', ...fieldCodes]
o	totalCount: true

5.3 差分バッファ
•	Map($id → { fieldCode: { value } })
•	入力値と初期値を比較してdirty判定
•	未保存: n は diff中の全セル数 を集計表示

5.4 保存
•	PUT /k/v1/records に records: [{ id, record }, …]
•	50件/バッチで分割送信
•	成功後：
o	records ローカル状態に値反映
o	該当セルの dirty を解除
o	差分バッファをクリア
•	失敗時：アラート or 行単位エラートースト（MVPは簡易通知でOK）
________________________________________
6) バリデーション & エラーハンドリング
•	NUMBER：Number(v) が NaN ならエラー
•	DATE：/^\d{4}-\d{2}-\d{2}$/ で検査（フォーマット不一致ならエラー）
•	PUT 失敗：
o	HTTPエラー時：内容をユーザーに表示（title/alert）
o	レート制限などはMVPでは再試行なし（将来：指数バックオフ）
•	閉じる時：差分があれば confirm で確認
________________________________________
7) パフォーマンス
•	初回は最大100件を想定（MVPはページングなしでもOK）
将来：仮想スクロール/追加読込
•	DOM描画は文字列連結で tbody に一括挿入
•	スクロール同期は scroll イベントでヘッダへ反映
________________________________________
8) 権限・安全
•	Kintoneのアクセス権に準拠（更新不可フィールドは PUT で弾かれる → エラー表示）
•	拡張の host_permissions は kintone ドメインのみに限定
•	収集・外部送信なし（MVP）。ストレージ利用は不要（任意）
________________________________________
9) 受け入れ基準（Acceptance Criteria）
1.	Ctrl+Shift+E でオーバーレイが表示/Escで閉じる
2.	SINGLE_LINE_TEXT/NUMBER/DATE のセルが編集できる
3.	無効な数値/日付は即時に赤枠・タイトルで警告され、保存時にエラーとして検知される
4.	変更セルに淡黄ハイライトが付き、ツールバーに 「未保存: n」 が正しく表示される
5.	Ctrl/⌘+S で最大50行/バッチに分割して PUT され、成功時に dirty が解除される
6.	失敗時、ユーザーにエラーが分かる（最低限 alert で OK）
7.	スクロールしても列ヘッダ/行ヘッダは同期して追従する
8.	Sum/Avg/Count が dirty な数値セルを対象に計算され表示される
9.	Kintone のどのアプリの画面でも起動できる（一覧条件が反映されたデータが出る）
________________________________________
10) テストケース（抜粋）
•	文字列/数値/日付をそれぞれ編集→保存→一覧再読込で反映確認
•	不正な日付（2025/01/01）でエラー表示→保存不可
•	未保存: n が差分数に追従
•	51件更新時に 50 + 1 の2回PUTが呼ばれる
•	権限で更新不可のフィールドを編集→保存エラー（期待通り失敗）
•	Ctrl+S 保存、Esc 閉じる、Enter/Tab移動が機能
•	大量行（例:100件）でもスクロール・編集がストレスなく動作
________________________________________
11) 将来拡張（備忘）
•	貼り付け（TSV→範囲反映）、Undo/Redo、ドロップダウン・チェックボックス対応
•	列幅リサイズ・固定列、フィルタ行、フィルハンドル、仮想スクロール
•	競合検知の高度化（更新日時/リビジョン差分プレビュー）
•	保存時の指数バックオフ・失敗行だけ再試行
•	ページング/追加読込、表示列のユーザー選択
________________________________________
12) 実装の雛形（コード化ガイド）
•	content.js
o	injected.js を <script> で挿入
o	postMessageのPromiseラッパ（callPage(action,payload)）
o	オーバーレイDOM生成・表示制御
o	getFormFields → getRecords で描画
o	入力→差分バッファ→PUT /records (50件バッチ)
o	キーボードショートカット
•	injected.js
o	PB_TO_PAGE を受信→kintone.app.*/kintone.api 呼び出し→PB_FROM_PAGE で返却
•	overlay.css
o	オーバーレイ最上位 / ツールバー / 列・行ヘッダ / セル（dirty/sel） / ステータス / トースト
