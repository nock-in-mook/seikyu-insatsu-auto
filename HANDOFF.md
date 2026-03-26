# 引き継ぎメモ

## 現在の状況
- 請求書自動印刷スクリプト完成・タスクスケジューラからの本番実行テスト成功
- 3月23日のスケジューラ失敗を調査・修正済み
- Claude API → Gemini 2.5 Flash に移行済み（無料枠活用）

## 3月23日失敗の原因と修正
1. **Anthropic APIクレジット不足** → Gemini Flashに移行して解決
2. **API判定エラー時のKeyError: 'reason'** → `.get()` で安全にアクセスするよう修正
3. **タスクスケジューラの日本語パス問題** → `C:\Users\msp\invoice_printer_launcher.py` でラップして解決
4. **dry-runがprocessed.jsonに記録してしまう** → dry-run時は保存しないよう修正

## 構成
- `invoice_printer.py` — メインスクリプト（IMAP取得→Gemini Flash判定→PDF印刷→Slack通知）
- `load_env_and_run.py` — 環境変数ローダー
- `run_invoice_printer.bat` → `C:\Users\msp\invoice_printer_launcher.bat` → `.py` — タスクスケジューラ起動チェーン
- `C:\Users\msp\invoice_printer_launcher.py` — 日本語パス回避用ランチャー（ログ出力・GDrive待機付き）
- `run_log.txt` — 実行ログ
- `SumatraPDF/` — ポータブル版PDF印刷ツール
- `processed.json` — 処理済みメール記録（二重印刷防止）

## 主要な仕様
- 対象期間: 当月1日〜実行日
- 全メールをGemini 2.5 Flash APIで判定（取引先マッピングをプロンプトに含む）
- 必須取引先: (有)ケーイング（yabu@k--ing.com）、合同会社トーシュー（sato@toshu-llc.co.jp）
- 未着チェック: 必須取引先からメールがなければ警告
- 印刷: SumatraPDF + win32printでスプーラ監視（用紙切れ等検知）
- Slack通知: 会社別ファイル一覧、スキップ、エラー、未着警告
- プリンター: Brother DCP-J528N Printer (2 コピー)
- `--dry-run` で印刷せずテスト可能（processed.jsonに記録しない）

## 環境変数（shared-env）
- GEMINI_API_KEY（旧ANTHROPIC_API_KEYは不要）
- INVOICE_IMAP_PASS
- INVOICE_SLACK_WEBHOOK

## 次のアクション
- 4月23日の自動実行を待って動作確認
- 取引先が増えたらREQUIRED_SENDERSに追加
