# 引き継ぎメモ

## 現在の状況
- 請求書自動印刷スクリプト完成・動作確認済み
- タスクスケジューラに毎月23日 9:00で登録済み
- 本番実行テスト済み（3月分の請求書3ファイル印刷成功）

## 構成
- `invoice_printer.py` — メインスクリプト（IMAP取得→Claude API判定→PDF印刷→Slack通知）
- `load_env_and_run.py` — タスクスケジューラ用の環境変数ローダー
- `run_invoice_printer.bat` — バッチファイル（タスクスケジューラから呼ばれる）
- `SumatraPDF/` — ポータブル版PDF印刷ツール
- `processed.json` — 処理済みメール記録（二重印刷防止）

## 主要な仕様
- 対象期間: 当月1日〜実行日
- 全メールをClaude APIで判定（取引先マッピングをプロンプトに含む）
- 必須取引先: (有)ケーイング（yabu@k--ing.com）、合同会社トーシュー（sato@toshu-llc.co.jp）
- 未着チェック: 必須取引先からメールがなければ警告
- 印刷: SumatraPDF + win32printでスプーラ監視（用紙切れ等検知）
- Slack通知: 会社別ファイル一覧、スキップ、エラー、未着警告
- プリンター: Brother DCP-J528N Printer (2 コピー)
- `--dry-run` で印刷せずテスト可能

## 環境変数（shared-env）
- ANTHROPIC_API_KEY
- INVOICE_IMAP_PASS
- INVOICE_SLACK_WEBHOOK

## 次のアクション
- 来月23日の自動実行を待って動作確認
- 取引先が増えたらREQUIRED_SENDERSに追加
