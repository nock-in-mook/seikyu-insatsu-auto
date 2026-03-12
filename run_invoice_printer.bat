@echo off
chcp 65001 >nul
set PYTHONUTF8=1

rem 環境変数を読み込み（shared-envはbash形式なのでここで個別設定）
rem パスワードはこのファイルに直書きしない → shared-envからsourceされる前提
rem バッチから直接実行する場合は以下のコメントを外して設定:
rem set INVOICE_IMAP_PASS=ここにパスワード
rem set ANTHROPIC_API_KEY=ここにAPIキー
rem set INVOICE_SLACK_WEBHOOK=ここにWebhook URL

py -3.14 "%~dp0invoice_printer.py" %*
