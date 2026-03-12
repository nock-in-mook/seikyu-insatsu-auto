"""shared-envから環境変数を読み込んでinvoice_printer.pyを実行する"""
import os
import sys
import re
from pathlib import Path

# shared-envのパス
SHARED_ENV = Path("G:/マイドライブ/_claude-sync/shared-env")

def load_shared_env():
    """bash形式のshared-envから環境変数を読み込む"""
    if not SHARED_ENV.exists():
        print(f"警告: shared-envが見つかりません: {SHARED_ENV}")
        return
    with open(SHARED_ENV, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            # "export KEY=VALUE" または "KEY=VALUE" をパース
            line = re.sub(r"^export\s+", "", line)
            match = re.match(r'([A-Za-z_][A-Za-z0-9_]*)=(.*)', line)
            if match:
                key = match.group(1)
                val = match.group(2).strip().strip('"').strip("'")
                # インラインコメントを除去
                val = re.split(r'\s+#\s+', val)[0].strip().strip('"').strip("'")
                os.environ[key] = val

if __name__ == "__main__":
    load_shared_env()
    # invoice_printer.pyを実行
    script = Path(__file__).parent / "invoice_printer.py"
    sys.argv[0] = str(script)
    exec(open(script, encoding="utf-8").read())
