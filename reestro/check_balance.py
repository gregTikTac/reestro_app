import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from reestro_parser import ReestroClient, load_config

cfg = load_config(Path(__file__).parent / "config.json")
client = ReestroClient(cfg, pause=0)
url = f"{cfg['baseUrl'].rstrip('/')}/realty/billing/v1/balance"
r = client.session.get(url, timeout=60)
print("balance HTTP", r.status_code)
print(r.text[:500])
