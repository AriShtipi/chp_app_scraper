import os

# ── env-controlled settings ──
VAT          = float(os.environ.get("VAT", 1.18))
DELAY_SEC    = float(os.environ.get("DELAY_SEC", 1.0))
BATCH_SIZE   = int(os.environ.get("BATCH_SIZE", 10))
TOP_N_LOCAL  = int(os.environ.get("TOP_N_LOCAL", 2))
TOP_N_ONLINE = int(os.environ.get("TOP_N_ONLINE", 1))

DEFAULT_CITIES = [
    {"name": "טירת כרמל ", "code1": "9000", "code2": "2100"},
    {"name": "אריאל ",     "code1": "9000", "code2": "3570"},
    {"name": "ביתר עילית ", "code1": "9000", "code2": "3780"},
]