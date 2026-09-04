# -*- coding: utf-8 -*-
"""生成打包内置的空库 seed/sheep_farm.db：仅表结构 + 6 个品种，无 users 表、无业务数据。"""
import os
import sys
import tempfile
import shutil

APP_DIR = os.path.dirname(os.path.abspath(__file__))
SEED_DIR = os.path.join(APP_DIR, "seed")
os.makedirs(SEED_DIR, exist_ok=True)
SEED_DB = os.path.join(SEED_DIR, "sheep_farm.db")

# 让 app 导入时的自动 init_db 落在临时目录，避免污染本机数据目录
TMP_ENV = tempfile.mkdtemp(prefix="seedenv_")
os.environ["SHEEP_DATA_DIR"] = TMP_ENV

import app  # noqa: E402

app.DB_PATH = SEED_DB
app.UPLOAD_FOLDER = os.path.join(SEED_DIR, "uploads")
if os.path.exists(SEED_DB):
    os.remove(SEED_DB)

app.init_db()

# 确保 WAL 数据落盘并清理临时文件，得到干净的单文件库
import sqlite3
con = sqlite3.connect(SEED_DB)
con.execute("PRAGMA wal_checkpoint(TRUNCATE)")
con.close()
for ext in ("-wal", "-shm"):
    p = SEED_DB + ext
    if os.path.exists(p):
        os.remove(p)

con = sqlite3.connect(SEED_DB)
sheep_n = con.execute("SELECT COUNT(*) FROM sheep").fetchone()[0]
breeds = [r[0] for r in con.execute("SELECT name FROM breeds ORDER BY id")]
users_n = con.execute("SELECT COUNT(*) FROM sqlite_master WHERE name='users'").fetchone()[0]
con.close()

print("seed sheep(应为0):", sheep_n)
print("seed breeds:", breeds)
print("seed users 表是否存在(应为0):", users_n)
print("SEED OK ->", SEED_DB)

shutil.rmtree(TMP_ENV, ignore_errors=True)
