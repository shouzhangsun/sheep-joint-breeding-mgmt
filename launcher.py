# -*- coding: utf-8 -*-
"""
山羊联合育种管理系统 — Windows 桌面端启动器
职责：
  1. 解析本机可写数据目录（%APPDATA%/SheepBreeding）
  2. 首次启动时从内置 seed 空库复制 sheep_farm.db 与创建 uploads
  3. 在 127.0.0.1 随机空闲端口后台线程运行 Flask（无登录、无 debug）
  4. 用 pywebview 打开原生窗口加载本地页面
  5. 单实例锁：已运行时退出（避免重复启动）
  6. 窗口关闭即退出进程（Flask 线程为守护线程）
"""
import os
import sys
import time
import socket
import threading
import webview  # pywebview 原生窗口能力（选择文件、关闭程序）

# ── 1. 数据目录（必须在 import app 之前设置，app 在导入时读取 SHEEP_DATA_DIR）──
APPDATA = os.environ.get("APPDATA") or os.path.expanduser("~")
DATA_DIR = os.path.join(APPDATA, "SheepBreeding")
os.makedirs(DATA_DIR, exist_ok=True)
os.environ["SHEEP_DATA_DIR"] = DATA_DIR

APP_DIR = os.path.dirname(os.path.abspath(__file__))
if getattr(sys, "frozen", False):
    RES_BASE = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
else:
    RES_BASE = APP_DIR
SEED_DB = os.path.join(RES_BASE, "seed", "sheep_farm.db")
DB_PATH = os.path.join(DATA_DIR, "sheep_farm.db")
UPLOAD_DIR = os.path.join(DATA_DIR, "uploads")
LOCK_FILE = os.path.join(DATA_DIR, "app.lock")

APP_NAME = "山羊联合育种管理系统"


def log(msg):
    print(f"[{APP_NAME}] {msg}", flush=True)


def log_error(where, exc):
    """把启动期异常写入日志文件，方便最终用户排查（即使无控制台也可见）。"""
    try:
        import traceback
        log_dir = os.environ.get("SHEEP_DATA_DIR", DATA_DIR)
        os.makedirs(log_dir, exist_ok=True)
        with open(os.path.join(log_dir, "launcher_error.log"), "a", encoding="utf-8") as f:
            f.write(f"\n[{datetime_now()}] ERROR in {where}:\n")
            f.write(traceback.format_exc())
    except Exception:
        pass


def datetime_now():
    try:
        return __import__("datetime").datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return "?"


class Api:
    """暴露给前端（pywebview JS 桥）调用的原生能力。"""

    def select_excel(self):
        """打开 Windows 原生文件选择框，返回选中文件路径列表；取消则为 None。"""
        try:
            win = webview.active_window()
            if not win:
                return None
            result = win.create_file_dialog(
                webview.OPEN_DIALOG,
                allow_multiple=False,
                file_types=("Excel 工作簿 (*.xlsx)", "*.xlsx"),
            )
            return result  # 列表（如 ['C:\\...\\a.xlsx']）或 None
        except Exception as e:
            log_error("select_excel", e)
            return None

    def select_any(self, allowed=("Excel 工作簿 (*.xlsx)", "*.xlsx")):
        """通用原生文件选择（预留给其他导入/导出场景）。"""
        try:
            win = webview.active_window()
            if not win:
                return None
            return win.create_file_dialog(
                webview.OPEN_DIALOG, allow_multiple=False, file_types=allowed
            )
        except Exception as e:
            log_error("select_any", e)
            return None

    def close_app(self):
        """关闭窗口并退出程序（Flask 为守护线程，会随进程结束）。"""
        try:
            win = webview.active_window()
            if win:
                win.destroy()
        except Exception as e:
            log_error("close_app", e)


# ── 2. 单实例锁 ──
def _pid_alive(pid):
    if sys.platform.startswith("win"):
        import subprocess
        out = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}"],
            capture_output=True, text=True,
        ).stdout
        return str(pid) in out
    else:
        try:
            os.kill(pid, 0)
            return True
        except OSError:
            return False


def acquire_lock():
    """返回 True 表示本进程获得锁；False 表示已有实例在运行。"""
    if os.path.exists(LOCK_FILE):
        try:
            with open(LOCK_FILE, "r") as f:
                old_pid = int(f.read().strip())
            if _pid_alive(old_pid):
                log(f"检测到已在运行的实例 (PID={old_pid})，退出。")
                return False
        except Exception:
            pass
    with open(LOCK_FILE, "w") as f:
        f.write(str(os.getpid()))
    return True


def release_lock():
    try:
        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)
    except Exception:
        pass


# ── 3. 首次运行：从 seed 复制空库 ──
def ensure_db():
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    if not os.path.exists(DB_PATH):
        if os.path.exists(SEED_DB):
            import shutil
            shutil.copyfile(SEED_DB, DB_PATH)
            log(f"已从内置模板创建数据库：{DB_PATH}")
        else:
            log("警告：未找到内置 seed 数据库，将尝试自动建表。")
    else:
        log(f"使用现有数据库：{DB_PATH}")


# ── 4. 随机空闲端口 ──
def find_free_port():
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    return port


# ── 5. 启动 Flask（后台线程）──
def start_flask(port):
    try:
        import app  # 必须在设置 SHEEP_DATA_DIR 后导入
        app.app.run(
            host="127.0.0.1",
            port=port,
            debug=False,
            use_reloader=False,
            threaded=True,
        )
    except Exception as e:
        log_error("start_flask", e)
        raise


def wait_for_server(port, timeout=30):
    import urllib.request
    url = f"http://127.0.0.1:{port}/"
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            urllib.request.urlopen(url, timeout=2)
            return True
        except Exception:
            time.sleep(0.3)
    return False


def main():
    if not acquire_lock():
        # 已有实例：尝试聚焦（best-effort），随后退出（使用模块顶部已导入的 webview）
        try:
            webview.create_window(APP_NAME + "（已在运行）",
                                  html="<h3 style='font-family:sans-serif;padding:20px'>程序已在运行，请勿重复打开。</h3>")
            webview.start()
        except Exception:
            pass
        sys.exit(0)

    try:
        ensure_db()
        port = find_free_port()
        log(f"Flask 将在 127.0.0.1:{port} 启动")

        flask_thread = threading.Thread(
            target=start_flask, args=(port,), daemon=True
        )
        flask_thread.start()

        if not wait_for_server(port):
            log("错误：本地服务启动超时。")
            release_lock()
            sys.exit(1)

        log("本地服务就绪，正在打开窗口…")
        api = Api()
        window = webview.create_window(
            APP_NAME,
            f"http://127.0.0.1:{port}/",
            width=1366,
            height=800,
            min_size=(1024, 700),
        )
        # 暴露 API 到 JS（通过 window.expose 而非 create_window 参数）
        window.expose(api.select_excel)
        window.expose(api.select_any)
        window.expose(api.close_app)
        webview.start()
    finally:
        release_lock()


if __name__ == "__main__":
    main()
