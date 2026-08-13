# -*- coding: utf-8 -*-
"""
OpenCode Usage Console · 一体化本地工具（前后端绑定）

一条命令启动：服务 + 浏览器自动打开前端页面。
前端可：点击更新（运行 scrape_usage.py，纯 HTTP 直调 _server API 爬取，速度快）、
输入账号密码、选择登录方式（GitHub/Google）与下载时间段（全部/指定/固定周期）、
开启自动更新（后台定时线程按设置周期自动爬取，即使前端页面关闭也持续生效）。

用法：
    python opencode_usage.py
    （自动打开 http://127.0.0.1:9901/）

API:
    GET  /index.html       前端页面（根路径 / 亦返回）
    GET  /api/excel          最新 Excel 文件字节
    GET  /api/status         Excel 信息（行数/修改时间）
    GET  /api/autofetch      自动更新配置
    POST /api/autofetch      {enabled, interval, unit, loginMethod, email, password,
                              timeMode, start, end, months} 保存自动更新配置
    POST /api/update         {loginMethod, email, password, start, end, months}
                             运行爬虫更新数据（SSE 流式返回爬虫进度）
"""
import json
import re
import subprocess
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

BASE = Path(__file__).parent
HTML_FILE = BASE / "index.html"
EXCEL_FILE = BASE / "opencode_usage_history.xlsx"
SCRAPE_SCRIPT = BASE / "scrape_usage.py"
CONFIG_FILE = BASE / "autofetch_config.json"
PORT = 9901

_update_lock = threading.Lock()

DEFAULT_AUTOFETCH = {
    "enabled": False,
    "interval": 30,
    "unit": "min",
    "loginMethod": "github",
    "workspace": "",
    "email": "",
    "password": "",
    "timeMode": "all",
    "start": "",
    "end": "",
    "months": 0,
    "incremental": False,
}


def build_scrape_cmd(login_method="github", email="", password="", workspace="", start="", end="", months=0,
                     incremental=False, force_login=False):
    """构造爬虫命令行（手动更新与后台自动更新共用）。

    默认优先复用缓存登录态（storage_state），不强制重新登录；仅当需要换账号或
    缓存失效需用账号密码重登时 force_login=True。email/password 仅作为凭据传入，
    供首次登录或失效重登使用，不因填了账号就强制重登。workspace 留空则交给爬虫自动获取。
    """
    cmd = ["python", "-u", str(SCRAPE_SCRIPT)]  # -u: 无缓冲输出，进度实时可读
    cmd += ["--login-method", login_method]
    if force_login:
        cmd += ["--force-login"]
    if workspace:
        cmd += ["--workspace", workspace]
    if email:
        cmd += ["--email", email]
    if password:
        cmd += ["--password", password]
    if start or end:
        # 指定了下载时间段 → 爬取所有月份后再按时间段过滤
        cmd += ["--months", "0"]
        if start:
            cmd += ["--start", start]
        if end:
            cmd += ["--end", end]
    else:
        # 未指定时间段：固定周期传 months，否则（全部数据）months=0 表示不过滤
        cmd += ["--months", str(months)]
    if incremental:
        cmd += ["--incremental"]
    return cmd


def load_autofetch_config() -> dict:
    """读取自动更新配置（本地文件持久化，非数据库）。"""
    cfg = dict(DEFAULT_AUTOFETCH)
    if CONFIG_FILE.exists():
        try:
            data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                cfg.update(data)
        except Exception:
            pass
    return cfg


def save_autofetch_config(cfg: dict) -> None:
    CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


def run_scraper_blocking(cmd) -> tuple:
    """阻塞式运行爬虫，返回 (returncode, 末尾日志)。"""
    proc = subprocess.Popen(
        cmd, cwd=str(BASE),
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        text=True, encoding="utf-8", errors="replace",
    )
    assert proc.stdout is not None
    log_buf = []
    for line in proc.stdout:
        line = line.rstrip()
        log_buf.append(line)
    proc.wait(timeout=1800)
    return proc.returncode, "\n".join(log_buf[-80:])


def autofetch_loop():
    """后台自动更新线程：按配置周期循环执行爬虫（前端页面关闭也持续生效）。"""
    while True:
        cfg = load_autofetch_config()
        if not cfg.get("enabled"):
            time.sleep(5)
            continue
        unit_sec = {"min": 60, "hour": 3600, "day": 86400}.get(cfg.get("unit", "min"), 60)
        interval = max(1, int(cfg.get("interval") or 30)) * unit_sec
        time.sleep(interval)
        # 到点：重新读取配置（期间可能被修改/关闭），且无更新任务进行中才执行
        cfg = load_autofetch_config()
        if not cfg.get("enabled"):
            continue
        if not _update_lock.acquire(blocking=False):
            print("[autofetch] 有更新任务进行中，跳过本轮")
            continue
        try:
            login_method = cfg.get("loginMethod") or "github"
            start = cfg.get("start") or ""
            end = cfg.get("end") or ""
            months = int(cfg.get("months") or 0)
            # 三层恢复由爬虫内部处理（缓存→关联账号→账号密码），此处以缓存登录态启动
            cmd = build_scrape_cmd(login_method=login_method, email=cfg.get("email") or "",
                                   password=cfg.get("password") or "", workspace=cfg.get("workspace") or "",
                                   start=start, end=end, months=months,
                                   incremental=bool(cfg.get("incremental")))
            print(f"[autofetch] 自动更新开始: {' '.join(cmd)}")
            rc, tail = run_scraper_blocking(cmd)
            print(f"[autofetch] 自动更新结束 rc={rc}\n{tail}")
        except Exception as e:
            print(f"[autofetch] 自动更新失败: {e}")
        finally:
            _update_lock.release()


def add_cors(handler):
    handler.send_header("Access-Control-Allow-Origin", "*")
    handler.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
    handler.send_header("Access-Control-Allow-Headers", "Content-Type")


class Handler(BaseHTTPRequestHandler):
    def _send(self, code, body: bytes, ctype="application/json; charset=utf-8"):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        add_cors(self)
        self.end_headers()
        self.wfile.write(body)

    def _send_json(self, code, obj):
        self._send(code, json.dumps(obj, ensure_ascii=False).encode())

    def log_message(self, fmt, *args):
        pass

    # ---- 预检 ----
    def do_OPTIONS(self):
        self.send_response(204)
        add_cors(self)
        self.end_headers()

    # ---- GET ----
    def do_GET(self):
        path = self.path.split("?")[0]
        if path in ("/", "/index.html"):
            if not HTML_FILE.exists():
                self._send(404, b"index.html not found")
                return
            self._send(200, HTML_FILE.read_bytes(), "text/html; charset=utf-8")
        elif path == "/favicon.ico":
            self._send(204, b"")
        elif path == "/api/excel":
            if not EXCEL_FILE.exists():
                self._send_json(404, {"ok": False, "error": "excel not found"})
                return
            data = EXCEL_FILE.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Length", str(len(data)))
            add_cors(self)
            self.end_headers()
            self.wfile.write(data)
        elif path == "/opencode-icon.svg":
            icon = BASE / "assets" / "opencode-icon.svg"
            if not icon.exists():
                self._send_json(404, {"ok": False, "error": "icon not found"})
                return
            self._send(200, icon.read_bytes(), "image/svg+xml")
        elif path == "/api/status":
            info = {"exists": EXCEL_FILE.exists()}
            if info["exists"]:
                info["size"] = EXCEL_FILE.stat().st_size
                info["mtime"] = EXCEL_FILE.stat().st_mtime
                try:
                    import openpyxl
                    wb = openpyxl.load_workbook(EXCEL_FILE, read_only=True)
                    ws = wb.active
                    info["rows"] = max(ws.max_row - 1, 0)
                    wb.close()
                except Exception as e:
                    info["rows_error"] = str(e)
            self._send_json(200, info)
        elif path == "/api/autofetch":
            cfg = load_autofetch_config()
            cfg.pop("password", None)  # 不返回密码明文
            self._send_json(200, cfg)
        else:
            self._send_json(404, {"ok": False, "error": "not found"})

    # ---- POST ----
    def do_POST(self):
        if self.path == "/api/update":
            self._run_update()
        elif self.path == "/api/autofetch":
            self._set_autofetch()
        else:
            self._send_json(404, {"ok": False, "error": "not found"})

    def _read_json_body(self):
        try:
            length = int(self.headers.get("Content-Length") or 0)
            if length <= 0:
                return {}
            return json.loads(self.rfile.read(length).decode("utf-8"))
        except Exception:
            return {}

    def _set_autofetch(self):
        body = self._read_json_body()
        cfg = load_autofetch_config()
        for key in DEFAULT_AUTOFETCH:
            if key in body:
                cfg[key] = body[key]
        # 账号密码持久化到本地配置文件（非代码、非数据库）：后端自动更新在缓存登录态
        # 失效时可用其重新登录，前端页面关闭后自动更新仍可持续生效
        save_autofetch_config(cfg)
        state = "开启" if cfg.get("enabled") else "关闭"
        unit_sec = {"min": "分钟", "hour": "小时", "day": "天"}.get(cfg.get("unit"), "分钟")
        self._send_json(200, {"ok": True, "enabled": cfg.get("enabled"),
                              "desc": f"自动更新已{state}（每 {cfg.get('interval')} {unit_sec} 一次）"})

    def _run_update(self):
        if not SCRAPE_SCRIPT.exists():
            self._send_json(500, {"ok": False, "error": "scrape_usage.py not found"})
            return
        if not _update_lock.acquire(blocking=False):
            self._send_json(409, {"ok": False, "error": "已有更新任务在运行，请稍候"})
            return
        body = self._read_json_body()
        email = (body.get("email") or "").strip()
        password = (body.get("password") or "").strip()
        workspace = (body.get("workspace") or "").strip()
        start = (body.get("start") or "").strip()
        end = (body.get("end") or "").strip()
        login_method = (body.get("loginMethod") or "github").strip() or "github"
        months = int(body.get("months") or 0)
        incremental = bool(body.get("incremental"))
        cmd = build_scrape_cmd(login_method=login_method, email=email, password=password,
                               workspace=workspace, start=start, end=end, months=months,
                               incremental=incremental)

        # SSE 流式响应：爬虫 stdout 逐行实时推送给前端
        self.send_response(200)
        self.send_header("Content-Type", "text/event-stream; charset=utf-8")
        self.send_header("Cache-Control", "no-cache")
        self.send_header("Connection", "keep-alive")
        add_cors(self)
        self.end_headers()

        def emit(event_type, **fields):
            try:
                payload = json.dumps({"type": event_type, **fields}, ensure_ascii=False)
                self.wfile.write(("data: " + payload + "\n\n").encode("utf-8"))
                self.wfile.flush()
            except Exception:
                pass

        def _stream_scraper(cmd, emit):
            """运行爬虫并逐行 SSE 推送 stdout，返回 (returncode, 完整日志)。"""
            proc = subprocess.Popen(
                cmd, cwd=str(BASE),
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, encoding="utf-8", errors="replace",
            )
            assert proc.stdout is not None
            log_buf = []
            for line in proc.stdout:
                line = line.rstrip()
                log_buf.append(line)
                if line:
                    emit("line", text=line)
            proc.wait(timeout=1800)
            return proc.returncode, "\n".join(log_buf)

        try:
            # 三层恢复由爬虫内部处理（缓存→关联账号→账号密码），此处仅以缓存登录态启动一次
            cmd = build_scrape_cmd(login_method=login_method, email=email, password=password,
                                   workspace=workspace, start=start, end=end, months=months,
                                   incremental=incremental)
            emit("start", cmd=" ".join(cmd))
            rc, full = _stream_scraper(cmd, emit)
            m = re.search(r"总计记录[:：]\s*(\d+)", full)
            rows = int(m.group(1)) if m else None
            emit("done", ok=rc == 0, returncode=rc, rows=rows, log=full[-2500:])
        except subprocess.TimeoutExpired:
            emit("done", ok=False, error="更新超时（超过 30 分钟）")
        except Exception as e:
            emit("done", ok=False, error=str(e))
        finally:
            _update_lock.release()
            # 关键：显式关闭 SSE 连接，否则前端 reader.read() 永不结束（done 事件后仍阻塞）
            try:
                self.wfile.flush()
                self.close_connection = True
            except Exception:
                pass


if __name__ == "__main__":
    print("=" * 60)
    print("  OpenCode Usage Console · 一体化工具")
    print(f"  页面地址 : http://127.0.0.1:{PORT}/")
    print("  功能     : 前端点更新 / 输入账号密码 / 选择下载时间段 / 自动更新")
    print("  停止服务 : 关闭本窗口或 Ctrl + C")
    print("=" * 60)
    # 启动后台自动更新线程（前端页面关闭也持续生效）
    threading.Thread(target=autofetch_loop, daemon=True).start()
    server = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    threading.Timer(1.0, lambda: webbrowser.open(f"http://127.0.0.1:{PORT}/")).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n服务已停止")
