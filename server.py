"""极简 Web 服务：浏览器上传文档 → 生成 PPTX → 直接下载。

只用标准库（不引入 Flask/FastAPI），与 CLI 共用同一条管线：
    python server.py [--host 127.0.0.1] [--port 8000] [--open]

转换是同步的；PowerPoint COM 校验/渲染不支持并发，因此用进程内互斥锁串行化。
运行日志写到项目外的 ../logs/server_runs.jsonl（不进入代码库）。
"""
import argparse
import json
import os
import re
import sys
import tempfile
import threading
import time
import traceback
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote

from doc2ppt import guard
from run import ROOT, _run_once

SUPPORTED = {".docx", ".pdf", ".md"}
PPTX_MIME = ("application/vnd.openxmlformats-officedocument."
             "presentationml.presentation")
MAX_BODY = 50 * 1024 * 1024   # 整个 multipart 请求
MAX_FILE = 30 * 1024 * 1024   # 单个文件
LOG_PATH = ROOT.parent / "logs" / "server_runs.jsonl"
_CONVERT_LOCK = threading.Lock()

_SAFE_CHARS = re.compile(r"[^\w\u4e00-\u9fff.\- ]+")
_WIN_RESERVED = {"CON", "PRN", "AUX", "NUL"} | {f"COM{i}" for i in range(1, 10)} \
    | {f"LPT{i}" for i in range(1, 10)}


def safe_name(name, allowed):
    """取安全的文件名：去目录、白名单字符、限长，并校验扩展名。"""
    base = Path(str(name or "").replace("\\", "/")).name.strip()
    base = _SAFE_CHARS.sub("_", base).strip(" .")[:90]
    ext = Path(base).suffix.lower()
    if ext not in allowed:
        raise ValueError(f"不支持的文件类型: {ext or '无扩展名'}"
                         f"（支持 {' / '.join(sorted(allowed))}）")
    stem = Path(base).stem.strip() or "document"
    if stem.upper() in _WIN_RESERVED:  # Windows 保留设备名，写盘会失败/被重定向
        stem = "_" + stem
    return stem + ext


def parse_multipart(content_type, body):
    """解析 multipart/form-data（Python 3.13 已移除 cgi，这里手写最小实现）。"""
    m = re.search(r'boundary=(?:"([^"]+)"|([^;]+))', content_type or "")
    if not m:
        raise ValueError("请求缺少 multipart boundary")
    boundary = ("--" + (m.group(1) or m.group(2)).strip()).encode()
    fields, files = {}, {}
    for chunk in body.split(boundary)[1:]:
        if chunk.startswith(b"--"):
            break
        if chunk.startswith(b"\r\n"):
            chunk = chunk[2:]
        head, sep, data = chunk.partition(b"\r\n\r\n")
        if not sep:
            continue
        if data.endswith(b"\r\n"):
            data = data[:-2]
        htxt = head.decode("utf-8", "replace")
        nm = re.search(r'\bname="([^"]*)"', htxt)
        if not nm:
            continue
        fn = re.search(r'\bfilename="([^"]*)"', htxt)
        if fn is None:
            fields[nm.group(1)] = data.decode("utf-8", "replace")
        else:
            files[nm.group(1)] = (fn.group(1), data)
    return fields, files


def _pages(raw):
    try:
        n = int(raw)
    except (TypeError, ValueError):
        return 8
    return min(max(n, 2), 30)


def convert_upload(doc_name, doc_bytes, tpl_name=None, tpl_bytes=None,
                   pages=8, use_llm=False):
    """把上传内容落到临时目录，跑完整管线，返回 (pptx 字节, 下载名, 统计)。"""
    if len(doc_bytes) > MAX_FILE:
        raise ValueError(f"文件超过 {MAX_FILE // 1024 // 1024}MB 限制")
    src_name = safe_name(doc_name, SUPPORTED)
    with tempfile.TemporaryDirectory(prefix="doc2ppt_") as td:
        src = Path(td) / src_name
        src.write_bytes(doc_bytes)
        tpl = ROOT / "templates" / "standard_16x9.pptx"
        if tpl_bytes:
            tpl_path = Path(td) / safe_name(tpl_name or "template.pptx", {".pptx"})
            tpl_path.write_bytes(tpl_bytes)
            guard.guard_ooxml(tpl_path)  # 模板同样是不可信输入
            tpl = tpl_path
        out = Path(td) / "out.pptx"
        stats = {}
        _blocks, items, rounds, _log, issues = _run_once(
            src, tpl, out, pages, preview=False, use_llm=use_llm, stats=stats)
        blob = out.read_bytes()
        meta = {"pages": len(items), "rounds": rounds, "issues": len(issues),
                "degraded": bool(stats.get("degraded"))}
        return blob, src.stem + "-输出.pptx", meta


def _log_run(ok, filename, meta, secs, ip):
    """尽力记录运行日志；日志目录不可写不影响转换。"""
    try:
        LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        row = {"ts": datetime.now().isoformat(timespec="seconds"), "ok": ok,
               "file": filename, "secs": round(secs, 1), "ip": ip, **meta}
        with LOG_PATH.open("a", encoding="utf-8") as f:
            f.write(json.dumps(row, ensure_ascii=False) + "\n")
    except OSError:
        pass


PAGE = """<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>文档转 PPT</title>
<style>
  :root{--blue:#1F4E79;--ink:#1a1a1a;--muted:#6b7280;--line:#dbe3ee;--bg:#f4f7fb}
  *{box-sizing:border-box}
  body{margin:0;font:15px/1.6 "Microsoft YaHei","Source Han Sans SC",system-ui,sans-serif;
       color:var(--ink);background:var(--bg);display:flex;min-height:100vh;align-items:center;
       justify-content:center;padding:24px}
  .card{width:100%;max-width:560px;background:#fff;border:1px solid var(--line);
        border-radius:14px;box-shadow:0 12px 40px rgba(31,78,121,.10);overflow:hidden}
  header{background:var(--blue);color:#fff;padding:26px 30px 22px}
  header h1{margin:0;font-size:22px;letter-spacing:.5px}
  header p{margin:6px 0 0;font-size:13px;opacity:.85}
  form{padding:24px 30px 28px}
  label{display:block;font-size:13px;color:var(--muted);margin:16px 0 6px}
  label:first-of-type{margin-top:0}
  input[type=file],input[type=number]{width:100%;padding:10px 12px;border:1px solid var(--line);
        border-radius:8px;background:#fbfdff;font-size:14px}
  .row{display:flex;gap:14px}.row>div{flex:1}
  .check{display:flex;align-items:center;gap:8px;margin-top:18px;color:var(--ink);font-size:14px}
  button{margin-top:22px;width:100%;padding:13px;border:0;border-radius:9px;background:var(--blue);
         color:#fff;font-size:16px;font-weight:600;cursor:pointer}
  button:disabled{opacity:.55;cursor:wait}
  #st{margin:14px 2px 0;font-size:13px;color:var(--muted);min-height:20px;word-break:break-all}
  #st.err{color:#b42318}
  .hint{margin-top:6px;font-size:12px;color:var(--muted)}
</style></head><body>
<div class="card">
  <header><h1>文档转 PPT</h1><p>docx / pdf（有文本层）/ md → 可编辑 .pptx</p></header>
  <form id="f">
    <label for="file">文档</label>
    <input id="file" name="file" type="file" accept=".docx,.pdf,.md" required>
    <div class="row">
      <div><label for="pages">期望页数</label>
        <input id="pages" name="pages" type="number" value="8" min="2" max="30"></div>
      <div><label for="template">模板（可选，.pptx）</label>
        <input id="template" name="template" type="file" accept=".pptx"></div>
    </div>
    <label class="check"><input type="checkbox" name="llm" value="1"> 使用大模型大纲（需服务端已配置 API Key）</label>
    <button id="btn" type="submit">转换并下载</button>
    <div id="st"></div>
    <div class="hint">转换通常 5–20 秒；单文件上限 30MB。生成结果为原生可编辑元素，非图片拼接。</div>
  </form>
</div>
<script>
const f=document.getElementById('f'),st=document.getElementById('st'),btn=document.getElementById('btn');
f.addEventListener('submit',async e=>{
  e.preventDefault();btn.disabled=true;st.className='';st.textContent='正在转换，请稍候…';
  try{
    const r=await fetch('/convert',{method:'POST',body:new FormData(f)});
    const ct=r.headers.get('content-type')||'';
    if(!r.ok||ct.includes('json')){
      let msg='转换失败';try{msg=(await r.json()).error||msg}catch(_){}
      throw new Error(msg);
    }
    const blob=await r.blob();
    const name=decodeURIComponent(r.headers.get('x-output-name')||'output.pptx');
    const degraded=r.headers.get('x-verify-degraded')==='1';
    const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;a.click();
    setTimeout(()=>URL.revokeObjectURL(a.href),4000);
    st.textContent='完成：'+name+'（已开始下载）'+(degraded?'　⚠ 本机校验降级：仅估算，未检查标点，建议复验':'');
    if(degraded) st.className='err';
  }catch(err){st.className='err';st.textContent='错误：'+err.message;}
  finally{btn.disabled=false;}
});
</script></body></html>""".encode()


class Handler(BaseHTTPRequestHandler):
    server_version = "doc2ppt/1.0"

    def do_GET(self):
        path = self.path.split("?", 1)[0]
        if path in ("/", "/index.html"):
            self._send(200, "text/html; charset=utf-8", PAGE)
        elif path == "/health":
            self._json(200, {"ok": True, "llm": bool(os.environ.get("DOC2PPT_API_KEY"))})
        elif path == "/favicon.ico":
            self._send(204, "image/x-icon", b"")
        else:
            self._json(404, {"error": "not found"})

    def do_POST(self):
        if self.path.split("?", 1)[0] != "/convert":
            self._json(404, {"error": "not found"})
            return
        name, t0 = "?", time.time()
        try:
            n = int(self.headers.get("Content-Length") or 0)
            if n <= 0:
                raise ValueError("请求为空")
            if n > MAX_BODY:
                raise ValueError("上传请求过大（整个表单上限 50MB）")
            fields, files = parse_multipart(self.headers.get("Content-Type", ""),
                                            self.rfile.read(n))
            if "file" not in files:
                raise ValueError("缺少上传文件")
            name, data = files["file"]
            if not data:
                raise ValueError("上传文件为空")
            tpl_name, tpl_data = files.get("template", (None, None))
            with _CONVERT_LOCK:  # COM 渲染不支持并发，串行执行
                blob, out_name, meta = convert_upload(
                    name, data, tpl_name, tpl_data,
                    _pages(fields.get("pages")),
                    fields.get("llm") in ("1", "on", "true"))
            extra = {
                "Content-Disposition":
                    f"attachment; filename=\"output.pptx\"; "
                    f"filename*=UTF-8''{quote(out_name)}",
                "X-Output-Name": quote(out_name),
            }
            if meta.get("degraded"):
                extra["X-Verify-Degraded"] = "1"
            self._send(200, PPTX_MIME, blob, extra)
            _log_run(True, name, meta, time.time() - t0, self.client_address[0])
        except (ValueError, SystemExit) as e:
            msg = (e.code if isinstance(e.code, str) else "转换失败") \
                if isinstance(e, SystemExit) else str(e)
            _log_run(False, name, {"error": msg}, time.time() - t0,
                     self.client_address[0])
            self._json(400, {"error": msg})
        except Exception as e:  # noqa: BLE001 - 服务端兜底：记日志，不向客户端泄漏堆栈
            traceback.print_exc()
            _log_run(False, name, {"error": str(e)}, time.time() - t0,
                     self.client_address[0])
            self._json(500, {"error": "服务器内部错误，请查看服务端日志"})

    def _send(self, code, ctype, body, extra=None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        try:
            self.wfile.write(body)
        except (BrokenPipeError, ConnectionResetError):
            pass  # 客户端提前断开（关页面/取消下载）不该刷 traceback

    def _json(self, code, payload):
        self._send(code, "application/json; charset=utf-8",
                   json.dumps(payload, ensure_ascii=False).encode("utf-8"))

    def log_message(self, fmt, *args):
        sys.stderr.write(f"[{self.log_date_time_string()}] {fmt % args}\n")


class _Server(ThreadingHTTPServer):
    # Windows 下 SO_REUSEADDR 允许抢绑已被监听的端口，会让「端口检测」失效；
    # 关闭后重复绑定会正常报错，bind_server 才能顺延。
    allow_reuse_address = False
    daemon_threads = True


def make_server(host="127.0.0.1", port=8000):
    return _Server((host, port), Handler)


def bind_server(host="127.0.0.1", port=8000, tries=20):
    """绑定端口；被占用时向后顺延。返回 (httpd, 实际端口, 是否顺延)。

    Windows 上端口被占用表现为 PermissionError(WinError 10013)，与 EADDRINUSE
    同样按“换一个端口”处理——否则用户只看到一屏 traceback。
    """
    last = None
    for p in range(port, port + tries):
        try:
            return make_server(host, p), p, p != port
        except OSError as e:
            last = e
    raise OSError(f"端口 {port}-{port + tries - 1} 均不可用: {last}")


def main(argv=None):
    ap = argparse.ArgumentParser(prog="server.py",
                                 description="文档转 PPT 的本地 Web 服务")
    ap.add_argument("--host", default="127.0.0.1",
                    help="监听地址（默认仅本机；对外服务需显式 0.0.0.0）")
    ap.add_argument("--port", type=int, default=8000,
                    help="监听端口（默认 8000；被占用时自动顺延）")
    ap.add_argument("--open", action="store_true", help="启动后自动打开浏览器")
    args = ap.parse_args(argv)
    if not 0 <= args.port <= 65535:
        print(f"启动失败: 端口 {args.port} 超出范围（0-65535）", file=sys.stderr)
        return 2
    if not (ROOT / "templates" / "standard_16x9.pptx").exists():
        print("启动失败: 缺少默认模板 templates/standard_16x9.pptx", file=sys.stderr)
        return 2
    try:
        httpd, port, shifted = bind_server(args.host, args.port)
    except OSError as e:
        print(f"启动失败: {e}", file=sys.stderr)
        print("提示: 用 --port 指定其他端口，或先关闭占用该端口的程序。", file=sys.stderr)
        return 2
    if shifted:
        print(f"提示: 端口 {args.port} 已被其他程序占用，已自动改用 {port}", flush=True)
    # 0.0.0.0/:: 是对外监听地址，浏览器不能直接访问，展示时换回本机地址
    shown = "127.0.0.1" if args.host in ("0.0.0.0", "::") else args.host
    url = f"http://{shown}:{port}/"
    print(f"服务已启动: {url}  （Ctrl+C 停止）", flush=True)
    if args.host not in ("127.0.0.1", "localhost", "::1"):
        print(f"注意: 正在对外监听 {args.host}，同一局域网的设备也可访问", flush=True)
    print("支持 .docx / .pdf（有文本层）/ .md；单文件上限 30MB", flush=True)
    if args.open:
        webbrowser.open(url)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止")
    finally:
        httpd.server_close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
