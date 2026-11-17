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
from doc2ppt.template import template_path
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


# 转换期间的通知：单槽环形缓冲，按运行 id 轮询取回写给前端
_PROGRESS = {}          # run_id -> [{"t", "phase", "msg"}]
_PROGRESS_LOCK = threading.Lock()
_RUN_SEQ = 0


def _progress_id():
    global _RUN_SEQ
    with _PROGRESS_LOCK:
        _RUN_SEQ += 1
        return f"run-{time.time_ns()}-{_RUN_SEQ}"


def _push_progress(run_id, phase, msg):
    with _PROGRESS_LOCK:
        buf = _PROGRESS.setdefault(run_id, [])
        buf.append({"t": time.time(), "phase": phase, "msg": str(msg)})
        buf[:] = buf[-200:]  # 只留最近 200 条


def _poll_progress(run_id):
    with _PROGRESS_LOCK:
        return list(_PROGRESS.get(run_id) or [])


def _clear_progress(run_id):
    with _PROGRESS_LOCK:
        _PROGRESS.pop(run_id, None)


def convert_upload(doc_name, doc_bytes, tpl_name=None, tpl_bytes=None,
                   pages=8, use_llm=False, progress=None):
    """把上传内容落到临时目录，跑完整管线，返回 (pptx 字节, 下载名, 统计)。

    progress: 可选回调 fn(phase, msg)，Web 端用它把阶段回显到前端。
    """
    if len(doc_bytes) > MAX_FILE:
        raise ValueError(f"文件超过 {MAX_FILE // 1024 // 1024}MB 限制")
    src_name = safe_name(doc_name, SUPPORTED)
    import run as run_mod

    _old = getattr(run_mod, "_run_progress", None)
    old_ph = getattr(run_mod, "_run_phase", None)

    def _report(phase, msg):
        if progress is not None:
            progress(phase, msg)
    run_mod._run_progress = _report
    run_mod._run_phase = _report
    try:
        with tempfile.TemporaryDirectory(prefix="doc2ppt_") as td:
            src = Path(td) / src_name
            src.write_bytes(doc_bytes)
            tpl = template_path()
            if tpl_bytes:
                tpl_path = Path(td) / safe_name(tpl_name or "template.pptx", {".pptx"})
                tpl_path.write_bytes(tpl_bytes)
                guard.guard_ooxml(tpl_path)  # 模板同样是不可信输入
                tpl = tpl_path
            out = Path(td) / "out.pptx"
            stats = {}
            _report("开始", "解析文档…")
            _blocks, items, rounds, _log, issues = _run_once(
                src, tpl, out, pages, preview=False, use_llm=use_llm, stats=stats)
            _report("完成", f"生成完成:{len(items)} 页,{len(issues)} 问题")
            blob = out.read_bytes()
            meta = {"pages": len(items), "rounds": rounds, "issues": len(issues),
                    "degraded": bool(stats.get("degraded"))}
            return blob, src.stem + "-输出.pptx", meta
    finally:
        if _old is not None:
            run_mod._run_progress = _old
        else:
            try:
                del run_mod._run_progress
            except AttributeError:
                pass
        if old_ph is not None:
            run_mod._run_phase = old_ph
        else:
            try:
                del run_mod._run_phase
            except AttributeError:
                pass


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


def preview_template(tpl_name, tpl_bytes):
    """模板向导：把上传的 .pptx 落到临时文件，探测语义版式，返回可回显的结构。

    只读不动：不转换、不出片，供前端展示『模板能命中哪些语义版式 + 槽位』。
    """
    from doc2ppt import template as T
    with tempfile.TemporaryDirectory(prefix="doc2ppt_tpl_") as td:
        tpl_path = Path(td) / safe_name(tpl_name or "template.pptx", {".pptx"})
        tpl_path.write_bytes(tpl_bytes)
        guard.guard_ooxml(tpl_path)  # 模板同样是不可信输入
        _prs, infos = T.load_template(tpl_path)
        layouts = []
        for sem in T.LAYOUT_MAP:
            info = infos.get(sem)
            if info is None:
                layouts.append({"sem": sem, "name": None, "desc": "未识别（将回落默认版式）"})
                continue
            slots = "、".join(role for role in ("title", "subtitle", "body", "body2", "caption")
                              if role in info.slots)
            layouts.append({"sem": sem, "name": info.layout.name,
                            "desc": f"槽位：{slots}"})
        return {"layouts": layouts,
                "note": "缺失的语义会按「content 版式」兜底；版式名含 标题/两栏/图表/目录/指标/流程 等关键词优先命中"}



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
    <div class="hint">转换通常 5–20 秒；单文件上限 30MB。生成结果为原生可编辑元素，非图片拼接。
      可选：先传自定义 .pptx 模板并点「预览版式」，确认自动探测出的语义版式后，再填文档转换。</div>
    <div class="tplhc" style="margin-top:10px"></div>
  </form>
</div>
<script>
const f=document.getElementById('f'),st=document.getElementById('st'),btn=document.getElementById('btn');
const tplIn=document.getElementById('template');
const tplWrap=document.createElement('div');
tplWrap.innerHTML='<button id="previewBtn" type="button" style="margin-top:4px">预览模板版式</button>';
tplIn.parentNode.appendChild(tplWrap);
const pv=document.getElementById('previewBtn');
pv.addEventListener('click',async()=>{
  const tf=tplIn.files[0];
  if(!tf){st.className='';st.textContent='请先选择模板文件';return;}
  st.className='';st.textContent='正在分析模板…';
  try{
    const fd=new FormData();fd.append('template',tf);
    const r=await fetch('/template-preview',{method:'POST',body:fd});
    const j=await r.json();
    if(!r.ok){throw new Error(j.error||'分析失败')}
    const rows=(j.layouts||[]).map(x=>`<tr><td>${x.sem}</td><td>${x.name||'-'}</td><td>${x.desc}</td></tr>`).join('');
    const box=document.createElement('div');
    box.style.cssText='margin-top:12px;background:#f6f9ff;border:1px solid #dbe3ee;border-radius:8px;padding:12px;font-size:13px';
    box.innerHTML='<b>模板版式探测结果</b><table style="width:100%;margin-top:6px;border-collapse:collapse"><tr style="color:#6b7280"><th style="text-align:left">语义</th><th style="text-align:left">版式</th><th style="text-align:left">说明</th></tr>'+
      rows+'<tr><td colspan=3 style="color:#6b7280;padding-top:8px">'+ (j.note||'') +'</td></tr></table>';
    document.querySelector('.tplhc').innerHTML='';document.querySelector('.tplhc').appendChild(box);
    st.textContent='模板已识别 '+(j.layouts||[]).length+' 个语义版式；可直接在下方填文档转换';
  }catch(err){st.className='err';st.textContent='错误：'+err.message;}
});
let runStart=null,pool=null;
f.addEventListener('submit',async e=>{
  e.preventDefault();btn.disabled=true;st.className='';st.textContent='正在转换，请稍候…';
  try{
    const rr=await fetch('/runid',{method:'POST'});
    const id=(await rr.json()).id;
    runStart=Date.now();
    pool=setInterval(async()=>{
      try{
        const p=await fetch('/progress/'+id);
        if(p.status===404){clearInterval(pool);return;}
        const j=await p.json();
        const last=j.steps&&j.steps[j.steps.length-1];
        if(last){const secs=Math.round((Date.now()-runStart)/1000);
          st.textContent='['+secs+'s] '+last.phase+'：'+last.msg;}
      }catch(_){}
    },700);
    const fd=new FormData(f);fd.append('runid',id);
    const r=await fetch('/convert',{method:'POST',body:fd});
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
    if(pool)clearInterval(pool);
  }catch(err){st.className='err';st.textContent='错误：'+err.message;if(pool)clearInterval(pool);}
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
        elif path.startswith("/progress/"):
            run_id = path.split("/", 2)[2]
            # Web 端轮询转换阶段：完成后该 id 会立即被清理，返回 404 表示已结束
            buf = _poll_progress(run_id)
            if not buf:
                self._json(404, {"error": "done"})
            else:
                self._json(200, {"steps": buf})
        elif path == "/favicon.ico":
            self._send(204, "image/x-icon", b"")
        else:
            self._json(404, {"error": "not found"})

    def do_POST(self):
        path = self.path.split("?", 1)[0]
        if path == "/runid":
            # 预申请一个运行 id，前端先拿 id 再轮询进度、后发转换
            self._json(200, {"id": _progress_id()})
            return
        if path == "/template-preview":
            # 自定义模板向导：只解析版式结构并回显命中语义，不落盘不转换
            try:
                n = int(self.headers.get("Content-Length") or 0)
                fields, files = parse_multipart(self.headers.get("Content-Type", ""),
                                                self.rfile.read(n) if n > 0 else b"")
                tpl_name, tpl_data = files.get("template", (None, None))
                if not tpl_name or not tpl_data:
                    raise ValueError("缺少模板文件")
                rows = preview_template(tpl_name, tpl_data)
                self._json(200, rows)
            except Exception as e:  # noqa: BLE001 向导分析失败的兜底
                self._json(400, {"error": str(e)})
            return
        if path != "/convert":
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
            # 前端上传前先 GET /runid 拿 id 轮询进度；没拿到时自己生成
            raw_id = (fields.get("runid") or [None])[0] if isinstance(fields.get("runid"), list) else fields.get("runid")
            run_id = str(raw_id) if raw_id else _progress_id()

            def _web_progress(phase, msg):
                _push_progress(run_id, phase, msg)
            with _CONVERT_LOCK:  # COM 渲染不支持并发，串行执行
                blob, out_name, meta = convert_upload(
                    name, data, tpl_name, tpl_data,
                    _pages(fields.get("pages")),
                    fields.get("llm") in ("1", "on", "true"),
                    progress=_web_progress)
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
            _clear_progress(run_id)  # 转换完成，前端轮询该 id 返回 404=结束
        except (ValueError, SystemExit) as e:
            msg = (e.code if isinstance(e.code, str) else "转换失败") \
                if isinstance(e, SystemExit) else str(e)
            _log_run(False, name, {"error": msg}, time.time() - t0,
                     self.client_address[0])
            _push_progress(run_id, "失败", str(msg))
            _clear_progress(run_id)
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
    if not template_path().exists():
        print("启动失败: 缺少默认模板（包内 templates/standard_16x9.pptx）", file=sys.stderr)
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
