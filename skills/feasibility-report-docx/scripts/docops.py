#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docops.py —— editor_sdk 的薄封装，专治「基于模板改长文档」时的坐标与样式坑。

所有命令都接受 --file <docx 绝对路径>，file_id 直接用文件路径。
坐标一律使用 editor_sdk 返回的 DOC 坐标（UTF-16 code unit），不要手算。

常用：
  python docops.py doctor
  python docops.py structure --file X.docx --headings
  python docops.py structure --file X.docx --range 60-90
  python docops.py locate  --file X.docx --text 关键词
  python docops.py delete  --file X.docx --idx 8012,7990,7944
  python docops.py fill    --file X.docx --begin 7700 --end 7886 --file-in c.md
  python docops.py insert  --file X.docx --idx 7699 --md c.md
  python docops.py split   --file X.docx --at "标题完整文本" [--level 2]
  python docops.py style   --file X.docx --pid 7FB57C61 --style HEADING_2
  python docops.py replace-all --file X.docx --old 旧 --new 新 [--dry-run]
  python docops.py verify  --file X.docx --forbid 旧词1,旧词2
  python docops.py save    --file X.docx
  python docops.py sheet   --file X.xlsx [--sheet-id id] [--rows 0-40]
"""

import argparse
import glob
import json
import os
import subprocess
import sys
import tempfile
import time

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

SKILL_DIRS = [
    r"C:\Program Files\WorkBuddy\resources\app.asar.unpacked\resources\plugins\workbuddy-builtin\skills",
    r"C:\Program Files\WorkBuddy\resources\app.asar.unpacked\resources\plugins\workbuddy-builtin\builtin-plugins",
    os.path.expanduser(r"~\.workbuddy\plugins\cache\workbuddy-builtin"),
]
SKILL_NAMES = ["workbuddy-local-office-edit", "tencent-local-office-edit", "workbuddy-docx"]


# ---------------------------------------------------------------- edsdk 定位

def find_edsdk():
    env = os.environ.get("EDSDK_PATH")
    if env and os.path.isfile(env):
        return env
    for root in SKILL_DIRS:
        for name in SKILL_NAMES:
            for pattern in (os.path.join(root, name, "edsdk.py"),
                            os.path.join(root, name, "*", "edsdk.py"),
                            os.path.join(root, "*", name, "*", "edsdk.py"),
                            os.path.join(root, "**", "edsdk.py")):
                hits = sorted(glob.glob(pattern, recursive=True))
                if hits:
                    return hits[0]
    raise SystemExit("找不到 edsdk.py。请设置环境变量 EDSDK_PATH 指向它。")


EDSDK = None


def edsdk_path():
    global EDSDK
    if EDSDK is None:
        EDSDK = find_edsdk()
    return EDSDK


def norm(path):
    """统一成正斜杠绝对路径，规避 Windows 反斜杠在 k=v 解析里的转义问题。"""
    return os.path.abspath(path).replace("\\", "/")


def call(tool, params=None, blob=None, quiet=False):
    """调用 edsdk 工具，返回解析后的 JSON dict。大 payload 自动走 --json-file。"""
    cmd = [sys.executable, edsdk_path(), "call", tool]
    for k, v in (params or {}).items():
        if v is None:
            continue
        cmd.append("%s=%s" % (k, v if isinstance(v, str) else json.dumps(v, ensure_ascii=False)))
    tmp = None
    if blob is not None:
        fd, tmp = tempfile.mkstemp(suffix=".json", text=True)
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(blob, f, ensure_ascii=False)
        cmd += ["--json-file", tmp]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    finally:
        if tmp and os.path.exists(tmp):
            os.remove(tmp)
    out = (proc.stdout or "").strip()
    if ("连接失败" in out or "ECONNREFUSED" in out or "10061" in out) and "editor_sdk" in out:
        print("\n[editor_sdk 服务已停止]\n"
              "  在编辑器里重新打开任意 Office 文件即可拉起服务；"
              "重启宿主程序也有效。\n"
              "  教训：每改完一章就 save，服务挂掉不会丢已保存的内容。\n", file=sys.stderr)
        raise SystemExit(2)
    try:
        data = json.loads(out)
    except Exception:
        if not quiet:
            print(out[:2000] if out else (proc.stderr or "")[:2000])
        return {"ok": False, "error": "返回非 JSON", "raw": out}
    if isinstance(data, dict) and data.get("ok") is False and not quiet:
        print("  ! %s 失败: %s" % (tool, str(data.get("error") or data)[:300]), file=sys.stderr)
    return data


def is_open(fid):
    """轻量探测：能读到结构就说明编辑器已在内存中。"""
    data = call("doc_resolve_document_structure",
                {"file_id": fid, "mode": "outline", "limit": 1}, quiet=True)
    return data.get("ok") is not False and data.get("nodes") is not None


def variants(path):
    """编辑器按 file_id 字符串索引，正/反斜杠是两份不同的键。首次失效时换一种写法。"""
    out = [path]
    alt = path.replace("/", "\\")
    if alt != path:
        out.append(alt)
    return out


def pick_open(path):
    for cand in variants(path):
        if is_open(cand):
            return cand
    return None


def ensure_open(path):
    """已打开则直接复用；只有未打开才 open_file，并返回可用的 file_id。

    三个坑：
    1. 对已打开的文件再次 open_file（默认 fresh open）会用磁盘内容重建编辑器，
       内存里未保存的编辑会全部丢失 —— 所以必须先探测。
    2. open_file 返回纯文本（open completed ... model ready），不是 JSON。
    3. 同一文件用正斜杠与反斜杠注册成两个键，换写法会撞锁并报 document is not open。
    """
    got = pick_open(path)
    if got:
        return got
    for cand in variants(path):
        data = call("open_file", {"file_path": cand, "wait": "true"}, quiet=True)
        raw = str(data.get("raw") or "")
        benign = (data.get("ok") is not False) or ("ready" in raw) or ("open completed" in raw)
        if not benign:
            msg = raw or str(data.get("error") or "")
            if ("10201" in msg or "occupied" in msg.lower()):
                got = pick_open(path)
                if got:
                    return got
                continue
            print("  ! open_file 失败: %s" % msg[:300], file=sys.stderr)
            continue
        for _ in range(30):
            got = pick_open(path)
            if got:
                return got
            time.sleep(0.5)
    print("  ! 无法打开或结构不可读：%s" % path, file=sys.stderr)
    return None


def need_file(args):
    if not args.file:
        raise SystemExit("缺少 --file")
    if not os.path.isfile(args.file):
        raise SystemExit("文件不存在: %s" % args.file)
    fid = ensure_open(norm(args.file))
    if not fid:
        raise SystemExit("无法打开文件：%s" % args.file)
    return fid


# ---------------------------------------------------------------- 结构

def get_nodes(path, mode="compact"):
    data = call("doc_resolve_document_structure",
                {"file_id": path, "mode": mode, "limit": 0})
    if data.get("ok") is False:
        raise SystemExit("读取结构失败: %s" % tr(json.dumps(data, ensure_ascii=False)))
    return data.get("nodes", [])


def tr(s, n):
    s = s or ""
    return s if len(s) <= n else s[:n] + "…"


def cmd_structure(args):
    path = need_file(args)
    mode = "outline" if args.headings else "compact"
    nodes = get_nodes(path, mode)
    if args.json_out:
        with open(args.json_out, "w", encoding="utf-8") as f:
            json.dump(nodes, f, ensure_ascii=False, indent=1)
        print("已写出 %d 个节点 → %s" % (len(nodes), args.json_out))

    lo, hi = (1, 10 ** 9)
    if args.range:
        a, _, b = args.range.partition("-")
        lo = int(a or 1)
        hi = int(b or a)

    rows = 0
    for n in nodes:
        pi = n.get("paragraph_index", 0)
        t = n.get("type", "")
        if args.headings and t != "Heading":
            continue
        if args.tables and t != "Table":
            continue
        if not (lo <= pi <= hi):
            continue
        if t == "Table":
            print("pi=%-4s [TABLE] id=%-10s %sx%s  start=%s end=%s  %s" % (
                pi, n.get("table_id", "-"), n.get("row_count"), n.get("col_count"),
                n.get("start_index"), n.get("end_index"),
                tr(" | ".join(n.get("first_row_texts") or []), 50)))
        elif t in ("Heading", "Title", "Subtitle"):
            print("pi=%-4s [H%s] pid=%-10s %s-%s  %s" % (
                pi, n.get("heading_level", "-"), n.get("paragraph_id", "-"),
                n.get("start_index"), n.get("end_index"),
                tr(n.get("text_preview", ""), args.width)))
        else:
            print("pi=%-4s [P ] pid=%-10s %s-%s  %s" % (
                pi, n.get("paragraph_id", "-"),
                n.get("start_index"), n.get("end_index"),
                tr(n.get("text_preview", ""), args.width)))
        rows += 1
    print("-- 共 %d / %d 个节点 --" % (rows, len(nodes)))


def cmd_locate(args):
    path = need_file(args)
    data = call("doc_find", {"file_id": path, "text": args.text})
    locs = data.get("locations", [])
    print("命中 %s 处「%s」" % (data.get("total", len(locs)), args.text))
    for i, loc in enumerate(locs[: args.limit]):
        extra = ""
        if loc.get("table_id"):
            extra = " [表 %s r%s c%s]" % (loc["table_id"], loc.get("row"), loc.get("col"))
        print("  #%d begin=%-6s end=%-6s pid=%s%s" % (
            i + 1, loc.get("begin"), loc.get("end"), loc.get("paragraph_id", "-"), extra))
        print("      %s" % tr(loc.get("related_text", ""), 110))
    if len(locs) > args.limit:
        print("  …还有 %d 处未显示（--limit 调整）" % (len(locs) - args.limit))


# ---------------------------------------------------------------- 删除

def cmd_delete(args):
    path = need_file(args)
    idxs = [int(x) for x in args.idx.replace(" ", "").split(",") if x.strip()]
    if not idxs:
        raise SystemExit("--idx 为空")
    order = sorted(set(idxs), reverse=True)
    print("删除 %d 段（降序执行）" % len(order))
    for idx in order:
        data = call("doc_delete_paragraph", {"file_id": path, "idx": idx}, quiet=True)
        ok = data.get("ok") is not False
        print("  idx=%-6s %s" % (idx, "ok" if ok else "FAIL " + tr(str(data.get('error')), 80)))
    print("完成。请重新 structure 确认坐标。")


# ---------------------------------------------------------------- 写入

def read_body(args):
    if getattr(args, "file_in", None):
        with open(args.file_in, encoding="utf-8") as f:
            return f.read()
    if getattr(args, "text", None) is not None:
        return args.text
    raise SystemExit("需要 --file-in <md路径> 或 --text")


def cmd_fill(args):
    path = need_file(args)
    text = read_body(args)
    data = call("doc_replace_text", {"file_id": path},
                blob={"ranges": [{"begin": args.begin, "end": args.end}], "text": text})
    print("替换 [%s,%s) 长度 %d → %s" % (
        args.begin, args.end, args.end - args.begin,
        "ok" if data.get("ok") is not False else "FAIL " + tr(str(data.get("error")), 120)))


def is_blank(node):
    return (not (node.get("text_preview") or "").strip()
            and (node.get("end_index", 0) - node.get("start_index", 0)) <= 2)


def clean_empty_at(path, start, limit=1):
    """删除起点恰好等于 start 的空段落。doc_insert_markdown 每次插入都会留下一个。

    返回删掉的段落数。
    """
    removed = 0
    for _ in range(limit):
        nodes = get_nodes(path, "outline")
        hit = None
        for n in nodes:
            if n.get("start_index") == start and is_blank(n):
                hit = n
                break
        if not hit:
            break
        res = call("doc_delete_paragraph", {"file_id": path, "idx": start}, quiet=True)
        if res.get("ok") is False:
            break
        removed += 1
    return removed


def resolve_after_heading(path, text, begin=None, hit=1):
    """返回「该标题下一段的 start_index」。

    关键结论（实测）：doc_insert_markdown 的 idx 语义是"插入点所在段落"。
    - 传标题的 end_index → 内容并进标题段，标题被污染；
    - 传标题的 start_index → 正文被塞到标题前面，标题被推到后面；
    - 传标题**下一段的 start_index** → 正确，标题与正文各自成段。

    同一标题文字在正文、目录里可能多次出现，所以优先选「段落类型是 Heading」的那一处。
    """
    data = call("doc_find", {"file_id": path, "text": text})
    locs = data.get("locations", [])
    if not locs:
        raise SystemExit("未找到标题文本：%s" % text)

    nodes = get_nodes(path, "outline")

    def node_of(b):
        for i, n in enumerate(nodes):
            if n.get("start_index") == b:
                return i
        for i, n in enumerate(nodes):
            if n.get("start_index", 0) <= b < n.get("end_index", 0):
                return i
        return None

    target = None
    if begin is not None:
        for cand in locs:
            if cand.get("begin") == begin:
                target = cand
                break
        if target is None:
            raise SystemExit("未在 begin=%s 处找到「%s」（%d 处命中）" % (begin, text, len(locs)))
    else:
        for cand in locs:                       # 优先命中标题段
            i = node_of(cand.get("begin"))
            if i is not None and nodes[i].get("type") == "Heading":
                target = cand
                break
        if target is None:
            target = locs[min(hit, len(locs)) - 1]

    b, e = target.get("begin"), target.get("end")
    i = node_of(b)
    if i is None:
        raise SystemExit("定位不到标题所在段落")

    head = nodes[i]
    if head.get("type") != "Heading":
        print("  ~ 注意：命中段落不是标题（type=%s），可能选错了位置，建议用 --begin 指定"
              % head.get("type"), file=sys.stderr)
    nxt = nodes[i + 1] if i + 1 < len(nodes) else None
    if nxt is None:
        print("  ! 该标题已是最后一个节点，回退用标题 end_index", file=sys.stderr)
        return e, head
    return nxt.get("start_index"), head


def cmd_insert(args):
    path = need_file(args)
    md = os.path.abspath(args.md).replace("\\", "/")

    if args.after_heading:
        idx, head = resolve_after_heading(path, args.after_heading, args.begin, args.hit)
        print("插入点 → %s（标题「%s」的下一段起点）" % (idx, args.after_heading))
    elif args.idx is not None:
        idx = args.idx
    else:
        raise SystemExit("需要 --idx 或 --after-heading")

    data = call("doc_insert_markdown",
                {"file_id": path, "idx": idx, "markdown": "file://" + md})
    if data.get("ok") is False:
        raise SystemExit("插入失败: %s" % tr(str(data.get("error")), 200))
    print("已插入。position=%s last_edit_index=%s" % (
        data.get("position"), data.get("last_edit_index")))
    print("下一步：structure --range 复查——标题应独占一段，正文紧随其后。")


def clean_blanks_in_window(path, lo, hi):
    """删除 [lo, hi] 区间内的空标题/空段落（插入残留），降序执行。"""
    nodes = get_nodes(path, "outline")
    hits = [n.get("start_index") for n in nodes
            if n.get("type") == "Heading" and is_blank(n)
            and lo <= (n.get("start_index") or 0) <= hi]
    removed = 0
    for s in sorted(set(hits), reverse=True):
        if call("doc_delete_paragraph", {"file_id": path, "idx": s}, quiet=True).get("ok") is False:
            break
        removed += 1
    return removed


def cmd_fix_heading(args):
    """一键修复 doc_insert_markdown 造成的标题污染。

    链条：删掉紧邻的空段 → 在标题文本末尾拆段 → 首段文本改回标题 → 恢复 HEADING_n。
    """
    path = need_file(args)
    data = call("doc_find", {"file_id": path, "text": args.at})
    locs = data.get("locations", [])
    if not locs:
        raise SystemExit("未找到标题文本：%s" % args.at)
    loc = None
    if args.begin is not None:
        for cand in locs:
            if cand.get("begin") == args.begin:
                loc = cand
                break
        if loc is None:
            raise SystemExit("未在 begin=%s 处找到「%s」（%d 处命中，改 --begin 或 --hit）"
                             % (args.begin, args.at, len(locs)))
    else:
        loc = locs[args.hit - 1] if len(locs) >= args.hit else locs[0]

    b, e = loc.get("begin"), loc.get("end")
    nodes = get_nodes(path, "outline")
    cur = None
    for n in nodes:
        if n.get("start_index") == b:
            cur = n
            break
    if cur is None:
        for n in nodes:
            if n.get("start_index", 0) <= b < n.get("end_index", 0):
                cur = n
                break

    span = (cur.get("end_index", 0) - cur.get("start_index", 0)) if cur else 0
    polluted = span > (e - b) + 2

    if args.dry_run:
        print("[dry-run] 「%s」begin=%s end=%s  段落跨度=%s  标题层级=%s" % (
            args.at, b, e, span, cur.get("heading_level") if cur else "-"))
        print("           判定：%s" % ("被污染，需要修复" if polluted else "正常，无需修复"))
        if polluted and cur:
            print("           修复将删掉起点 %s 处的空段、在 %s 拆段、并把 %s-%s 改回标题"
                  % (cur.get("end_index"), e, b, e))
        return

    removed = clean_empty_at(path, cur.get("end_index") if cur else e)
    if removed:
        print("清理空段 %d 个" % removed)

    if polluted:
        res = call("doc_insert_paragraph_with_text",
                   {"file_id": path, "idx": e, "text": "", "type": 0})
        print("拆段 @%s → %s" % (e, "ok" if res.get("ok") is not False else "FAIL"))

    res = call("doc_replace_text", {"file_id": path},
               blob={"ranges": [{"begin": b, "end": e}], "text": args.at})
    print("标题文本复原 → %s" % ("ok" if res.get("ok") is not False else "FAIL"))

    style = args.style or ("HEADING_%d" % args.level)
    res = call("doc_apply_named_style", {"file_id": path},
               blob={"ranges": [{"begin": b, "end": e}], "named_style_type": style})
    print("套用 %s → %s" % (style, "ok" if res.get("ok") is not False else "FAIL"))

    swept = clean_blanks_in_window(path, e, e + 300)
    if swept:
        print("清理区间内空标题 %d 个" % swept)
    print("完成。请 structure 复查该处：标题应独占一段，正文紧随其后。")


def cmd_split(args):
    path = need_file(args)
    data = call("doc_find", {"file_id": path, "text": args.at})
    locs = data.get("locations", [])
    if not locs:
        raise SystemExit("未找到文本：%s" % args.at)
    loc = None
    if args.begin is not None:
        for cand in locs:
            if cand.get("begin") == args.begin:
                loc = cand
                break
        if loc is None:
            raise SystemExit("未在 begin=%s 处找到「%s」（共 %d 处命中，去掉 --begin 或换 --hit）"
                             % (args.begin, args.at, len(locs)))
    else:
        loc = locs[args.hit - 1] if len(locs) >= args.hit else locs[0]
    end = loc.get("end")
    res = call("doc_insert_paragraph_with_text",
               {"file_id": path, "idx": end, "text": "", "type": 0})
    print("在 %s 处拆分 → %s" % (end, "ok" if res.get("ok") is not False else "FAIL"))
    print("下一步：fill --begin %s --end %s --text \"%s\"，再 style 恢复 HEADING_n"
          % (loc.get("begin"), end, args.at))


def cmd_style(args):
    path = need_file(args)
    style = str(args.style).upper().replace("H", "HEADING_") if str(args.style).isdigit() else str(args.style).upper()
    style = style.replace("HEADING__", "HEADING_")
    if style.startswith("HEADING_") and not style[8:].isdigit():
        style = "HEADING_" + str(args.level)

    if args.pid:
        blob = {"paragraph_ids": [args.pid], "named_style_type": style}
    elif args.text:
        data = call("doc_find", {"file_id": path, "text": args.text})
        locs = data.get("locations", [])
        if not locs:
            raise SystemExit("未找到文本：%s" % args.text)
        pid = locs[0].get("paragraph_id")
        if pid:
            blob = {"paragraph_ids": [pid], "named_style_type": style}
        else:
            blob = {"ranges": [{"begin": locs[0]["begin"], "end": locs[0]["end"]}],
                    "named_style_type": style}
    else:
        raise SystemExit("需要 --pid 或 --text")

    res = call("doc_apply_named_style", {"file_id": path}, blob=blob)
    print("套用 %s → %s" % (style, "ok" if res.get("ok") is not False else "FAIL " + tr(str(res.get("error")), 120)))


def cmd_replace_all(args):
    path = need_file(args)
    if args.dry_run:
        data = call("doc_find", {"file_id": path, "text": args.old})
        locs = data.get("locations", [])
        print("[dry-run] 「%s」命中 %s 处" % (args.old, data.get("total", len(locs))))
        for loc in locs[:20]:
            print("  begin=%-6s %s" % (loc.get("begin"), tr(loc.get("related_text", ""), 90)))
        return
    data = call("doc_find_and_replace",
                {"file_id": path, "old_text": args.old, "new_text": args.new})
    print("替换「%s」→「%s」: %s" % (
        args.old, args.new, "ok" if data.get("ok") is not False else tr(str(data.get("error")), 120)))
    chk = call("doc_find", {"file_id": path, "text": args.old})
    print("复查残留: %s 处" % chk.get("total", 0))


# ---------------------------------------------------------------- 校验

def cmd_verify(args):
    path = need_file(args)
    nodes = get_nodes(path, "outline")
    heads = [n for n in nodes if n.get("type") == "Heading"]

    tdata = call("doc_list_tables", {"file_id": path}, quiet=True)
    tables = (tdata.get("tables") or tdata.get("data") or []) if tdata.get("ok") is not False else []
    if not isinstance(tables, list):
        tables = []

    problems = 0
    print("[总览] 标题 %d 个，表格 %d 张，节点 %d 个" % (len(heads), len(tables), len(nodes)))

    # 空标题：preview 为空且坐标跨度极小（只剩段落符）
    empty = [n for n in heads
             if not (n.get("text_preview") or "").strip()
             and (n.get("end_index", 0) - n.get("start_index", 0)) <= 2]
    if empty:
        problems += len(empty)
        print("  ! 空标题 %d 个：" % len(empty))
        for n in empty[:15]:
            print("      pi=%s H%s pid=%s start=%s" % (
                n.get("paragraph_index"), n.get("heading_level"),
                n.get("paragraph_id"), n.get("start_index")))
    else:
        print("  空标题 0 个")

    # 标题污染：靠坐标跨度判断，不依赖被截断的 preview
    polluted = []
    for n in heads:
        span = n.get("end_index", 0) - n.get("start_index", 0)
        if span > args.max_head_span:
            polluted.append(n)
    if polluted:
        problems += len(polluted)
        print("  ! 疑似被正文污染的标题 %d 个（跨度 > %d）：" % (len(polluted), args.max_head_span))
        for n in polluted[:15]:
            span = n.get("end_index", 0) - n.get("start_index", 0)
            print("      pi=%s H%s span=%s start=%s  %s" % (
                n.get("paragraph_index"), n.get("heading_level"), span,
                n.get("start_index"), tr(n.get("text_preview"), 40)))
        print("      处置：split --at <标题文本> → fill 首段 → style HEADING_n")
    else:
        print("  标题跨度均正常（≤ %d）" % args.max_head_span)

    # 重名标题：preview 有截断，只作线索
    seen = {}
    for n in heads:
        key = (n.get("text_preview") or "").strip()
        if key:
            seen.setdefault(key, []).append(n.get("paragraph_index"))
    dups = {k: v for k, v in seen.items() if len(v) > 1}
    if dups:
        print("  ~ 前缀相同的标题 %d 组（preview 有截断，需人工确认）：" % len(dups))
        for k, v in list(dups.items())[:10]:
            print("      「%s」@ pi=%s" % (tr(k, 24), v))

    prev, jumps = None, []
    for n in heads:
        lv = n.get("heading_level") or 0
        if prev is not None and lv - prev > 1:
            jumps.append((n.get("paragraph_index"), "H%s→H%s" % (prev, lv)))
        prev = lv
    if jumps:
        print("  ~ 层级跳级 %d 处（确认是否漏了标题）：%s" % (len(jumps), jumps[:8]))

    if args.forbid:
        print("[残留词]")
        for word in [w.strip() for w in args.forbid.split(",") if w.strip()]:
            data = call("doc_find", {"file_id": path, "text": word}, quiet=True)
            total = data.get("total", 0)
            if total:
                problems += 1
            print("  %-14s %s" % (word, "ok" if total == 0 else "残留 %s 处" % total))
            for loc in (data.get("locations") or [])[:3]:
                print("       begin=%-6s %s" % (loc.get("begin"), tr(loc.get("related_text"), 90)))

    print("[占位符]")
    for word in ["TODO", "待补充", "待定", "XX", "{", "【"]:
        data = call("doc_find", {"file_id": path, "text": word}, quiet=True)
        total = data.get("total", 0)
        if total:
            problems += 1
            print("  ! %-6s %s 处" % (word, total))
        else:
            print("  %-6s 0" % word)

    print("\n合计待处理项：%d" % problems)
    return 0 if problems == 0 else 1


# ---------------------------------------------------------------- 表格数据源

def cmd_sheet(args):
    path = need_file(args)
    info = call("sheet_get_sheet_info", {"file_id": path})
    sheets = info.get("sheets") or info.get("sheet_list") or []
    if not sheets and isinstance(info.get("data"), dict):
        sheets = info["data"].get("sheets") or []
    if not sheets:
        print(json.dumps(info, ensure_ascii=False)[:1500])
        return
    print("子表：")
    for s in sheets:
        print("  id=%s  %s" % (s.get("id") or s.get("sheet_id"), s.get("name")))

    sid = args.sheet_id or (sheets[0].get("id") or sheets[0].get("sheet_id"))
    used = call("sheet_get_used_range", {"file_id": path, "sheet_id": sid}, quiet=True)
    ur = used.get("used_range") or used
    print("\n已用区域：%s" % json.dumps(ur, ensure_ascii=False)[:200])

    a, _, b = args.rows.partition("-")
    r0, r1 = int(a or 0), int(b or (a or 60))
    data = call("sheet_get_cell_data", {
        "file_id": path, "sheet_id": sid,
        "start_row": r0, "end_row": r1,
        "start_col": 0, "end_col": args.cols,
        "return_csv": "true",
    })
    print("\n--- CSV (rows %d-%d) ---" % (r0, r1))
    print(data.get("csv") or data.get("data") or json.dumps(data, ensure_ascii=False)[:4000])


# ---------------------------------------------------------------- 其它

def cmd_save(args):
    path = need_file(args)
    data = call("save_file", {"file_id": path, "file_path": norm(args.to or args.file)})
    raw = str(data.get("raw") or "")
    ok = (data.get("ok") is not False) or ("saved" in raw.lower())
    print("保存 → %s" % ("ok" if ok else tr(raw or str(data.get("error")), 200)))


def cmd_doctor(args):
    p = edsdk_path()
    print("edsdk: %s" % p)
    data = call("doc_resolve_document_structure", {"file_id": "__probe__", "limit": 1}, quiet=True)
    print("连通性: %s" % ("ok" if data.get("ok") is not False or "未找到" not in str(data) else "异常"))
    print("提示：连通性仅供参考，真正可用性以 structure 命令读到你自己的文件为准。")


def build_parser():
    ap = argparse.ArgumentParser(prog="docops.py", description="editor_sdk 文档操作助手")
    sub = ap.add_subparsers(dest="cmd", required=True)

    def add(name, fn, help_):
        p = sub.add_parser(name, help=help_)
        p.add_argument("--file", help="docx/xlsx 绝对路径")
        p.set_defaults(func=fn)
        return p

    p = add("structure", cmd_structure, "打印文档结构（坐标/标题/表格）")
    p.add_argument("--headings", action="store_true", help="只看标题")
    p.add_argument("--tables", action="store_true", help="只看表格")
    p.add_argument("--range", help="按 paragraph_index 过滤，如 60-90")
    p.add_argument("--width", type=int, default=40, help="preview 截断长度")
    p.add_argument("--json-out", help="把全量节点写入该 JSON 文件")

    p = add("locate", cmd_locate, "查找文本，打印坐标")
    p.add_argument("--text", required=True)
    p.add_argument("--limit", type=int, default=20)

    p = add("delete", cmd_delete, "按 start_index 批量删段（内部降序）")
    p.add_argument("--idx", required=True, help="逗号分隔的 start_index")

    p = add("fill", cmd_fill, "用文件/文本替换坐标区间")
    p.add_argument("--begin", type=int, required=True)
    p.add_argument("--end", type=int, required=True)
    p.add_argument("--file-in", dest="file_in", help="内容来源 md/txt")
    p.add_argument("--text", help="直接给文本")

    p = add("insert", cmd_insert, "在坐标处插入 Markdown")
    p.add_argument("--idx", type=int, help="显式插入坐标（推荐改用 --after-heading）")
    p.add_argument("--after-heading", dest="after_heading", help="在该标题的下一段插入")
    p.add_argument("--begin", type=int, help="配合 --after-heading 选定匹配项")
    p.add_argument("--hit", type=int, default=1, help="配合 --after-heading 选第几个匹配")
    p.add_argument("--md", required=True)

    p = add("split", cmd_split, "在指定文本末尾拆段（治标题污染）")
    p.add_argument("--at", required=True, help="标题完整文本")
    p.add_argument("--hit", type=int, default=1, help="用第几个匹配，默认 1")
    p.add_argument("--begin", type=int, help="按起始坐标选定匹配项（更可靠）")
    p.add_argument("--level", type=int, default=None, help="仅用于 style 默认值")

    p = add("fix-heading", cmd_fix_heading, "一键修复标题污染（删空段+拆段+复原+样式）")
    p.add_argument("--at", required=True, help="标题完整文本")
    p.add_argument("--begin", type=int, help="按起始坐标选定匹配项（推荐）")
    p.add_argument("--hit", type=int, default=1)
    p.add_argument("--level", type=int, default=2)
    p.add_argument("--style", help="直接指定样式名，如 HEADING_2")
    p.add_argument("--dry-run", dest="dry_run", action="store_true")

    p = add("style", cmd_style, "按 pid/文本套命名样式")
    p.add_argument("--pid")
    p.add_argument("--text")
    p.add_argument("--style", default="HEADING_2", help="HEADING_1..9 / TITLE / NORMAL_TEXT / 2")
    p.add_argument("--level", type=int, default=2)

    p = add("replace-all", cmd_replace_all, "全局文字替换（含 dry-run）")
    p.add_argument("--old", required=True)
    p.add_argument("--new", default="")
    p.add_argument("--dry-run", dest="dry_run", action="store_true")

    p = add("verify", cmd_verify, "定稿自查")
    p.add_argument("--forbid", default="", help="逗号分隔的旧业务残留词")
    p.add_argument("--max-head-span", dest="max_head_span", type=int, default=60,
                   help="标题段落坐标跨度上限，超过视为被正文污染")

    add("save", cmd_save, "保存")
    sub.choices["save"].add_argument("--to", help="另存为路径")

    p = add("sheet", cmd_sheet, "读 xlsx 报价单/清单")
    p.add_argument("--sheet-id")
    p.add_argument("--rows", default="0-60")
    p.add_argument("--cols", type=int, default=12)

    add("doctor", cmd_doctor, "环境自检")
    return ap


def main(argv=None):
    args = build_parser().parse_args(argv)
    if args.cmd != "doctor" and getattr(args, "file", None) is None:
        raise SystemExit("缺少 --file")
    rc = args.func(args)
    sys.exit(rc if isinstance(rc, int) else 0)


if __name__ == "__main__":
    main()
