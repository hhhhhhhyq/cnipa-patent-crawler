#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CNIPA 专利检索爬虫（复用已登录 Chrome + CDP 网络抓包，UI 驱动）

目标: https://pss-system.cponline.cnipa.gov.cn/conventionalSearch

经实测确认的关键机制：
  - 登录令牌在 localStorage（同源跨标签共享）；必须操作 CNIPA 标签且置于前台(否则页面不布局)。
  - 检索接口 POST .../search/results/getResults 带 WAF 动态令牌，不能直接重放，只能 UI 触发。
  - “自动识别”检索框需真实键盘输入；回车提交后整页导航到 retrieveList（导航瞬间请求抓不到）。
  - 结果页内操作(设排序/每页条数/翻页/日期筛选+确定)会原地重发 getResults → 可抓到。
  - 日期范围是明文参数 apdInterval；在结果页填“申请日”起止日期点“确定”即服务端精确过滤。
  - 故流程：搜索→设申请日起止日期(确定)→设40条/页(抓第1页)→翻页抓完所有页。

用法:
  python3 cnipa_crawler.py run --input 公司名.xlsx --output 专利汇总.xlsx \
          --start 2020-01-01 --end 2023-12-31
"""

import argparse
import base64
import json
import os
import random
import re
import subprocess
import sys
import time
import urllib.request

import pandas as pd

# ── 配置 ──────────────────────────────────────────────────────────────────────
CDP_PORT = "9222"
SEARCH_URL = "https://pss-system.cponline.cnipa.gov.cn/conventionalSearch"
PAGE_LIMIT_LABEL = "40 条/页"
MAX_PAGES = 600
INV_TYPE = {"FM": "发明", "SX": "实用新型", "SY": "实用新型",
            "WG": "外观设计", "WS": "外观设计", "XX": "实用新型"}

# 等待时间均为【随机区间】(秒)，每次取区间内随机值，避免固定节奏被识别为机器
WAIT_SEARCH = (6.0, 10.0)        # 检索/导航后
WAIT_ACTION = (3.5, 6.5)         # 翻页/下拉/确定等页内操作后
DELAY_COMPANIES = (3.0, 8.0)     # 公司与公司之间
WAIT_CITE = (1.0, 2.4)           # --citations 时每条专利之间
WAIT_TINY = (0.4, 1.0)           # 输入/聚焦等小停顿


def nap(span):
    """在区间 (lo, hi) 内随机睡眠。"""
    lo, hi = span
    time.sleep(random.uniform(lo, hi))


# ── agent-browser 封装 ────────────────────────────────────────────────────────
def _ab(*args, timeout=60):
    try:
        r = subprocess.run(["agent-browser", "--cdp", CDP_PORT, *args],
                           capture_output=True, text=True, timeout=timeout)
    except FileNotFoundError:
        sys.exit("✗ 未找到 agent-browser，请先: npm install -g agent-browser")
    except subprocess.TimeoutExpired:
        return ""
    return r.stdout


def ab_eval(js, timeout=60):
    out = _ab("eval", "-b", base64.b64encode(js.encode()).decode(), timeout=timeout).strip()
    if len(out) >= 2 and out[0] == '"' and out[-1] == '"':
        try:
            return json.loads(out)
        except Exception:
            return out[1:-1]
    return out


def ab_type(text):
    _ab("keyboard", "type", text, timeout=30)


def ab_press(key):
    _ab("press", key, timeout=20)


def net_clear():
    _ab("network", "requests", "--clear", timeout=30)


def net_last_getresults():
    """返回最近一次 getResults 的 (postData_dict, parsed_body) 或 (None, None)。"""
    out = _ab("network", "requests", "--filter", "getResults", "--json", timeout=60)
    try:
        reqs = json.loads(out).get("data", {}).get("requests", [])
    except Exception:
        return None, None
    if not reqs:
        return None, None
    det = _ab("network", "request", str(reqs[-1].get("requestId")), "--json", timeout=60)
    try:
        data = json.loads(det).get("data", {})
        pd_ = json.loads(data.get("postData") or "{}")
        body = json.loads(data.get("responseBody") or "{}")
        return pd_, body
    except Exception:
        return None, None


def net_all_getresults_records():
    """取当前已捕获的所有 getResults 响应里的记录列表（合并）。"""
    out = _ab("network", "requests", "--filter", "getResults", "--json", timeout=60)
    try:
        reqs = json.loads(out).get("data", {}).get("requests", [])
    except Exception:
        return []
    recs = []
    for r in reqs:
        det = _ab("network", "request", str(r.get("requestId")), "--json", timeout=60)
        try:
            body = json.loads(json.loads(det).get("data", {}).get("responseBody") or "{}")
            rl = (body.get("t") or {}).get("searchResultRecord") or []
            recs.extend(rl)
        except Exception:
            continue
    return recs


# ── CDP / 标签页 / 登录 ───────────────────────────────────────────────────────
def ensure_cdp():
    detect = os.path.join(os.path.expanduser("~"),
                          ".claude/skills/browser-cdp/scripts/setup-cdp-chrome.js")
    if os.path.exists(detect):
        out = subprocess.run(["node", detect, CDP_PORT, "--detect-only"],
                             capture_output=True, text=True).stdout
        if "CDP_STATUS=ready" not in out:
            print("✗ CDP 未就绪。先启动 debug Chrome 并登录 CNIPA：")
            print(f"    node {detect} {CDP_PORT} --yes")
            sys.exit(1)
    print("✓ CDP 已就绪")


def cdp_activate_cnipa():
    """把 CNIPA 标签置于前台（否则后台标签不布局，元素不可见/键盘失效）。"""
    try:
        tabs = json.loads(urllib.request.urlopen(
            f"http://127.0.0.1:{CDP_PORT}/json", timeout=10).read().decode())
    except Exception:
        return
    for t in tabs:
        if t.get("type") == "page" and "cnipa.gov.cn" in t.get("url", ""):
            try:
                urllib.request.urlopen(
                    f"http://127.0.0.1:{CDP_PORT}/json/activate/{t['id']}", timeout=10)
            except Exception:
                pass
            return


def switch_to_cnipa_tab():
    out = _ab("tab", "list", timeout=30)
    for line in out.splitlines():
        m = re.search(r"\[(t\d+)\].*?(https?://\S+)", line)
        if m and "cnipa.gov.cn" in m.group(2):
            _ab("tab", m.group(1), timeout=30)
            return True
    _ab("tab", "new", timeout=30)
    return False


def goto_search_page():
    _ab("open", SEARCH_URL, timeout=90)
    cdp_activate_cnipa()
    _ab("wait", "4000", timeout=40)


def check_logged_in():
    return ab_eval("localStorage.getItem('token')?'yes':'no'") == "yes"


# ── 风控 / 异常检测 + 退避重试 ───────────────────────────────────────────────────
BASE_BACKOFF = 30.0      # 首次退避秒数
MAX_BACKOFF = 600.0      # 单次退避上限（10 分钟）
RECOVER_TRIES = 6        # 最多退避重试次数


def page_health():
    """返回 'ok' | 'logout' | 'block'。
    - logout：token 丢失或跳转登录页（需你在 Chrome 重新登录）。
    - block ：页面出现限流/风控提示弹窗（频繁/繁忙/验证/稍后/拒绝/限制 等）。
    只看错误弹窗/对话框，避免把正文里的'验证'等词误判。"""
    js = ("(function(){var href=location.href.toLowerCase();"
          "if(!localStorage.getItem('token')||href.indexOf('login')>=0)return 'logout';"
          "var ws=['频繁','稍后','繁忙','人机','滑动验证','拒绝访问','访问受限','请求过','too many','forbidden','rate limit'];"
          "var ns=document.querySelectorAll('.el-message,.el-message-box__message,.el-notification__content,.el-dialog__body');"
          "for(var i=0;i<ns.length;i++){if(ns[i].offsetParent){var t=ns[i].textContent||'';"
          "for(var j=0;j<ws.length;j++){if(t.indexOf(ws[j])>=0)return 'block';}}}"
          "return 'ok';})()")
    h = ab_eval(js)
    return h if h in ("ok", "logout", "block") else "ok"


def wait_recover(reason, navigate=False):
    """检测到异常后指数退避并重新检查健康；恢复返回 True，超限返回 False。"""
    for i in range(RECOVER_TRIES):
        wait = min(MAX_BACKOFF, BASE_BACKOFF * (2 ** i))
        h = page_health()
        if h == "logout":
            print(f"   ⏸ {reason}：登录态失效，请在 Chrome 里重新登录 CNIPA。"
                  f"等待 {int(wait)}s 后重试（{i + 1}/{RECOVER_TRIES}）…")
        else:
            print(f"   ⏸ {reason}：疑似限流/风控，退避 {int(wait)}s 后重试（{i + 1}/{RECOVER_TRIES}）…")
        time.sleep(wait)
        if navigate:
            goto_search_page()
        if page_health() == "ok":
            print("   ▶ 已恢复，继续")
            return True
    return False


# ── 字段解析（基于实测真实字段）─────────────────────────────────────────────────
def extract_item(record, en_name):
    items = record.get("items") or {}
    if isinstance(items, dict):
        for group in items.values():
            if isinstance(group, list):
                for it in group:
                    if isinstance(it, dict) and it.get("indexEnName") == en_name:
                        return it.get("value", "")
    return ""


def extract_abstract(record):
    ab = record.get("abview")
    if not ab or not isinstance(ab, list) or not isinstance(ab[0], dict):
        return ""
    val = ab[0].get("value", "")
    m = re.search(r"<base:Paragraphs[^>]*>(.*?)</base:Paragraphs>", val, re.S)
    text = m.group(1) if m else val
    text = re.sub(r"<!\[CDATA\[.*?\]\]>", "", text, flags=re.S)
    text = re.sub(r"<[^>]+>", "", text)
    return re.sub(r"\s+", " ", text).strip()


def join_ipc(record):
    raw = []
    if record.get("ipcMain"):
        raw.append(record["ipcMain"])
    for it in (record.get("ipcDetail") or []):
        if isinstance(it, dict) and it.get("value"):
            raw.append(it["value"])
    seen, out = set(), []
    for chunk in raw:
        for v in re.split(r"[;,\s]+", str(chunk)):
            v = v.strip()
            if v and v not in seen:
                seen.add(v); out.append(v)
    return "; ".join(out)


def normalize(record, company):
    return {
        "检索公司": company,
        "申请号": record.get("apo") or record.get("ap") or "",
        "专利名称": record.get("ti") or "",
        "申请日": record.get("apd") or "",
        "公开(公告)号": record.get("pn") or "",
        "公开(公告)日": record.get("pd") or "",
        "公开日期(公开)": extract_item(record, "GK_PD"),
        "公开日期(授权)": extract_item(record, "SQ_PD"),
        "专利类型": INV_TYPE.get(record.get("invType"), record.get("invType") or ""),
        "申请人": record.get("pa") or "",
        "发明人": record.get("inv") or "",
        "代理机构": extract_item(record, "AGY"),
        "申请人所在国/地区": extract_item(record, "AC"),
        "主IPC分类号": record.get("ipcMain") or "",
        "IPC分类号": join_ipc(record),
        "被引证次数": record.get("_citedCount", ""),
        "摘要": extract_abstract(record),
    }


# ── 时间过滤（客户端兜底）────────────────────────────────────────────────────────
def to_yyyymmdd(s):
    digits = re.sub(r"\D", "", str(s))[:8]
    if len(digits) < 4:
        return None
    digits = (digits + "0101")[:8]
    try:
        return int(digits)
    except ValueError:
        return None


def in_range(record, start, end, basis):
    key = "apd" if basis == "申请日" else "pd"
    # apd 可能形如 "2023.12.05;2024.01.09"，取任一落在区间即保留
    raw = str(record.get(key) or "")
    cands = [to_yyyymmdd(x) for x in re.split(r"[;,\s]+", raw) if to_yyyymmdd(x)]
    if not cands:
        return True
    for d in cands:
        if (not start or d >= start) and (not end or d <= end):
            return True
    return False


def fmt_date(s):
    """'2020-01-01' / '20200101' -> ('2020-01-01', 20200101)；None -> (None,None)。"""
    d = to_yyyymmdd(s) if s else None
    if d is None:
        return None, None
    ds = str(d)
    return f"{ds[:4]}-{ds[4:6]}-{ds[6:8]}", d


# ── 结果页操作（均经实测）─────────────────────────────────────────────────────────
def select_dropdown_option(open_value_pat, option_text, scope=".el-select"):
    ab_eval("(function(){var s=document.querySelectorAll(%s+' .el-input__inner');"
            "for(var i=0;i<s.length;i++){if(new RegExp(%r).test(s[i].value)){"
            "s[i].dispatchEvent(new MouseEvent('mousedown',{bubbles:true}));s[i].click();return 'ok';}}"
            "return 'none';})()" % (json.dumps(scope), open_value_pat))
    nap((0.9, 1.7))
    return ab_eval("(function(){var its=document.querySelectorAll('.el-select-dropdown__item');"
                   "for(var i=0;i<its.length;i++){if(its[i].textContent.trim()===%s){its[i].click();return 'picked';}}"
                   "return 'noopt';})()" % json.dumps(option_text))


def set_apd_filter(start_disp, end_disp):
    """在结果页填【申请日】起止日期并点确定（服务端过滤）。返回是否成功设置。"""
    def fill(idx, value):
        ok = ab_eval("(function(){var s=document.querySelectorAll('input[placeholder=\"开始日期\"]'),"
                     "e=document.querySelectorAll('input[placeholder=\"结束日期\"]');"
                     "var el=(%d===0)?s[0]:e[0];if(!el)return 'no';el.focus();return 'ok';})()" % idx)
        if ok != "ok":
            return False
        ab_press("Control+a"); ab_press("Delete")
        ab_type(value)
        nap(WAIT_TINY)
        ab_press("Enter")
        nap(WAIT_TINY)
        ab_eval("document.body.click()")
        nap((0.3, 0.7))
        return True

    if start_disp:
        fill(0, start_disp)
    if end_disp:
        fill(1, end_disp)
    # 点确定
    net_clear()
    ab_eval("(function(){var bs=document.querySelectorAll('button.el-button--primary');"
            "for(var i=0;i<bs.length;i++){if((bs[i].textContent||'').trim()==='确定'&&bs[i].offsetParent){bs[i].click();return 'ok';}}return 'no';})()")
    nap(WAIT_ACTION)
    return True


def click_next_page():
    return ab_eval("(function(){var b=document.querySelector('.el-pagination .btn-next');"
                   "if(b&&!b.disabled){b.click();return 'ok';}return 'end';})()") == "ok"


def fetch_citations_current_page():
    """点当前页每条专利的【被引证】按钮，抓 patcitedinfos 的 totalCount。

    返回 {anId: 被引证次数}。逐条点击较慢，仅在 --citations 时启用。
    """
    net_clear()
    n = ab_eval("(function(){var a=document.querySelectorAll('li,span,div'),c=0;"
                "for(var i=0;i<a.length;i++){if((a[i].textContent||'').trim()==='被引证'&&a[i].offsetParent)c++;}return c;})()")
    try:
        n = int(n)
    except ValueError:
        n = 0
    for idx in range(n):
        ab_eval("(function(){var a=document.querySelectorAll('li,span,div'),c=-1;"
                "for(var i=0;i<a.length;i++){if((a[i].textContent||'').trim()==='被引证'&&a[i].offsetParent){c++;"
                "if(c===%d){a[i].click();return 'ok';}}}return 'no';})()" % idx)
        nap(WAIT_CITE)
    nap((1.0, 1.8))
    # 收集所有 patcitedinfos 响应，按 anId 映射次数
    out = _ab("network", "requests", "--filter", "patcitedinfos", "--json", timeout=60)
    try:
        reqs = json.loads(out).get("data", {}).get("requests", [])
    except Exception:
        return {}
    mapping = {}
    for r in reqs:
        det = _ab("network", "request", str(r.get("requestId")), "--json", timeout=60)
        try:
            data = json.loads(det).get("data", {})
            an = (json.loads(data.get("postData") or "{}").get("param") or {}).get("anId")
            tc = (json.loads(data.get("responseBody") or "{}").get("t") or {}).get("pagination", {}).get("totalCount")
            if an is not None and tc is not None:
                mapping[an] = tc
        except Exception:
            continue
    return mapping


def do_search_once(company):
    goto_search_page()
    if ab_eval("(function(){var el=document.querySelector('input[placeholder*=\"智能识别检索\"],textarea[placeholder*=\"智能识别检索\"]');"
               "if(el){el.focus();el.value='';return 'ok';}return 'no';})()") != "ok":
        return False
    nap(WAIT_TINY)
    ab_type(company)
    nap(WAIT_TINY)
    net_clear()
    ab_press("Enter")
    nap(WAIT_SEARCH)
    return "retrieveList" in ab_eval("location.href")


def do_search(company):
    """检索并导航到结果页；失败时检测风控/登录态并退避重试。"""
    for attempt in range(3):
        if do_search_once(company):
            return True
        h = page_health()
        if h == "ok":
            return False  # 不是风控，纯粹没进结果页（如检索框没找到），交由上层处理
        if not wait_recover("检索未成功", navigate=True):
            return False
    return False


def search_company(company, start_disp, end_disp, start_i, end_i, basis, with_citations):
    if not do_search(company):
        print("   ! 检索后未进入结果页，跳过"); return []

    # 服务端日期过滤
    if start_disp or end_disp:
        set_apd_filter(start_disp, end_disp)
        pd_, _ = net_last_getresults()
        if pd_:
            iv = pd_.get("apdInterval") or {}
            print(f"     服务端日期过滤 apdInterval={iv.get('startDate','')}~{iv.get('endDate','')}")

    # 设 40 条/页 —— 原地触发并抓到第 1 页
    net_clear()
    select_dropdown_option(r"条/页", PAGE_LIMIT_LABEL, scope=".el-pagination .el-select")
    nap(WAIT_ACTION)

    records, seen = [], set()

    def harvest():
        """收割当前已捕获 getResults 的新记录，返回本批新记录列表。"""
        page_recs = []
        for rec in net_all_getresults_records():
            key = rec.get("vid") or json.dumps(rec, sort_keys=True, ensure_ascii=False)[:200]
            if key in seen:
                continue
            seen.add(key); records.append(rec); page_recs.append(rec)
        return page_recs

    def attach_citations(page_recs):
        if not with_citations or not page_recs:
            return
        cmap = fetch_citations_current_page()
        for rec in page_recs:
            if rec.get("anId") in cmap:
                rec["_citedCount"] = cmap[rec["anId"]]

    attach_citations(harvest())
    # 翻页直到没有下一页
    for _ in range(MAX_PAGES):
        net_clear()
        if not click_next_page():
            break
        nap(WAIT_ACTION)
        page_recs = harvest()
        if not page_recs:
            # 区分“真的没有下一页”与“被风控/掉登录导致空”
            if page_health() != "ok":
                if wait_recover("翻页无数据", navigate=False):
                    net_clear(); nap(WAIT_ACTION)
                    page_recs = harvest()
            if not page_recs:
                break
        attach_citations(page_recs)
    else:
        print(f"     ⚠ 达到翻页上限 {MAX_PAGES}，可能未取完")
    return records


# ── run ───────────────────────────────────────────────────────────────────────
def read_companies(path):
    df = pd.read_excel(path, header=0)
    if df.shape[1] == 0:
        sys.exit("✗ 输入 Excel 没有列")
    return [str(x).strip() for x in df[df.columns[0]].dropna().tolist() if str(x).strip()]


def cmd_run(args):
    ensure_cdp()
    switch_to_cnipa_tab()
    goto_search_page()
    if not check_logged_in():
        sys.exit("✗ 未登录。请在该 Chrome 里登录 CNIPA 后重试。")
    print("✓ 已登录")

    start_disp, start_i = fmt_date(args.start)
    end_disp, end_i = fmt_date(args.end)
    if start_i or end_i:
        print(f"→ 时间过滤(申请日): [{start_i or '不限'}, {end_i or '不限'}]")
    else:
        print("⚠ 未设时间范围：将抓取每家公司的全部专利（大公司可能极多）")
    companies = read_companies(args.input)
    print(f"→ {len(companies)} 家公司")

    raw_path = args.output + ".raw.jsonl"
    progress_path = args.output + ".progress.txt"

    # ── 断点续抓：读取已完成公司，并从备份恢复已抓数据 ──
    done = set()
    if os.path.exists(progress_path):
        with open(progress_path, encoding="utf-8") as f:
            done = {ln.strip() for ln in f if ln.strip()}

    all_rows = []
    if done:
        print(f"↻ 断点续抓：检测到已完成 {len(done)} 家，将跳过并恢复其数据")
        # 从 raw 备份重建已完成公司的行，并把 raw 清理成只含已完成公司（避免重复）
        if os.path.exists(raw_path):
            keep_lines = []
            with open(raw_path, encoding="utf-8") as f:
                for ln in f:
                    try:
                        obj = json.loads(ln)
                    except Exception:
                        continue
                    if obj.get("company") in done:
                        keep_lines.append(ln if ln.endswith("\n") else ln + "\n")
                        all_rows.append(normalize(obj["record"], obj["company"]))
            with open(raw_path, "w", encoding="utf-8") as f:
                f.writelines(keep_lines)
        if all_rows:
            pd.DataFrame(all_rows).to_excel(args.output, index=False)
    else:
        # 全新开始：清掉可能残留的旧备份，避免污染
        if os.path.exists(raw_path):
            os.remove(raw_path)

    todo = sum(1 for c in companies if c not in done)
    print(f"→ 待抓 {todo} 家（共 {len(companies)} 家，已完成 {len(companies) - todo} 家）")

    for i, company in enumerate(companies, 1):
        if company in done:
            print(f"[{i}/{len(companies)}] {company} —— 已完成，跳过")
            continue
        print(f"[{i}/{len(companies)}] {company}")
        # 每家公司前先体检：异常则退避；持续无法恢复就停止（已抓数据已落盘，恢复后可继续）
        if page_health() != "ok":
            if not wait_recover("公司检索前检测到异常", navigate=True):
                print("✗ 持续异常无法恢复，已停止。已完成的公司不会重抓，恢复后重跑同一命令即可续抓。")
                break
        try:
            recs = search_company(company, start_disp, end_disp, start_i, end_i,
                                  args.date_basis, args.citations)
        except Exception as e:
            print(f"   ! 出错: {e}"); recs = []
        kept = [r for r in recs if in_range(r, start_i, end_i, args.date_basis)]
        all_rows.extend(normalize(r, company) for r in kept)
        print(f"   → 抓取 {len(recs)} 条，过滤后 {len(kept)} 条（累计 {len(all_rows)}）")
        if all_rows:
            pd.DataFrame(all_rows).to_excel(args.output, index=False)
        with open(raw_path, "a", encoding="utf-8") as f:
            for r in kept:
                f.write(json.dumps({"company": company, "record": r}, ensure_ascii=False) + "\n")
        # 标记本公司完成（写在数据落盘之后，确保只有完整处理过的才算完成）
        with open(progress_path, "a", encoding="utf-8") as f:
            f.write(company + "\n")
        done.add(company)
        nap(DELAY_COMPANIES)

    if all_rows:
        print(f"\n✓ 完成，共 {len(all_rows)} 条 → {args.output}")
        if len(done) >= len(set(companies)):
            print(f"  全部 {len(set(companies))} 家已抓完。如需重新抓取，请删除 {progress_path}")
    else:
        print("\n⚠ 没抓到数据。")


def main():
    p = argparse.ArgumentParser(description="CNIPA 专利检索爬虫")
    sub = p.add_subparsers(dest="cmd", required=True)
    rp = sub.add_parser("run", help="按 Excel 公司名批量检索并汇总")
    rp.add_argument("--input", required=True, help="输入 Excel（第一列为公司名）")
    rp.add_argument("--output", default="专利汇总.xlsx", help="输出 Excel")
    rp.add_argument("--start", help="起始日期 如 2020-01-01")
    rp.add_argument("--end", help="截止日期 如 2023-12-31")
    rp.add_argument("--date-basis", default="申请日", choices=["申请日", "公开日"],
                    help="客户端兜底过滤依据（服务端过滤固定按申请日）")
    rp.add_argument("--citations", action="store_true",
                    help="额外抓取每条专利的被引证次数（逐条点击，很慢、请求多，慎用）")
    args = p.parse_args()
    if args.cmd == "run":
        cmd_run(args)


if __name__ == "__main__":
    main()
