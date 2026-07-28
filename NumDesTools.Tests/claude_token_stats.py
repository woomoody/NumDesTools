"""
Claude Code Token 使用统计 — 生成 HTML 报告并自动用浏览器打开
按实际使用的模型分别计价（$/MTok）
支持 --date 参数筛选日期范围
"""
import os, json, sys, subprocess, webbrowser, argparse, sqlite3
from collections import defaultdict
from datetime import datetime, date, timedelta

sys.stdout.reconfigure(encoding='utf-8')

parser = argparse.ArgumentParser()
parser.add_argument('--date', default='today', help='today / 2026-07-20 / 2026-07-15..2026-07-20')
args = parser.parse_args()

today = date.today()
if args.date == 'today':
    date_start = today.isoformat()
    date_end   = today.isoformat()
    date_label = f"今天（{today.isoformat()}）"
elif '..' in args.date:
    parts = args.date.split('..')
    date_start = parts[0].strip()
    date_end   = parts[1].strip()
    date_label = f"{date_start} ~ {date_end}"
else:
    date_start = args.date.strip()
    date_end   = date_start
    date_label = date_start

# ── 模型价格（每天从 JSON 刷新一次）──
_PRICE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'model_prices.json')
_SNAP_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'token_stats_history.json')

_DEFAULT_PRICES = [
    # Anthropic 官方价（$/MTok，prefix 顺序匹配——精确在前，避免 opus-4-1 被 opus-4 误匹配）
    {'prefix': 'claude-fable',    'input': 10.00, 'output': 50.00, 'cache_read': 1.00,  'cache_write': 12.50},
    {'prefix': 'claude-mythos',   'input': 10.00, 'output': 50.00, 'cache_read': 1.00,  'cache_write': 12.50},
    {'prefix': 'claude-opus-4-1', 'input': 15.00, 'output': 75.00, 'cache_read': 1.50,  'cache_write': 18.75},  # 4.1 老版
    {'prefix': 'claude-opus-4',   'input':  5.00, 'output': 25.00, 'cache_read': 0.50,  'cache_write':  6.25},  # 4.5+ (4-5/4-6/4-7/4-8)
    {'prefix': 'claude-opus-5',   'input':  5.00, 'output': 25.00, 'cache_read': 0.50,  'cache_write':  6.25},
    {'prefix': 'claude-sonnet-5', 'input':  2.00, 'output': 10.00, 'cache_read': 0.20,  'cache_write':  2.50},  # 9/1 前促销价
    {'prefix': 'claude-sonnet-4', 'input':  3.00, 'output': 15.00, 'cache_read': 0.30,  'cache_write':  3.75},
    {'prefix': 'claude-haiku-4',  'input':  1.00, 'output':  5.00, 'cache_read': 0.10,  'cache_write':  1.25},
    {'prefix': 'claude-haiku-3',  'input':  0.80, 'output':  4.00, 'cache_read': 0.08,  'cache_write':  1.00},
    # ── 以下价格：input/output 用户填实价；cache 按各家规则估（无官方数据）──
    # DeepSeek/GLM/Kimi: cache_read≈0.1×input, cache_write≈1.25×input；GPT: cache_read≈0.5×input, cache_write≈1.25×input
    {'prefix': 'glm-5.2',           'input': 0.74, 'output': 2.47,  'cache_read': 0.07, 'cache_write': 0.93},
    {'prefix': 'glm-5.1',           'input': 0.57, 'output': 2.29,  'cache_read': 0.06, 'cache_write': 0.71},
    {'prefix': 'deepseek-v4-pro',   'input': 1.06, 'output': 2.02, 'cache_read': 0.11, 'cache_write': 1.06},
    {'prefix': 'deepseek-v4-flash', 'input': 0.09, 'output': 0.18, 'cache_read': 0.01, 'cache_write': 0.09},
    {'prefix': 'gpt-5.6-sol',       'input': 5.00, 'output': 30.00,'cache_read': 2.50, 'cache_write': 6.25},
    {'prefix': 'gpt-5.4-mini',      'input': 0.75, 'output': 4.50, 'cache_read': 0.38, 'cache_write': 0.94},
    {'prefix': 'gpt-5.3-codex',     'input': 1.75, 'output': 14.00,'cache_read': 0.88, 'cache_write': 2.19},
    {'prefix': 'kimi',              'input': 2.94, 'output': 14.71,'cache_read': 0.29, 'cache_write': 3.68},
]
_DEFAULT_FALLBACK = {'input': 5.00, 'output': 25.00, 'cache_read': 0.50, 'cache_write': 6.25}

def _load_prices():
    ts = today.isoformat()
    if os.path.exists(_PRICE_FILE):
        try:
            with open(_PRICE_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if data.get('date') == ts:
                pu = data.get('price_updated') or data.get('date')
                try:
                    age = (today - date.fromisoformat(pu)).days
                    if age >= 7: print(f"  [提醒] 价格表已 {age} 天未更新，考虑更新 _DEFAULT_PRICES（搜官方价后改此处）")
                except: pass
                return data['prices'], data['fallback']
        except: pass
    prices = _DEFAULT_PRICES[:]
    fallback = dict(_DEFAULT_FALLBACK)
    with open(_PRICE_FILE, 'w', encoding='utf-8') as f:
        json.dump({'date': ts, 'price_updated': ts, 'prices': prices, 'fallback': fallback}, f, ensure_ascii=False, indent=2)
    print(f"  prices updated ({ts}), {len(prices)} models")
    return prices, fallback

MODEL_PRICES, _PRICE_FALLBACK = _load_prices()

def _model_price(model: str):
    m = (model or '').lower()
    for p in MODEL_PRICES:
        if m.startswith(p['prefix']):
            return (p['input'], p['output'], p['cache_read'], p['cache_write'])
    return (_PRICE_FALLBACK['input'], _PRICE_FALLBACK['output'],
            _PRICE_FALLBACK['cache_read'], _PRICE_FALLBACK['cache_write'])

def calc_cost(inp, out, cr, cw, model=''):
    pi, po, pcr, pcw = _model_price(model)
    return (inp * pi + out * po + cr * pcr + cw * pcw) / 1_000_000

def _load_snap():
    """读完整快照（含 frozen_date）。"""
    if not os.path.exists(_SNAP_FILE): return None
    try:
        with open(_SNAP_FILE, 'r', encoding='utf-8') as f: return json.load(f)
    except Exception: return None

def _frozen_unix(frozen_str):
    """frozen_date(YYYY-MM-DD) -> unix 秒（UTC 当天 00:00）。"""
    from datetime import timezone
    try: return datetime.fromisoformat(frozen_str).replace(tzinfo=timezone.utc).timestamp()
    except: return 0

def _collect_cc(frozen_str):
    """CC: ~/.claude/projects jsonl（增量：mtime 筛 + 行 timestamp>frozen）。"""
    records = []
    mtime_cut = _frozen_unix(frozen_str) - 2 * 86400
    for BASE, prefix in BASES:
        if not os.path.isdir(BASE): continue
        for proj in sorted(os.listdir(BASE)):
            pp = os.path.join(BASE, proj)
            if not os.path.isdir(pp): continue
            proj_key = f"{prefix}{proj}"
            for dp, _, files in os.walk(pp):
                if os.sep + 'subagents' in dp or os.sep + 'workflows' in dp: continue
                for f in sorted(files):
                    if not f.endswith('.jsonl'): continue
                    fp = os.path.join(dp, f)
                    try:
                        if os.path.getmtime(fp) < mtime_cut: continue
                        with open(fp, 'r', encoding='utf-8') as fh:
                            for line in fh:
                                line = line.strip()
                                if not line: continue
                                try: obj = json.loads(line)
                                except: continue
                                ts = obj.get('timestamp') or obj.get('ts') or obj.get('created_at')
                                if not ts:
                                    msg = obj.get('message', {})
                                    if isinstance(msg, dict): ts = msg.get('timestamp') or msg.get('created_at')
                                if not ts: continue
                                try:
                                    date_str = datetime.fromtimestamp(ts).strftime('%Y-%m-%d') if isinstance(ts, (int, float)) else str(ts)[:10]
                                except: continue
                                if date_str <= frozen_str: continue
                                usage = obj.get('usage') or (isinstance(obj.get('message'), dict) and obj['message'].get('usage'))
                                if not isinstance(usage, dict): continue
                                inp = usage.get('input_tokens',0) or 0
                                out = usage.get('output_tokens',0) or 0
                                cr  = usage.get('cache_read_input_tokens',0) or 0
                                cw  = usage.get('cache_creation_input_tokens',0) or 0
                                if inp+out+cr+cw == 0: continue
                                if out < 200 and cr == 0 and cw == 0: continue
                                model = (obj.get('model') or (obj.get('message') or {}).get('model') or '<empty>')
                                records.append((date_str, model, inp, out, cr, cw, proj_key))
                    except Exception as e:
                        print(f'  [warn] {fp}: {e}')
    return records

def _collect_hermes(frozen_str):
    """hermes: state.db sessions 表（started_at Unix 浮点秒，WHERE >frozen_unix）。"""
    records = []
    db = os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'hermes', 'state.db')
    if not os.path.exists(db): return records
    fu = _frozen_unix(frozen_str)
    try:
        c = sqlite3.connect(db)
        for model, ts, inp, out, cr, cw, cwd in c.execute(
            "SELECT model, started_at, input_tokens, output_tokens, cache_read_tokens, cache_write_tokens, cwd FROM sessions WHERE started_at > ?", (fu,)):
            if not model or not ts: continue
            try: date_str = datetime.fromtimestamp(ts).strftime('%Y-%m-%d')
            except: continue
            if date_str <= frozen_str: continue
            inp, out, cr, cw = inp or 0, out or 0, cr or 0, cw or 0
            if inp+out+cr+cw == 0: continue
            records.append((date_str, model, inp, out, cr, cw, f'[hermes]{cwd or "None"}'))
    except Exception as e: print(f'  [warn] hermes db: {e}')
    return records

def _collect_omp(frozen_str):
    """omp: ~/.omp/agent/sessions jsonl（message.usage 简写 input/output/cacheRead/cacheWrite）。"""
    records = []
    base = os.path.expanduser('~/.omp/agent/sessions')
    if not os.path.isdir(base): return records
    mtime_cut = _frozen_unix(frozen_str) - 2 * 86400
    for dp, _, files in os.walk(base):
        for f in sorted(files):
            if not f.endswith('.jsonl'): continue
            fp = os.path.join(dp, f)
            try:
                if os.path.getmtime(fp) < mtime_cut: continue
                with open(fp, 'r', encoding='utf-8') as fh:
                    for line in fh:
                        line = line.strip()
                        if not line: continue
                        try: obj = json.loads(line)
                        except: continue
                        msg = obj.get('message')
                        if not isinstance(msg, dict): continue
                        u = msg.get('usage')
                        if not isinstance(u, dict): continue
                        inp = u.get('input') or u.get('input_tokens') or u.get('prompt_tokens') or 0
                        out = u.get('output') or u.get('output_tokens') or u.get('completion_tokens') or 0
                        cr = u.get('cacheRead') or u.get('cache_read_input_tokens') or u.get('cache_read_tokens') or 0
                        cw = u.get('cacheWrite') or u.get('cache_creation_input_tokens') or u.get('cache_write_tokens') or 0
                        if inp+out+cr+cw == 0: continue
                        model = msg.get('model') or obj.get('model') or '<empty>'
                        ts = obj.get('timestamp')
                        date_str = str(ts)[:10] if ts else ''
                        if not date_str or date_str <= frozen_str: continue
                        records.append((date_str, model, inp, out, cr, cw, '[omp]'))
            except Exception as e: print(f'  [warn] {fp}: {e}')
    return records

def _collect_opencode(frozen_str):
    """opencode: opencode.db session 表（time_created Unix 毫秒 /1000，model 是 JSON 字符串取 id）。"""
    records = []
    db = os.path.join(os.path.expanduser('~'), '.local', 'share', 'opencode', 'opencode.db')
    if not os.path.exists(db): return records
    fu = _frozen_unix(frozen_str)
    try:
        c = sqlite3.connect(db)
        for m_json, tc, inp, out, cr, cw, d in c.execute(
            "SELECT model, time_created, tokens_input, tokens_output, tokens_cache_read, tokens_cache_write, directory FROM session WHERE time_created/1000.0 > ?", (fu,)):
            if not m_json or not tc: continue
            try: model = json.loads(m_json).get('id', m_json) if m_json.startswith('{') else m_json
            except: model = m_json
            try: date_str = datetime.fromtimestamp(tc/1000.0).strftime('%Y-%m-%d')
            except: continue
            if date_str <= frozen_str: continue
            inp, out, cr, cw = inp or 0, out or 0, cr or 0, cw or 0
            if inp+out+cr+cw == 0: continue
            records.append((date_str, model, inp, out, cr, cw, f'[opencode]{d or "None"}'))
    except Exception as e: print(f'  [warn] opencode db: {e}')
    return records

def _save_snap(daily, model_daily, proj_daily, frozen_date):
    tmp = _SNAP_FILE + '.tmp'
    try:
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump({'daily': dict(daily), 'model_daily': {d: dict(md) for d, md in model_daily.items()},
                       'proj_daily': {p: dict(pd) for p, pd in proj_daily.items() if pd},
                       'frozen_date': frozen_date, 'updated': datetime.now().isoformat()},
                      f, ensure_ascii=False)
        os.replace(tmp, _SNAP_FILE)
    except Exception as e:
        print(f'  [warn] 快照保存失败: {e}')
        if os.path.exists(tmp):
            try: os.remove(tmp)
            except OSError: pass

def cn_num(n):
    if n >= 1_0000_0000: return f'{n/1_0000_0000:.2f}亿'
    if n >= 1_0000:      return f'{n/1_0000:.1f}万'
    return f'{n:,}'

BASES = [(os.path.expanduser(r'~/.claude/projects'), '[local]')]

# ── 数据采集（4 源统一，按 model 不分 harness）+ 增量 frozen_date ────────────────
_zero = lambda: {'input':0,'output':0,'cache_read':0,'cache_write':0,'cost':0.0}
daily       = defaultdict(_zero)
monthly     = defaultdict(_zero)
proj_daily  = defaultdict(lambda: defaultdict(_zero))
model_daily = defaultdict(lambda: defaultdict(_zero))  # model_daily[date][model]
total_msgs = skipped = 0

snap = _load_snap()
# frozen_cc = CC 在快照里的最后日期（proj_daily [local] 项目 max），不用 frozen_date——防其他源拉高 daily max 导致 CC 漏扫边界
if snap:
    cc_dates = [d for p, pd in (snap.get('proj_daily') or {}).items() if p.startswith('[local]') for d in pd]
    frozen_cc = max(cc_dates) if cc_dates else '1970-01-01'
else:
    frozen_cc = '1970-01-01'
# 其他源：有 frozen_date 用(增量)，首次无则全量
frozen_other = (snap.get('frozen_date') if snap else None) or '1970-01-01'
print(f"  frozen: cc={frozen_cc} other={frozen_other}")

for fn, fz in ((_collect_cc, frozen_cc), (_collect_hermes, frozen_other), (_collect_omp, frozen_other), (_collect_opencode, frozen_other)):
    for date_str, model, inp, out, cr, cw, proj_key in fn(fz):
        cost = calc_cost(inp, out, cr, cw, model)
        total_msgs += 1
        for d in (daily[date_str], monthly[date_str[:7]], proj_daily[proj_key][date_str], model_daily[date_str][model]):
            d['input'] += inp; d['output'] += out; d['cache_read'] += cr; d['cache_write'] += cw; d['cost'] += cost

# ── 合并历史快照（累加：snap 不同源 + fresh 不同源不重复；frozen 控制各源不重扫同日期）──
if snap:
    for d, sv in snap.get('daily', {}).items():
        dv = daily[d]
        for k in ('input','output','cache_read','cache_write','cost'): dv[k] += sv.get(k, 0)
    for d, md in snap.get('model_daily', {}).items():
        for m, v in md.items():
            mv = model_daily[d][m]
            for k in ('input','output','cache_read','cache_write','cost'): mv[k] += v.get(k, 0)
    for proj, pd in snap.get('proj_daily', {}).items():
        for d, v in pd.items():
            pv = proj_daily[proj][d]
            for k in ('input','output','cache_read','cache_write','cost'): pv[k] += v.get(k, 0)

# monthly 从合并后 daily 重算（含历史）
monthly.clear()
for d, v in daily.items():
    m = monthly[d[:7]]
    m['input'] += v['input']; m['output'] += v['output']
    m['cache_read'] += v['cache_read']; m['cache_write'] += v['cache_write']
    m['cost'] += v['cost']

# ── 固化：<today 的进快照（昨天及以前已稳定），frozen_date=today-1 ──
today_iso = today.isoformat()
frozen_new = (today - timedelta(days=1)).isoformat()
new_daily = {d: v for d, v in daily.items() if d < today_iso}
new_md = {d: dict(md) for d, md in model_daily.items() if d < today_iso}
new_pd = {p: {d: v for d, v in pd.items() if d < today_iso} for p, pd in proj_daily.items()}
new_pd = {p: pd for p, pd in new_pd.items() if pd}
_save_snap(new_daily, new_md, new_pd, frozen_new)
print(f"  固化到 {frozen_new}（{len(new_daily)} 天历史进快照）")

# ── 汇总计算 ──────────────────────────────────────────────────────────────────
grand_in = grand_out = grand_cr = grand_cw = grand_cost = 0
for v in daily.values():
    grand_in += v['input']; grand_out += v['output']
    grand_cr += v['cache_read']; grand_cw += v['cache_write']
    grand_cost += v['cost']

def period_stats(days):
    start = (today - timedelta(days=days-1)).isoformat() if days else '0000-00-00'
    si=so=scr=scw=dc=0; sc=0.0
    for d,v in daily.items():
        if d >= start:
            si+=v['input']; so+=v['output']; scr+=v['cache_read']; scw+=v['cache_write']
            sc+=v['cost']; dc+=1
    return dc, si, so, scr, scw, sc

# 填充完整日期轴
if daily:
    raw = sorted(daily.keys())
    d0 = date.fromisoformat(raw[0])
    d1 = today
    all_dates = [(d0+timedelta(days=i)).isoformat() for i in range((d1-d0).days+1)]
else:
    all_dates = []

empty = {'input':0,'output':0,'cache_read':0,'cache_write':0}
chart_dates   = all_dates
chart_output  = [daily.get(d,empty)['output']/1000 for d in chart_dates]
chart_input   = [daily.get(d,empty)['input']/1000  for d in chart_dates]
chart_cr      = [daily.get(d,empty)['cache_read']/1000 for d in chart_dates]
chart_cost    = [round(daily.get(d, {'cost':0.0})['cost'], 2) for d in chart_dates]

# 每日明细行
detail_rows = ''
for d in all_dates:
    v = daily.get(d) or {'input':0,'output':0,'cache_read':0,'cache_write':0,'cost':0.0}
    i,o,cr,cw,c = v['input'],v['output'],v['cache_read'],v['cache_write'],v['cost']
    detail_rows += f'<tr><td>{d}</td><td>{cn_num(i)}</td><td>{cn_num(o)}</td><td>{cn_num(cr)}</td><td>{cn_num(cw)}</td><td>{cn_num(i+o)}</td><td>{cn_num(i+o+cr+cw)}</td><td>${c:.2f}</td></tr>\n'

# 项目汇总行
proj_rows = ''
proj_list = []
for proj, pd in proj_daily.items():
    pi=po=pcr=pcw=0; pc=0.0
    for v in pd.values():
        pi+=v['input']; po+=v['output']; pcr+=v['cache_read']; pcw+=v['cache_write']; pc+=v['cost']
    if pi+po+pcr+pcw == 0: continue
    proj_list.append((proj, pi, po, pcr, pcw, pc))
proj_list.sort(key=lambda x: -x[5])
for proj,pi,po,pcr,pcw,pc in proj_list:
    short = proj[-60:] if len(proj)>60 else proj
    proj_rows += f'<tr><td title="{proj}">{short}</td><td>{cn_num(pi)}</td><td>{cn_num(po)}</td><td>{cn_num(pcr)}</td><td>{cn_num(pcw)}</td><td>${pc:.2f}</td></tr>\n'

# 模型汇总行（全历史，扁平化聚合）
model_rows = ''
model_flat = defaultdict(_zero)
for d, md in model_daily.items():
    for m, v in md.items():
        for k in ('input','output','cache_read','cache_write','cost'):
            model_flat[m][k] += v[k]
model_list = []
for m, v in model_flat.items():
    mi,mo,mcr,mcw,mc = v['input'],v['output'],v['cache_read'],v['cache_write'],v['cost']
    if mi+mo+mcr+mcw == 0: continue
    model_list.append((m, mi, mo, mcr, mcw, mc))
model_list.sort(key=lambda x: -x[1])
for m,mi,mo,mcr,mcw,mc in model_list:
    model_rows += f'<tr><td>{m}</td><td>{cn_num(mi)}</td><td>{cn_num(mo)}</td><td>{cn_num(mcr)}</td><td>{cn_num(mcw)}</td><td>{cn_num(mi+mo+mcr+mcw)}</td><td>${mc:.2f}</td></tr>\n'

# ── 按日期筛选的模型汇总 ──
model_date_rows = ''
model_date_list = []
filtered_model = defaultdict(_zero)
for d, md in model_daily.items():
    if date_start <= d <= date_end:
        for m, v in md.items():
            for k in ('input','output','cache_read','cache_write','cost'):
                filtered_model[m][k] += v[k]
for m, v in filtered_model.items():
    mi,mo,mcr,mcw,mc = v['input'],v['output'],v['cache_read'],v['cache_write'],v['cost']
    if mi+mo+mcr+mcw == 0: continue
    model_date_list.append((m, mi, mo, mcr, mcw, mc))
model_date_list.sort(key=lambda x: -x[1])
for m,mi,mo,mcr,mcw,mc in model_date_list:
    model_date_rows += f'<tr><td>{m}</td><td>{cn_num(mi)}</td><td>{cn_num(mo)}</td><td>{cn_num(mcr)}</td><td>{cn_num(mcw)}</td><td>{cn_num(mi+mo+mcr+mcw)}</td><td>${mc:.2f}</td></tr>\n'

# 汇总卡数据
dc7,si7,so7,scr7,scw7,cost7   = period_stats(7)
dc30,si30,so30,scr30,scw30,cost30 = period_stats(30)

def card(title, days, dc, si, so, scr, scw, cost):
    quota = si+so+scr+scw
    return f'''
    <div class="card">
      <div class="card-title">{title}</div>
      <div class="card-cost">${cost:.2f}</div>
      <div class="card-sub">有效天数 {dc} 天</div>
      <table class="card-table">
        <tr><td>input</td><td>{cn_num(si)}</td></tr>
        <tr><td>output</td><td>{cn_num(so)}</td></tr>
        <tr><td>缓存读</td><td>{cn_num(scr)}</td></tr>
        <tr><td>缓存写</td><td>{cn_num(scw)}</td></tr>
        <tr class="sep"><td>实计(in+out)</td><td>{cn_num(si+so)}</td></tr>
        <tr><td>配额消耗(全)</td><td>{cn_num(quota)}</td></tr>
      </table>
    </div>'''

def month_card(label, ym):
    v = monthly.get(ym, {'input':0,'output':0,'cache_read':0,'cache_write':0,'cost':0.0})
    mi,mo,mcr,mcw,mc = v['input'],v['output'],v['cache_read'],v['cache_write'],v['cost']
    days_in = sum(1 for d in daily if d.startswith(ym))
    return card(f'{label}（{ym}）', None, days_in, mi, mo, mcr, mcw, mc)

this_month = today.strftime('%Y-%m')
last_month = (today.replace(day=1) - timedelta(days=1)).strftime('%Y-%m')

# 月度明细行
month_rows = ''
for ym in sorted(monthly.keys(), reverse=True):
    v = monthly[ym]
    mi,mo,mcr,mcw,mc = v['input'],v['output'],v['cache_read'],v['cache_write'],v['cost']
    days_in = sum(1 for d in daily if d.startswith(ym))
    month_rows += (f'<tr><td>{ym}</td><td>{days_in}</td>'
                   f'<td>{cn_num(mi)}</td><td>{cn_num(mo)}</td>'
                   f'<td>{cn_num(mcr)}</td><td>{cn_num(mcw)}</td>'
                   f'<td>{cn_num(mi+mo)}</td><td>{cn_num(mi+mo+mcr+mcw)}</td>'
                   f'<td>${mc:.2f}</td></tr>\n')

cards = (card('最近 7 天', 7, dc7, si7, so7, scr7, scw7, cost7)
       + card('最近 30 天', 30, dc30, si30, so30, scr30, scw30, cost30)
       + month_card('本月', this_month)
       + month_card('上月', last_month)
       + card('历史累计', None, len(daily), grand_in, grand_out, grand_cr, grand_cw, grand_cost))

import json as _json
labels_js   = _json.dumps(chart_dates)
output_js   = _json.dumps(chart_output)
input_js    = _json.dumps(chart_input)
cr_js       = _json.dumps(chart_cr)
cost_js     = _json.dumps(chart_cost)

# 全量 model_daily 嵌进 HTML，供交互式区间筛选（按模型消耗表）
_model_daily_plain = {
    d: {m: {'input': v['input'], 'output': v['output'],
           'cache_read': v['cache_read'], 'cache_write': v['cache_write'],
           'cost': round(v['cost'], 4)} for m, v in md.items()}
    for d, md in model_daily.items()
}
model_daily_js = _json.dumps(_model_daily_plain)
_daily_plain = {d: {'input': v['input'], 'output': v['output'], 'cache_read': v['cache_read'], 'cache_write': v['cache_write'], 'cost': round(v['cost'], 4)} for d, v in daily.items()}
daily_js = _json.dumps(_daily_plain)
_today_iso = today.isoformat()
_model_price_map = {m: dict(zip(['in','out','cr','cw'], _model_price(m))) for m in {mm for _md in _model_daily_plain.values() for mm in _md}}
model_price_js = _json.dumps(_model_price_map)
_all_dates_sorted = sorted(_model_daily_plain.keys())
_date_min = _all_dates_sorted[0] if _all_dates_sorted else today.isoformat()
_date_max = _all_dates_sorted[-1] if _all_dates_sorted else today.isoformat()
model_date_js = r'''const MODEL_DAILY = ''' + model_daily_js + r''';
const MODEL_PRICE = ''' + model_price_js + r''';
const DAILY = ''' + daily_js + r''';
const TODAY = "''' + _today_iso + r'''";
const DATE_MIN = "''' + _date_min + r'''";
const DATE_MAX = "''' + _date_max + r'''";
let mdChart=null, barChart=null, crChart=null, costChart=null;
const gridColor='rgba(255,255,255,0.06)', tickColor='#666';
function fmtN(n){ if(n>=1e8) return (n/1e8).toFixed(2)+'亿'; if(n>=1e4) return (n/1e4).toFixed(1)+'万'; return n.toLocaleString(); }
function offsetDate(d,n){ var dt=new Date(d+'T00:00:00Z'); dt.setUTCDate(dt.getUTCDate()+n); return dt.toISOString().substring(0,10); }
function rangeKeys(start,end){ return Object.keys(DAILY).filter(function(d){return d>=start&&d<=end;}).sort(); }
function renderMdTable(start, end, sortBy){
  var agg={};
  rangeKeys(start,end).forEach(function(d){ Object.keys(MODEL_DAILY[d]||{}).forEach(function(m){ var v=MODEL_DAILY[d][m]; if(!agg[m])agg[m]={input:0,output:0,cache_read:0,cache_write:0,cost:0}; agg[m].input+=v.input;agg[m].output+=v.output;agg[m].cache_read+=v.cache_read;agg[m].cache_write+=v.cache_write;agg[m].cost+=v.cost; }); });
  var arr=Object.keys(agg).map(function(m){var x=agg[m];x.m=m;x.quota=x.input+x.output+x.cache_read+x.cache_write;return x;}).filter(function(x){return x.quota>0;});
  arr.sort(function(a,b){return (b[sortBy]||0)-(a[sortBy]||0);});
  var tbody=document.getElementById('mdTbody');
  tbody.innerHTML = arr.length===0 ? '<tr><td colspan="8" style="text-align:center;color:#666;">该时间段无数据</td></tr>' : arr.map(function(x){ return '<tr><td>'+x.m+'</td><td>'+(MODEL_PRICE[x.m]?(MODEL_PRICE[x.m].in.toFixed(2)+'/'+MODEL_PRICE[x.m].out.toFixed(2)):'-')+'</td><td>'+fmtN(x.input)+'</td><td>'+fmtN(x.output)+'</td><td>'+fmtN(x.cache_read)+'</td><td>'+fmtN(x.cache_write)+'</td><td>'+fmtN(x.quota)+'</td><td>$'+x.cost.toFixed(2)+'</td></tr>'; }).join('');
  document.getElementById('mdTitle').textContent='📅 '+start+' ~ '+end+' · 按模型消耗 Token';
  if(mdChart)mdChart.destroy();
  mdChart=new Chart(document.getElementById('mdChart'),{type:'bar',data:{labels:arr.map(function(x){return x.m;}),datasets:[{label:'费用 USD',data:arr.map(function(x){return +x.cost.toFixed(2);}),backgroundColor:'rgba(245,166,35,0.7)'}]},options:{responsive:true,plugins:{legend:{labels:{color:'#aaa'}}},scales:{x:{ticks:{color:tickColor,maxRotation:45},grid:{color:gridColor}},y:{ticks:{color:tickColor,callback:function(v){return '$'+v;}},grid:{color:gridColor}}}}});
}
function renderDetailBody(start,end){
  var keys=rangeKeys(start,end).reverse();
  var tbody=document.getElementById('detailBody');
  if(keys.length===0){ tbody.innerHTML='<tr><td colspan="8" style="text-align:center;color:#666;">该时间段无数据</td></tr>'; return; }
  tbody.innerHTML=keys.map(function(d){ var v=DAILY[d]; var sum=v.input+v.output, quota=sum+v.cache_read+v.cache_write; return '<tr><td>'+d+'</td><td>'+fmtN(v.input)+'</td><td>'+fmtN(v.output)+'</td><td>'+fmtN(v.cache_read)+'</td><td>'+fmtN(v.cache_write)+'</td><td>'+fmtN(sum)+'</td><td>'+fmtN(quota)+'</td><td>$'+v.cost.toFixed(2)+'</td></tr>'; }).join('');
}
function renderMainCharts(start,end){
  var keys=rangeKeys(start,end), labels=keys;
  var out=keys.map(function(d){return DAILY[d].output/1000;}), inp=keys.map(function(d){return DAILY[d].input/1000;}), cr=keys.map(function(d){return DAILY[d].cache_read/1000;}), cost=keys.map(function(d){return DAILY[d].cost;});
  if(barChart)barChart.destroy();
  barChart=new Chart(document.getElementById('barChart'),{type:'bar',data:{labels:labels,datasets:[{label:'output (K)',data:out,backgroundColor:'rgba(168,216,234,0.75)',order:1},{label:'input (K)',data:inp,backgroundColor:'rgba(100,149,237,0.55)',order:2}]},options:{responsive:true,plugins:{legend:{labels:{color:'#aaa'}}},scales:{x:{ticks:{color:tickColor,maxRotation:45},grid:{color:gridColor}},y:{ticks:{color:tickColor},grid:{color:gridColor}}}}});
  if(crChart)crChart.destroy();
  crChart=new Chart(document.getElementById('crChart'),{type:'bar',data:{labels:labels,datasets:[{label:'缓存读 (K)',data:cr,backgroundColor:'rgba(245,166,35,0.6)'}]},options:{responsive:true,plugins:{legend:{labels:{color:'#aaa'}}},scales:{x:{ticks:{color:tickColor,maxRotation:45},grid:{color:gridColor}},y:{ticks:{color:tickColor},grid:{color:gridColor}}}}});
  if(costChart)costChart.destroy();
  costChart=new Chart(document.getElementById('costChart'),{type:'line',data:{labels:labels,datasets:[{label:'费用 USD',data:cost,borderColor:'#f5a623',backgroundColor:'rgba(245,166,35,0.15)',pointRadius:3,tension:0.3,fill:true}]},options:{responsive:true,plugins:{legend:{labels:{color:'#aaa'}}},scales:{x:{ticks:{color:tickColor,maxRotation:45},grid:{color:gridColor}},y:{ticks:{color:tickColor,callback:function(v){return '$'+v;}},grid:{color:gridColor}}}}});
}
function applyRange(start,end){
  document.getElementById('rangeStart').value=start;
  document.getElementById('rangeEnd').value=end;
  document.getElementById('rangeLabel').textContent='📅 '+start+' ~ '+end;
  renderMainCharts(start,end);
  renderDetailBody(start,end);
  renderMdTable(start,end,document.getElementById('mdSort').value);
}
function setPreset(preset){
  var end=TODAY, start;
  if(preset==='today') start=end;
  else if(preset==='7d') start=offsetDate(end,-6);
  else if(preset==='30d') start=offsetDate(end,-29);
  else if(preset==='3m') start=offsetDate(end,-89);
  else if(preset==='1y') start=offsetDate(end,-364);
  else if(preset==='month') start=end.substring(0,8)+'01';
  document.querySelectorAll('.range-pill').forEach(function(b){b.classList.toggle('active', b.dataset.preset===preset);});
  applyRange(start,end);
}
applyRange(offsetDate(TODAY,-6), TODAY);
'''

html = f'''<!DOCTYPE html>
<html lang="zh">
<head>
<meta charset="UTF-8">
<title>Claude Code Token 统计</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4/dist/chart.umd.min.js"></script>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: "Microsoft YaHei", Arial, sans-serif; background: #1a1a2e; color: #e0e0e0; padding: 20px; }}
  h1 {{ font-size: 1.4em; margin-bottom: 4px; color: #a8d8ea; }}
  .meta {{ font-size: .85em; color: #888; margin-bottom: 20px; }}
  .cards {{ display: flex; gap: 16px; margin-bottom: 28px; flex-wrap: wrap; }}
  .card {{ background: #16213e; border-radius: 10px; padding: 16px 20px; min-width: 220px; flex: 1; }}
  .card-title {{ font-size: .9em; color: #888; margin-bottom: 4px; }}
  .card-cost {{ font-size: 2em; font-weight: bold; color: #f5a623; margin-bottom: 6px; }}
  .card-sub {{ font-size: .8em; color: #666; margin-bottom: 10px; }}
  .card-table {{ width: 100%; font-size: .85em; border-collapse: collapse; }}
  .card-table td {{ padding: 2px 0; }}
  .card-table td:last-child {{ text-align: right; color: #a8d8ea; }}
  .card-table tr.sep td {{ border-top: 1px solid #333; padding-top: 6px; }}
  .chart-box {{ background: #16213e; border-radius: 10px; padding: 16px; margin-bottom: 20px; }}
  .chart-box h2 {{ font-size: 1em; color: #888; margin-bottom: 12px; }}
  .section {{ background: #16213e; border-radius: 10px; padding: 16px; margin-bottom: 20px; }}
  .section h2 {{ font-size: 1em; color: #888; margin-bottom: 12px; }}
  table.data {{ width: 100%; border-collapse: collapse; font-size: .82em; }}
  table.data th {{ background: #0f3460; color: #a8d8ea; padding: 6px 10px; text-align: right; white-space: nowrap; }}
  table.data th:first-child {{ text-align: left; }}
  table.data td {{ padding: 5px 10px; text-align: right; border-bottom: 1px solid #222; white-space: nowrap; }}
  table.data td:first-child {{ text-align: left; color: #ccc; }}
  table.data tr:hover td {{ background: #1e2a4a; }}
  .range-bar {{ background: #16213e; border-radius: 10px; padding: 12px 16px; margin-bottom: 20px; display: flex; gap: 14px; align-items: center; flex-wrap: wrap; }}
  .range-label {{ color: #a8d8ea; font-size: .95em; font-weight: bold; min-width: 200px; }}
  .range-pills {{ display: flex; gap: 6px; flex-wrap: wrap; }}
  .range-pill {{ background: #0f3460; color: #a8d8ea; border: 1px solid #1e2a4a; border-radius: 16px; padding: 5px 14px; font-size: .82em; cursor: pointer; transition: all .15s; }}
  .range-pill:hover {{ background: #1e2a4a; }}
  .range-pill.active {{ background: #f5a623; color: #1a1a2e; border-color: #f5a623; font-weight: bold; }}
  .range-custom {{ display: flex; gap: 6px; align-items: center; margin-left: auto; }}
  .range-custom input[type=date], .section select {{ background: #0f3460; color: #e0e0e0; border: 1px solid #1e2a4a; border-radius: 6px; padding: 4px 10px; font-size: .82em; color-scheme: dark; cursor: pointer; }}
  .section select:focus {{ outline: none; border-color: #f5a623; }}
  .section select option {{ background: #0f3460; color: #e0e0e0; }}
  .note {{ font-size: .78em; color: #555; margin-top: 10px; line-height: 1.6; }}
</style>
</head>
<body>
<h1>📊 Claude Code Token 使用统计</h1>
<div class="meta">扫描消息: {total_msgs:,} 条 &nbsp;|&nbsp; 筛选: {date_label} &nbsp;|&nbsp; 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</div>

<div class="cards">{cards}</div>

<div class="range-bar">
  <span class="range-label" id="rangeLabel"></span>
  <div class="range-pills">
    <button class="range-pill" data-preset="today" onclick="setPreset('today')">本日</button>
    <button class="range-pill active" data-preset="7d" onclick="setPreset('7d')">最近7天</button>
    <button class="range-pill" data-preset="month" onclick="setPreset('month')">本月</button>
    <button class="range-pill" data-preset="30d" onclick="setPreset('30d')">最近30天</button>
    <button class="range-pill" data-preset="3m" onclick="setPreset('3m')">最近3个月</button>
    <button class="range-pill" data-preset="1y" onclick="setPreset('1y')">最近1年</button>
  </div>
  <div class="range-custom">
    <input type="date" id="rangeStart" min="{_date_min}" max="{_date_max}" onchange="applyRange(this.value, document.getElementById('rangeEnd').value)">
    <span>~</span>
    <input type="date" id="rangeEnd" min="{_date_min}" max="{_date_max}" onchange="applyRange(document.getElementById('rangeStart').value, this.value)">
  </div>
</div>

<div class="chart-box">
  <h2>每日 Output / Input tokens（K）</h2>
  <canvas id="barChart" height="80"></canvas>
</div>
<div class="chart-box">
  <h2>每日缓存读取 tokens（K）</h2>
  <canvas id="crChart" height="60"></canvas>
</div>
<div class="chart-box">
  <h2>每日费用（USD）</h2>
  <canvas id="costChart" height="60"></canvas>
</div>

<div class="section">
  <h2 id="mdTitle">📅 {date_label} · 按模型消耗 Token</h2>
  <div style="margin-bottom:12px;display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
    <span style="color:#888;font-size:.85em;">区间用顶部选择器 · 排序</span>
    <select id="mdSort" onchange="applyRange(document.getElementById('rangeStart').value, document.getElementById('rangeEnd').value)">
      <option value="cost">费用 USD ↓</option>
      <option value="input">input ↓</option>
      <option value="output">output ↓</option>
      <option value="quota">配额消耗 ↓</option>
    </select>
  </div>
  <canvas id="mdChart" height="60"></canvas>
  <table class="data">
    <thead><tr><th>模型</th><th>价格 $/MTok (in/out)</th><th>input</th><th>output</th><th>缓存读</th><th>缓存写</th><th>配额消耗(全)</th><th>费用USD</th></tr></thead>
    <tbody id="mdTbody"></tbody>
  </table>
</div>

<div class="section">
  <h2>按自然月汇总</h2>
  <table class="data">
    <thead><tr>
      <th>月份</th><th>有效天</th><th>input</th><th>output</th><th>缓存读</th><th>缓存写</th>
      <th>实计(in+out)</th><th>配额消耗(全)</th><th>费用USD</th>
    </tr></thead>
    <tbody>{month_rows}</tbody>
  </table>
</div>

<div class="section">
  <h2>每日明细</h2>
  <table class="data">
    <thead><tr>
      <th>日期</th><th>input</th><th>output</th><th>缓存读</th><th>缓存写</th>
      <th>实计(in+out)</th><th>配额消耗(全)</th><th>费用USD</th>
    </tr></thead>
    <tbody id="detailBody"></tbody>
  </table>
</div>

<div class="section">
  <h2>按模型汇总（全历史，按 input 降序）</h2>
  <table class="data">
    <thead><tr><th>模型</th><th>input</th><th>output</th><th>缓存读</th><th>缓存写</th><th>配额消耗(全)</th><th>费用USD</th></tr></thead>
    <tbody>{model_rows}</tbody>
  </table>
</div>

<div class="section">
  <h2>按项目汇总（按费用降序）</h2>
  <table class="data">
    <thead><tr><th>项目</th><th>input</th><th>output</th><th>缓存读</th><th>缓存写</th><th>费用USD</th></tr></thead>
    <tbody>{proj_rows}</tbody>
  </table>
</div>

<div class="note">
  口径说明：<br>
  · 实计(in+out) = input + output，纯生成 token 量<br>
  · 配额消耗(全) = input + output + 缓存读 + 缓存写<br>
  · 费用 USD 按当天价格文件计算（model_prices.json），每天刷新一次<br>
  · 增减模型价格：编辑 claude_token_stats.py 中 _DEFAULT_PRICES，次日自动生效
</div>

<script>
{model_date_js}
</script>
</body>
</html>'''

out_path = os.path.join(os.path.expanduser('~'), 'Documents', 'claude_token_stats.html')
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(html)

print(f'  报告已生成: {out_path}')
webbrowser.open(f'file:///{out_path}')