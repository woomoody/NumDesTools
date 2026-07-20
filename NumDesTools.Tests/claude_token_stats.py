"""
Claude Code Token 使用统计 — 生成 HTML 报告并自动用浏览器打开
按实际使用的模型分别计价（$/MTok）
支持 --date 参数筛选日期范围
"""
import os, json, sys, subprocess, webbrowser, argparse
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

_DEFAULT_PRICES = [
    {'prefix': 'claude-fable',   'input': 10.00, 'output': 50.00, 'cache_read': 1.00,  'cache_write': 12.50},
    {'prefix': 'claude-mythos',  'input': 10.00, 'output': 50.00, 'cache_read': 1.00,  'cache_write': 12.50},
    {'prefix': 'claude-opus-4',  'input':  5.00, 'output': 25.00, 'cache_read': 0.50,  'cache_write':  6.25},
    {'prefix': 'claude-sonnet-4','input':  3.00, 'output': 15.00, 'cache_read': 0.30,  'cache_write':  3.75},
    {'prefix': 'claude-haiku-4', 'input':  1.00, 'output':  5.00, 'cache_read': 0.10,  'cache_write':  1.25},
]
_DEFAULT_FALLBACK = {'input': 5.00, 'output': 25.00, 'cache_read': 0.50, 'cache_write': 6.25}

def _load_prices():
    ts = today.isoformat()
    if os.path.exists(_PRICE_FILE):
        try:
            with open(_PRICE_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if data.get('date') == ts:
                return data['prices'], data['fallback']
        except: pass
    prices = _DEFAULT_PRICES[:]
    fallback = dict(_DEFAULT_FALLBACK)
    with open(_PRICE_FILE, 'w', encoding='utf-8') as f:
        json.dump({'date': ts, 'prices': prices, 'fallback': fallback}, f, ensure_ascii=False, indent=2)
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

def cn_num(n):
    if n >= 1_0000_0000: return f'{n/1_0000_0000:.2f}亿'
    if n >= 1_0000:      return f'{n/1_0000:.1f}万'
    return f'{n:,}'

BASES = [(os.path.expanduser(r'~/.claude/projects'), '[local]')]

# ── 数据采集 ──────────────────────────────────────────────────────────────────
_zero = lambda: {'input':0,'output':0,'cache_read':0,'cache_write':0,'cost':0.0}
daily       = defaultdict(_zero)
monthly     = defaultdict(_zero)
proj_daily  = defaultdict(lambda: defaultdict(_zero))
model_daily = defaultdict(lambda: defaultdict(_zero))  # model_daily[date][model]
total_msgs = skipped = 0

for BASE, prefix in BASES:
    if not os.path.isdir(BASE): continue
    for proj in sorted(os.listdir(BASE)):
        proj_path = os.path.join(BASE, proj)
        if not os.path.isdir(proj_path): continue
        proj_key = f"{prefix}{proj}"
        for dirpath, _, files in os.walk(proj_path):
            if os.sep + 'subagents' in dirpath or os.sep + 'workflows' in dirpath:
                continue
            for f in sorted(files):
                if not f.endswith('.jsonl'): continue
                fpath = os.path.join(dirpath, f)
                try:
                    with open(fpath, 'r', encoding='utf-8') as fp:
                        for line in fp:
                            line = line.strip()
                            if not line: continue
                            try: obj = json.loads(line)
                            except: skipped += 1; continue
                            ts = obj.get('timestamp') or obj.get('ts') or obj.get('created_at')
                            if not ts:
                                msg = obj.get('message', {})
                                if isinstance(msg, dict): ts = msg.get('timestamp') or msg.get('created_at')
                            if not ts: continue
                            try:
                                date_str = datetime.fromtimestamp(ts).strftime('%Y-%m-%d') if isinstance(ts, (int, float)) else str(ts)[:10]
                            except: continue
                            usage = obj.get('usage') or (isinstance(obj.get('message'), dict) and obj['message'].get('usage'))
                            if not isinstance(usage, dict): continue
                            inp = usage.get('input_tokens',0) or 0
                            out = usage.get('output_tokens',0) or 0
                            cr  = usage.get('cache_read_input_tokens',0) or 0
                            cw  = usage.get('cache_creation_input_tokens',0) or 0
                            if inp+out+cr+cw == 0: continue
                            if out < 200 and cr == 0 and cw == 0:
                                skipped += 1; continue
                            model = (obj.get('model') or
                                     (obj.get('message') or {}).get('model') or '<empty>')
                            cost  = calc_cost(inp, out, cr, cw, model)
                            total_msgs += 1
                            for d in (daily[date_str], monthly[date_str[:7]],
                                      proj_daily[proj_key][date_str],
                                      model_daily[date_str][model]):
                                d['input']      += inp
                                d['output']     += out
                                d['cache_read'] += cr
                                d['cache_write']+= cw
                                d['cost']       += cost
                except Exception as e:
                    print(f'  [warn] {fpath}: {e}')

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
  .note {{ font-size: .78em; color: #555; margin-top: 10px; line-height: 1.6; }}
</style>
</head>
<body>
<h1>📊 Claude Code Token 使用统计</h1>
<div class="meta">扫描消息: {total_msgs:,} 条 &nbsp;|&nbsp; 筛选: {date_label} &nbsp;|&nbsp; 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</div>

<div class="cards">{cards}</div>

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
  <h2>📅 {date_label} · 按模型消耗 Token</h2>
  <table class="data">
    <thead><tr><th>模型</th><th>input</th><th>output</th><th>缓存读</th><th>缓存写</th><th>配额消耗(全)</th><th>费用USD</th></tr></thead>
    <tbody>{model_date_rows if model_date_rows else '<tr><td colspan="7" style="text-align:center;color:#666;">该时间段无数据</td></tr>'}</tbody>
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
    <tbody>{detail_rows}</tbody>
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
const labels = {labels_js};
const outputData = {output_js};
const inputData  = {input_js};
const crData     = {cr_js};
const costData   = {cost_js};

const gridColor = 'rgba(255,255,255,0.06)';
const tickColor = '#666';

new Chart(document.getElementById('barChart'), {{
  type: 'bar',
  data: {{
    labels,
    datasets: [
      {{ label: 'output (K)', data: outputData, backgroundColor: 'rgba(168,216,234,0.75)', order: 1 }},
      {{ label: 'input (K)',  data: inputData,  backgroundColor: 'rgba(100,149,237,0.55)', order: 2 }},
    ]
  }},
  options: {{
    responsive: true,
    plugins: {{ legend: {{ labels: {{ color: '#aaa' }} }} }},
    scales: {{
      x: {{ ticks: {{ color: tickColor, maxRotation: 45 }}, grid: {{ color: gridColor }} }},
      y: {{ ticks: {{ color: tickColor }}, grid: {{ color: gridColor }} }}
    }}
  }}
}});

new Chart(document.getElementById('crChart'), {{
  type: 'bar',
  data: {{ labels, datasets: [{{ label: '缓存读 (K)', data: crData, backgroundColor: 'rgba(245,166,35,0.6)' }}] }},
  options: {{
    responsive: true,
    plugins: {{ legend: {{ labels: {{ color: '#aaa' }} }} }},
    scales: {{
      x: {{ ticks: {{ color: tickColor, maxRotation: 45 }}, grid: {{ color: gridColor }} }},
      y: {{ ticks: {{ color: tickColor }}, grid: {{ color: gridColor }} }}
    }}
  }}
}});

new Chart(document.getElementById('costChart'), {{
  type: 'line',
  data: {{ labels, datasets: [{{ label: '费用 USD', data: costData,
    borderColor: '#f5a623', backgroundColor: 'rgba(245,166,35,0.15)',
    pointRadius: 3, tension: 0.3, fill: true }}] }},
  options: {{
    responsive: true,
    plugins: {{ legend: {{ labels: {{ color: '#aaa' }} }} }},
    scales: {{
      x: {{ ticks: {{ color: tickColor, maxRotation: 45 }}, grid: {{ color: gridColor }} }},
      y: {{ ticks: {{ color: tickColor, callback: v => '$'+v }}, grid: {{ color: gridColor }} }}
    }}
  }}
}});
</script>
</body>
</html>'''

out_path = os.path.join(os.path.expanduser('~'), 'Documents', 'claude_token_stats.html')
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(html)

print(f'  报告已生成: {out_path}')
webbrowser.open(f'file:///{out_path}')