"""
Claude Code 每日 Token 使用统计
扫描 ~/.claude/projects/ 下所有 .jsonl 会话文件，按日期汇总 token 用量

计价基准 (claude-opus-4-x, 美元/MTok):
  input        $5.00
  output       $25.00
  cache_read   $0.50
  cache_write  $6.25
"""
import os
import json
import sys
import subprocess
import tempfile
from collections import defaultdict
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')

REMOTE_SSH  = "admin@100.96.48.30"
REMOTE_PATH = r"C:\Users\admin\.claude\projects"

# 计价单位: 美元 / 百万 token (MTok)
PRICE = {
    'input':       5.00,
    'output':      25.00,
    'cache_read':  0.50,
    'cache_write': 6.25,
}

def calc_cost(inp, out, cr, cw):
    return (inp * PRICE['input'] + out * PRICE['output']
            + cr * PRICE['cache_read'] + cw * PRICE['cache_write']) / 1_000_000

BASES = [
    (os.path.expanduser(r'~/.claude/projects'), '[local]'),
]

def _pull_remote_jsonl():
    """通过 SSH 读取远程 jsonl，yield (proj_key, date_str, inp, out, cr, cw)"""
    r = subprocess.run(
        ["ssh", "-o", "StrictHostKeyChecking=no", "-o", "BatchMode=yes",
         REMOTE_SSH, f'dir "{REMOTE_PATH}" /s /b'],
        capture_output=True, text=True, timeout=30
    )
    if r.returncode != 0:
        print(f"  [warn] 远程 SSH 不可用，跳过远程数据", file=sys.stderr)
        return
    for fpath in r.stdout.splitlines():
        fpath = fpath.strip()
        if not fpath.endswith('.jsonl'):
            continue
        rel = fpath[len(REMOTE_PATH):].lstrip('\\')
        proj_key = "[remote]" + rel.split('\\')[0]
        r2 = subprocess.run(
            ["ssh", "-o", "StrictHostKeyChecking=no", "-o", "BatchMode=yes",
             REMOTE_SSH, f'type "{fpath}"'],
            capture_output=True, text=True, encoding='utf-8', errors='ignore', timeout=30
        )
        for line in r2.stdout.splitlines():
            line = line.strip()
            if not line:
                continue
            try:
                obj = json.loads(line)
            except Exception:
                continue
            ts = obj.get('timestamp') or obj.get('ts') or obj.get('created_at')
            if not ts:
                msg = obj.get('message', {})
                if isinstance(msg, dict):
                    ts = msg.get('timestamp') or msg.get('created_at')
            if not ts:
                continue
            try:
                date_str = datetime.fromtimestamp(ts).strftime('%Y-%m-%d') if isinstance(ts, (int, float)) else str(ts)[:10]
            except Exception:
                continue
            usage = obj.get('usage') or (obj.get('message') or {}).get('usage')
            if not isinstance(usage, dict):
                continue
            inp = usage.get('input_tokens', 0) or 0
            out = usage.get('output_tokens', 0) or 0
            cr  = usage.get('cache_read_input_tokens', 0) or 0
            cw  = usage.get('cache_creation_input_tokens', 0) or 0
            if inp + out + cr + cw == 0:
                continue
            yield proj_key, date_str, inp, out, cr, cw

# date -> {input, output, cache_read, cache_write}
daily = defaultdict(lambda: {'input': 0, 'output': 0, 'cache_read': 0, 'cache_write': 0})
proj_daily = defaultdict(lambda: defaultdict(lambda: {'input': 0, 'output': 0, 'cache_read': 0, 'cache_write': 0}))

total_msgs = 0
skipped = 0

for BASE, prefix in BASES:
    if not os.path.isdir(BASE):
        continue
    for proj in sorted(os.listdir(BASE)):
        proj_path = os.path.join(BASE, proj)
        if not os.path.isdir(proj_path):
            continue
        proj_key = f"{prefix}{proj}"
        for dirpath, _, files in os.walk(proj_path):
          for f in sorted(files):
            if not f.endswith('.jsonl'):
                continue
            fpath = os.path.join(dirpath, f)
            try:
                with open(fpath, 'r', encoding='utf-8') as fp:
                    for line in fp:
                        line = line.strip()
                        if not line:
                            continue
                        try:
                            obj = json.loads(line)
                        except json.JSONDecodeError:
                            skipped += 1
                            continue

                        ts = obj.get('timestamp') or obj.get('ts') or obj.get('created_at')
                        if not ts:
                            msg = obj.get('message', {})
                            if isinstance(msg, dict):
                                ts = msg.get('timestamp') or msg.get('created_at')
                        if not ts:
                            continue

                        try:
                            if isinstance(ts, (int, float)):
                                date_str = datetime.fromtimestamp(ts).strftime('%Y-%m-%d')
                            else:
                                date_str = str(ts)[:10]
                        except Exception:
                            continue

                        usage = None
                        if 'usage' in obj:
                            usage = obj['usage']
                        elif isinstance(obj.get('message'), dict):
                            usage = obj['message'].get('usage')

                        if not isinstance(usage, dict):
                            continue

                        inp = usage.get('input_tokens', 0) or 0
                        out = usage.get('output_tokens', 0) or 0
                        cr  = usage.get('cache_read_input_tokens', 0) or 0
                        cw  = usage.get('cache_creation_input_tokens', 0) or 0

                        if inp + out + cr + cw == 0:
                            continue

                        total_msgs += 1
                        daily[date_str]['input']       += inp
                        daily[date_str]['output']      += out
                        daily[date_str]['cache_read']  += cr
                        daily[date_str]['cache_write'] += cw
                        proj_daily[proj_key][date_str]['input']       += inp
                        proj_daily[proj_key][date_str]['output']      += out
                        proj_daily[proj_key][date_str]['cache_read']  += cr
                        proj_daily[proj_key][date_str]['cache_write'] += cw
            except Exception as e:
                print(f'  [warn] 读取失败 {fpath}: {e}')

# ── 远程数据 ──────────────────────────────────────────────────────────────────
print("  正在读取远程数据...", flush=True)
for proj_key, date_str, inp, out, cr, cw in _pull_remote_jsonl():
    total_msgs += 1
    daily[date_str]['input']       += inp
    daily[date_str]['output']      += out
    daily[date_str]['cache_read']  += cr
    daily[date_str]['cache_write'] += cw
    proj_daily[proj_key][date_str]['input']       += inp
    proj_daily[proj_key][date_str]['output']      += out
    proj_daily[proj_key][date_str]['cache_read']  += cr
    proj_daily[proj_key][date_str]['cache_write'] += cw

# ── 输出 ──────────────────────────────────────────────────────────────────────
SEP = '─' * 108

print()
print('╔══════════════════════════════════════════════════════════════════════════════════════════════════════════╗')
print('║              Claude Code  每日 Token 使用统计                                                           ║')
print('║  计价: input $5/MTok  output $25/MTok  cache_read $0.50/MTok  cache_write $6.25/MTok                   ║')
print('╚══════════════════════════════════════════════════════════════════════════════════════════════════════════╝')
print(f'  扫描消息条数: {total_msgs:,}  |  跳过损坏行: {skipped}')
print()

print(SEP)
print(f"{'日期':<12} {'input':>12} {'output':>12} {'cache_read':>16} {'cache_write':>14}  {'实计(in+out)':>14}  {'费用(USD $)':>12}")
print(SEP)

grand_in = grand_out = grand_cr = grand_cw = 0
for date in sorted(daily.keys()):
    d = daily[date]
    i, o, cr, cw = d['input'], d['output'], d['cache_read'], d['cache_write']
    grand_in += i; grand_out += o; grand_cr += cr; grand_cw += cw
    cost = calc_cost(i, o, cr, cw)
    print(f"{date:<12} {i:>12,} {o:>12,} {cr:>16,} {cw:>14,}  {i+o:>14,}  ${cost:>10.4f}")

print(SEP)
grand_cost = calc_cost(grand_in, grand_out, grand_cr, grand_cw)
print(f"{'合计':<12} {grand_in:>12,} {grand_out:>12,} {grand_cr:>16,} {grand_cw:>14,}  {grand_in+grand_out:>14,}  ${grand_cost:>10.4f}")
print(SEP)

print()
print('── 按项目汇总 ──')
print(f"  {'项目':<44} {'input':>10}  {'output':>10}  {'cache_r':>12}  {'cache_w':>12}  {'费用(USD)':>10}")
print(f"  {'─'*44} {'─'*10}  {'─'*10}  {'─'*12}  {'─'*12}  {'─'*10}")
for proj in sorted(proj_daily.keys(), key=lambda p: -sum(calc_cost(d['input'],d['output'],d['cache_read'],d['cache_write']) for d in proj_daily[p].values())):
    pd = proj_daily[proj]
    pi = po = pcr = pcw = 0
    for d in pd.values():
        pi += d['input']; po += d['output']
        pcr += d['cache_read']; pcw += d['cache_write']
    if pi + po + pcr + pcw == 0:
        continue
    pcost = calc_cost(pi, po, pcr, pcw)
    short = proj[-44:] if len(proj) > 44 else proj
    print(f"  {short:<44} {pi:>10,}  {po:>10,}  {pcr:>12,}  {pcw:>12,}  ${pcost:>9.4f}")

print()

# ── 字符图表 ──────────────────────────────────────────────────────────────────
def ascii_chart(daily, calc_cost):
    raw_dates = sorted(daily.keys())
    if not raw_dates:
        return

    # 填充日期轴：从最早到最晚，缺失日期补 0
    from datetime import date, timedelta
    d0 = date.fromisoformat(raw_dates[0])
    d1 = date.fromisoformat(raw_dates[-1])
    dates = [(d0 + timedelta(days=i)).isoformat() for i in range((d1 - d0).days + 1)]
    empty = {'input': 0, 'output': 0, 'cache_read': 0, 'cache_write': 0}

    costs   = [calc_cost(daily.get(d, empty)['input'], daily.get(d, empty)['output'],
                         daily.get(d, empty)['cache_read'], daily.get(d, empty)['cache_write']) for d in dates]
    outputs = [daily.get(d, empty)['output'] / 1000 for d in dates]   # K tokens
    inputs  = [daily.get(d, empty)['input']  / 1000 for d in dates]

    ROWS   = 10   # 图高（行数）
    BAR_W  = 6    # 每天占宽
    YLABEL = 8    # 左侧标签宽

    def make_bar_chart(values, title, unit, bar_char='█', sub_char='▄'):
        max_v = max(values) if max(values) > 0 else 1
        print(f'\n  {title}')
        for row in range(ROWS, 0, -1):
            threshold = max_v * row / ROWS
            line = f'  {max_v * row / ROWS:{YLABEL-2}.0f}{unit} │'
            for v in values:
                filled = v >= threshold
                half   = v >= max_v * (row - 0.5) / ROWS and not filled
                if filled:
                    line += f' {bar_char*4} '
                elif half:
                    line += f' {sub_char*4} '
                else:
                    line += ' ' * BAR_W
            print(line)
        # x 轴
        print(' ' * YLABEL + '  └' + '──────' * len(dates))
        # 日期标签
        label_line = ' ' * (YLABEL + 3)
        for d in dates:
            label_line += f'{d[5:]:^6}'
        print(label_line)

    def make_cost_chart(values, title):
        max_v = max(values) if max(values) > 0 else 1
        # 折线图：用字符模拟
        rows_data = []
        for row in range(ROWS, 0, -1):
            threshold_hi = max_v * row / ROWS
            threshold_lo = max_v * (row - 1) / ROWS
            row_chars = []
            for v in values:
                if threshold_lo < v <= threshold_hi:
                    row_chars.append('●')
                elif v > threshold_hi:
                    row_chars.append('│')
                else:
                    row_chars.append(' ')
            rows_data.append((threshold_hi, row_chars))

        print(f'\n  {title}')
        for threshold, row_chars in rows_data:
            line = f'  {threshold:{YLABEL-2}.2f}$ │'
            for i, ch in enumerate(row_chars):
                # 连接相邻点
                if ch == '●' and i + 1 < len(row_chars) and row_chars[i+1] == '●':
                    line += f'  {ch}───'
                elif ch == '●':
                    line += f'  {ch}   '
                elif ch == '│':
                    line += f'  {ch}   '
                else:
                    # 检查是否需要画水平连接线
                    line += ' ' * BAR_W
            print(line)
        print(' ' * YLABEL + '  └' + '──────' * len(dates))
        label_line = ' ' * (YLABEL + 3)
        for d in dates:
            label_line += f'{d[5:]:^6}'
        print(label_line)
        # 数值行
        val_line = ' ' * (YLABEL + 3)
        for c in values:
            val_line += f'{"$"+f"{c:.1f}":^6}'
        print(val_line)

    SEP2 = '─' * (YLABEL + 4 + BAR_W * len(dates))
    print()
    print(SEP2)
    print('  📊 Token 使用趋势（字符图）')
    print(SEP2)
    make_bar_chart(outputs, '■ output tokens (K)', 'K', '█')
    make_bar_chart(inputs,  '■ input tokens (K)',  'K', '░')
    make_cost_chart(costs,  '● 每日费用 (USD)')
    print()

ascii_chart(daily, calc_cost)
