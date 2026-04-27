"""
Claude Code 每日 Token 使用统计
扫描 ~/.claude/projects/ 下所有 .jsonl 会话文件，按日期汇总 token 用量

计价基准 (claude-opus-4-x, 美元/MTok):
  input        $3.00
  output       $15.00
  cache_read   $0.30
  cache_write  $3.75
"""
import os
import json
import sys
from collections import defaultdict
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')

BASE = os.path.expanduser(r'~/.claude/projects')

# 计价单位: 美元 / 百万 token (MTok)
PRICE = {
    'input':       3.00,
    'output':      15.00,
    'cache_read':  0.30,
    'cache_write': 3.75,
}

def calc_cost(inp, out, cr, cw):
    return (inp * PRICE['input'] + out * PRICE['output']
            + cr * PRICE['cache_read'] + cw * PRICE['cache_write']) / 1_000_000

# date -> {input, output, cache_read, cache_write}
daily = defaultdict(lambda: {'input': 0, 'output': 0, 'cache_read': 0, 'cache_write': 0})
proj_daily = defaultdict(lambda: defaultdict(lambda: {'input': 0, 'output': 0, 'cache_read': 0, 'cache_write': 0}))

total_msgs = 0
skipped = 0

for proj in sorted(os.listdir(BASE)):
    proj_path = os.path.join(BASE, proj)
    if not os.path.isdir(proj_path):
        continue
    for f in sorted(os.listdir(proj_path)):
        if not f.endswith('.jsonl'):
            continue
        fpath = os.path.join(proj_path, f)
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

                    # 提取时间戳
                    ts = obj.get('timestamp') or obj.get('ts') or obj.get('created_at')
                    if not ts:
                        # 尝试从 message 里取
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

                    # 提取 usage
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
                    daily[date_str]['input']      += inp
                    daily[date_str]['output']     += out
                    daily[date_str]['cache_read'] += cr
                    daily[date_str]['cache_write']+= cw
                    proj_daily[proj][date_str]['input']      += inp
                    proj_daily[proj][date_str]['output']     += out
                    proj_daily[proj][date_str]['cache_read'] += cr
                    proj_daily[proj][date_str]['cache_write']+= cw
        except Exception as e:
            print(f'  [warn] 读取失败 {fpath}: {e}')

# ── 输出 ──────────────────────────────────────────────────────────────────────
SEP = '─' * 108

print()
print('╔══════════════════════════════════════════════════════════════════════════════════════════════════════════╗')
print('║              Claude Code  每日 Token 使用统计                                                           ║')
print('║  计价: input $3/MTok  output $15/MTok  cache_read $0.30/MTok  cache_write $3.75/MTok                   ║')
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
for proj in sorted(proj_daily.keys()):
    pd = proj_daily[proj]
    pi = po = pcr = pcw = 0
    for d in pd.values():
        pi += d['input']; po += d['output']
        pcr += d['cache_read']; pcw += d['cache_write']
    if pi + po == 0:
        continue
    pcost = calc_cost(pi, po, pcr, pcw)
    short = proj[-40:] if len(proj) > 40 else proj
    print(f"  {short:<42}  input={pi:>10,}  output={po:>10,}  实计={pi+po:>10,}  费用=${pcost:.4f}")

print()
