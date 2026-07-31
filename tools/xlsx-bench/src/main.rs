//! 方案A（纯Rust）性能实测：用真实 Item.xlsx（~6.5万行×85列）测「读入 / 排序 / 整列复制」
//! 三段耗时，跟 C# 方案C（TuiMigrationBenchmarkC.cs）同口径对比。只读，不写回。

use calamine::{open_workbook, Data, Reader, Xlsx};
use std::time::Instant;

const ITEM_PATH: &str = r"C:\M1Work\public\Excels\Tables\Item.xlsx";
const SORT_COL: usize = 1; // 对齐 C# 方案 SortBy(1, ...)（B 列/id 列）

fn median(mut v: Vec<f64>) -> f64 {
    v.sort_by(|a, b| a.partial_cmp(b).unwrap());
    v[v.len() / 2]
}

/// 列式存储：cols[col][row]，对齐 C# ColumnStore 的列式布局。
fn load() -> (Vec<Vec<String>>, usize, usize) {
    let mut wb: Xlsx<_> = open_workbook(ITEM_PATH).expect("open Item.xlsx");
    let sheet = wb.sheet_names()[0].clone();
    let range = wb.worksheet_range(&sheet).expect("read worksheet range");
    let row_count = range.rows().count();
    let col_count = range.width();

    let mut cols: Vec<Vec<String>> = (0..col_count).map(|_| Vec::with_capacity(row_count)).collect();
    for row in range.rows() {
        for (c, cell) in row.iter().enumerate() {
            cols[c].push(match cell {
                Data::Empty => String::new(),
                Data::String(s) => s.clone(),
                Data::Float(f) => f.to_string(),
                Data::Int(i) => i.to_string(),
                Data::Bool(b) => b.to_string(),
                other => other.to_string(),
            });
        }
    }
    (cols, row_count, col_count)
}

/// 只排 usize 索引数组，不拷贝整表——对齐 C# VirtualizingSortableView.SortBy 的设计。
fn sort_index(cols: &[Vec<String>], by_col: usize, row_count: usize) -> Vec<usize> {
    let mut order: Vec<usize> = (0..row_count).collect();
    order.sort_by(|&a, &b| cols[by_col][a].as_bytes().cmp(cols[by_col][b].as_bytes()));
    order
}

/// 复刻 C# MainWindow.CopySelectionToClipboard 的 EntireColumn 分支：按排序后的行顺序，
/// 把全部列拼成 Tab 分隔、换行分隔的大字符串（模拟拷进剪贴板前的文本构建）。
fn copy_all_columns(cols: &[Vec<String>], order: &[usize], col_count: usize) -> usize {
    let mut out = String::new();
    for &r in order {
        let parts: Vec<&str> = (0..col_count).map(|c| cols[c][r].as_str()).collect();
        out.push_str(&parts.join("\t"));
        out.push('\n');
    }
    out.len()
}

fn main() {
    let (cols, row_count, col_count) = load(); // warm-up，顺带拿到维度

    let mut load_times = Vec::with_capacity(3);
    for _ in 0..3 {
        let t = Instant::now();
        let _ = load();
        load_times.push(t.elapsed().as_secs_f64() * 1000.0);
    }

    let mut sort_times = Vec::with_capacity(3);
    let mut order = Vec::new();
    for _ in 0..3 {
        let t = Instant::now();
        order = sort_index(&cols, SORT_COL, row_count);
        sort_times.push(t.elapsed().as_secs_f64() * 1000.0);
    }

    let mut copy_times = Vec::with_capacity(3);
    for _ in 0..3 {
        let t = Instant::now();
        let _ = copy_all_columns(&cols, &order, col_count);
        copy_times.push(t.elapsed().as_secs_f64() * 1000.0);
    }

    println!(
        "[方案A 纯Rust] rows={} cols={} | 读入(median of 3)={:.0}ms 排序={:.0}ms 整列复制(全{}列)={:.0}ms",
        row_count,
        col_count,
        median(load_times),
        median(sort_times),
        col_count,
        median(copy_times)
    );
}
