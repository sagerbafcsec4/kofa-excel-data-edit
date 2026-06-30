#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
サッカー記録ワークブック バッチ整形スクリプト (excel-data skill 準拠)

★このファイルがロジックの唯一の正本。Webアプリ(index.html)は実行時にこれを読み込むため、
  ロジックの修正はこのファイルだけ直せばよい(index.html 側に書き写す必要はない)。
"""
import sys, os, glob
from datetime import datetime
import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy
from collections import Counter

try:
    BASE = os.path.dirname(os.path.abspath(__file__))
except NameError:
    BASE = os.getcwd()
OUT  = os.path.join(BASE, "_processed")
LOGS = os.path.join(BASE, "_logs")
DELETE_KEIREKI = None
MATCH_DATE = None       # 試合日 "YYYY/MM/DD"。None(未指定)なら歳列は変更しない

def is_empty(v):
    return v is None or (isinstance(v, str) and v.strip() == "")

def jwidth(text):
    w = 0
    for ch in str(text):
        w += 2 if ord(ch) > 0x2E7F else 1
    return w

def set_font_size(cell, size):
    f = copy(cell.font); f.size = size; cell.font = f

def set_shrink(cell):
    a = copy(cell.alignment); a.shrink_to_fit = True; cell.alignment = a

def read_grid(ws):
    return [list(row) for row in ws.iter_rows(values_only=True)]

def g(grid, r, c):
    if 1 <= r <= len(grid) and 1 <= c <= len(grid[r-1]):
        return grid[r-1][c-1]
    return None

def last_data_row(grid, c1, c2):
    last = 1
    for r in range(1, len(grid) + 1):
        row = grid[r-1]
        for c in range(c1, min(c2, len(row)) + 1):
            if not is_empty(row[c-1]):
                last = r; break
    return last

def col_empty(grid, c, r1, r2):
    for r in range(r1, r2 + 1):
        if not is_empty(g(grid, r, c)):
            return False
    return True

WIDTHS = {"節":10,"開催日":10,"H":3,"A":3,"ｽｺｱ":8,"スコア":8,"対戦相手":30,"退場":20,"NEWS":60}

def _is_note(v):
    """表外の注記セルか(縮小・書式変更してはいけない)。"""
    return isinstance(v, str) and ("用語説明" in v or v.lstrip().startswith("【"))

def fmt_season(ws, log):
    grid = read_grid(ws)
    last = last_data_row(grid, 1, 18)
    for cd in ws.column_dimensions.values():
        cd.hidden = False
    score_cols = [c for c in range(1, 19)
                  if isinstance(g(grid, 2, c), str) and g(grid, 2, c).strip().startswith("得点")]
    del_score = []
    if score_cols:
        start = score_cols[0]; block = list(range(start, start + 8))
        del_score = [block[i] for i in range(len(block) - 1, 2, -1) if col_empty(grid, block[i], 3, last)]
    del_empty = [c for c in range(1, 19) if is_empty(g(grid, 2, c)) and col_empty(grid, c, 2, last)]
    for c in sorted(set(del_score) | set(del_empty), reverse=True):
        lab = "得点空" if c in del_score else "空ヘッダー"
        log(f"  {get_column_letter(c)}({lab}) -> 削除")
        ws.delete_cols(c, 1)
    grid = read_grid(ws)
    fmax = max((c for c in range(1, 26) if not is_empty(g(grid, 2, c))), default=18)
    # フォント11pt + 縮小は「表の枠内」だけ。表外の注記(【用語説明】等)は触らない。
    tbl_last = max((r for r in range(3, last + 1) if not is_empty(g(grid, r, 1))), default=2)
    for row in ws.iter_rows(min_row=2, max_row=tbl_last, min_col=1, max_col=fmax):
        for cell in row:
            if _is_note(cell.value):
                continue
            set_font_size(cell, 11); set_shrink(cell)
    # 表外の注記(用語説明等)に残っている「縮小して全体を表示」を能動的に解除する
    for row in ws.iter_rows():
        for cell in row:
            if _is_note(cell.value) and cell.alignment.shrink_to_fit:
                a = copy(cell.alignment); a.shrink_to_fit = False; cell.alignment = a
                log(f"  注記 {cell.coordinate} の縮小を解除")
    for c in range(1, fmax + 1):
        h = g(grid, 2, c)
        h = h.strip() if isinstance(h, str) else h
        if h in WIDTHS:
            w = WIDTHS[h]
        elif isinstance(h, str) and (h.startswith("得点") or h.isdigit()):
            w = 20
        else:
            continue
        ws.column_dimensions[get_column_letter(c)].width = w
        log(f"  幅 {get_column_letter(c)}('{h}') -> {w}")

PLAY_POS = {"GK","DF","MF","FW"}

def fmt_appearance(ws, log):
    grid = read_grid(ws)
    H = 8
    maxc = max((len(r) for r in grid), default=0)
    last = last_data_row(grid, 1, maxc)
    w = max((jwidth(g(grid, r, 3)) for r in range(1, last+1) if not is_empty(g(grid, r, 3))), default=0)
    if w:
        width = min(round(w * 1.2) + 4, 80)
        ws.column_dimensions["C"].width = width
        log(f"  名前列 C autofit -> {width}")
    def match_empty_2row(r1, r2):
        for r in (r1, r2):
            for c in range(H, maxc + 1):
                if not is_empty(g(grid, r, c)):
                    return False
        return True
    del_players = []
    r = 6
    while r <= last:
        pos = g(grid, r, 1)
        pos_s = pos.strip() if isinstance(pos, str) else pos
        if pos_s in PLAY_POS:
            if is_empty(g(grid, r, 2)) and match_empty_2row(r, r+1):
                del_players.append((r, g(grid, r, 3)))
            r += 2
        else:
            r += 1
    for r, name in sorted(del_players, reverse=True):
        log(f"  記録ゼロ選手 行{r}-{r+1} '{name}' -> 2行削除")
        ws.delete_rows(r, 2)
    grid2 = read_grid(ws)
    maxc2 = max((len(r) for r in grid2), default=0)
    last2 = last_data_row(grid2, 1, maxc2)
    def match_unplayed(c):
        if not is_empty(g(grid2, 4, c)):
            return False
        for cc in (c, c+1):
            for r in range(6, last2 + 1):
                if not is_empty(g(grid2, r, cc)):
                    return False
        return True
    leftmost = None
    c = H
    while c <= maxc2:
        if match_unplayed(c):
            leftmost = c; break
        c += 2
    if leftmost is not None:
        cut = leftmost + 2
        if cut <= maxc2:
            n = maxc2 - cut + 1
            log(f"  未実施試合 {get_column_letter(cut)}:{get_column_letter(maxc2)} ({n}列) -> 削除")
            ws.delete_cols(cut, n)
    # 凡例(◎：フル出場…)の結合は指示外なので解除する(最終レイアウトで判定)
    g3 = read_grid(ws)
    legend_row = next((r for r in range(1, len(g3) + 1)
                       if isinstance(g(g3, r, 1), str) and "フル出場" in g(g3, r, 1)), None)
    if legend_row is not None:
        removed = 0
        for m in list(ws.merged_cells.ranges):
            if m.min_row == legend_row:
                ws.unmerge_cells(str(m)); removed += 1
        if removed:
            log(f"  凡例(行{legend_row})の結合を {removed}件 解除")

PANELS = [(2,25),(26,49),(50,73),(74,97)]
RIGHT_COL = {1:"Y", 2:"AW", 3:"BU", 4:"CS"}
FILLED_MIN = 40

def fmt_formation(ws, wb, log):
    grid = read_grid(ws)
    n = 0
    for r1, r2 in ((1,49),(50,97)):
        for c1, c2 in PANELS:
            cnt = 0
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    if not is_empty(g(grid, r, c)):
                        cnt += 1
                if cnt >= FILLED_MIN: break
            if cnt >= FILLED_MIN: n += 1
    log(f"  フォーメーション数 = {n}")
    if n == 0:
        log(f"  -> 空シート '{ws.title}' 削除"); wb.remove(ws); return
    if n <= 4:
        # 上段の下端を動的に決定。テンプレの「負傷」行が上段/下段に1つずつあるので、
        # その間隔=バンド高さ。上段の下端=1+バンド高さ → スタッツ(枠/CK)行まで含まれる。
        injuries = [r for r in range(1, len(grid) + 1)
                    if any(isinstance(g(grid, r, c), str) and g(grid, r, c).strip() == "負傷"
                           for c in (2, 26, 50, 74))]
        bottom = (1 + (injuries[1] - injuries[0])) if len(injuries) >= 2 else 49
        right = RIGHT_COL[n]
        ws.print_area = f"A1:{right}{bottom}"
        log(f"  -> 印刷範囲 A1:{right}{bottom}")
    else:
        log(f"  -> 5試合以上、印刷範囲は変更なし")

def fmt_hyoki(ws, log):
    grid = read_grid(ws)
    maxc = max((len(r) for r in grid), default=0)
    last = last_data_row(grid, 1, maxc)
    age_col = bd_col = None
    for c in range(1, maxc + 1):
        h = g(grid, 2, c)
        h = h.strip() if isinstance(h, str) else h
        if h == "歳": age_col = c
        elif h == "生年月日": bd_col = c
    if MATCH_DATE and age_col and bd_col:
        bdL = get_column_letter(bd_col)
        date_str = MATCH_DATE.replace("-", "/")
        cnt_age = 0
        for r in range(3, last + 1):
            if not is_empty(g(grid, r, bd_col)):
                ws.cell(r, age_col).value = f'=DATEDIF({bdL}{r}, "{date_str}", "Y")'
                cnt_age += 1
        log(f"  歳列({get_column_letter(age_col)}) に {date_str} 基準の年齢式を {cnt_age}件 設定")
        ws.cell(last + 1, age_col).value = "※年齢は試合当日の年齢"
        log(f"  注記 {get_column_letter(age_col)}{last + 1} に『※年齢は試合当日の年齢』を設定")
    elif age_col:
        log("  歳列: 試合日未指定のため変更なし")
    # 列幅(表記のみ): 狭めるの禁止。広げるのは文字が切れる(内容>現在幅)時だけ。タイトル(行1)は除外。
    DEFAULT_W = 8.43
    for c in range(1, maxc + 1):
        need = 0
        for r in range(2, last + 1):
            v = g(grid, r, c)
            if not is_empty(v):
                need = max(need, jwidth(v))
        if need == 0:
            continue
        cur = ws.column_dimensions[get_column_letter(c)].width or DEFAULT_W
        # 明確に溢れている時だけ広げる(僅差=表示上切れていない範囲は触らない)。狭めるのは禁止。
        if need > cur + 1:
            ws.column_dimensions[get_column_letter(c)].width = round(need + 0.5, 2)
            log(f"  幅 {get_column_letter(c)} {round(cur,1)}->{round(need+0.5,2)} (切れ防止で拡張)")

def _row_set(ws, r, cmax=14):
    s = set()
    for c in range(1, cmax + 1):
        v = ws.cell(r, c).value
        if v not in (None, "") and str(v).strip() != "":
            s.add(str(v).strip())
    return s

def _cell(ws, r, c):
    v = ws.cell(r, c).value
    return str(v).strip() if v not in (None, "") else ""

def is_appearance_sheet(ws):
    if _cell(ws, 1, 4) == "出場記録":
        return True
    return _cell(ws, 4, 1) == "Pos" and _cell(ws, 4, 2) == "No" and _cell(ws, 4, 3) == "Name"

def is_season_sheet(ws):
    h = _row_set(ws, 2)
    return "節" in h and "開催日" in h and "対戦相手" in h

def is_hyoki_sheet(ws):
    return "統一表記" in _row_set(ws, 2)

def is_keireki_sheet(ws):
    h = _row_set(ws, 2)
    return "主な下部組織" in h and "経歴" in h

def is_formation_sheet(ws):
    return 90 <= ws.max_column <= 110

def process_wb(wb, log):
    for ws in list(wb.worksheets):
        t = ws.title
        if is_appearance_sheet(ws):
            log(f"[§A 出場記録] {t}"); fmt_appearance(ws, log)
        elif is_season_sheet(ws):
            log(f"[§B シーズン] {t}"); fmt_season(ws, log)
        elif is_hyoki_sheet(ws):
            log(f"[§D 表記] {t}"); fmt_hyoki(ws, log)
        elif is_keireki_sheet(ws):
            if DELETE_KEIREKI:
                log(f"[§E 経歴] {t} -> 削除"); wb.remove(ws)
            else:
                log(f"[§E 経歴] {t} -> 保持")
        elif is_formation_sheet(ws):
            log(f"[§C フォーメーション] {t}"); fmt_formation(ws, wb, log)
        else:
            log(f"[--] {t} -> 対象外シート、スキップ")
    return wb

def _exempt_cols(ws):
    if MATCH_DATE and is_hyoki_sheet(ws):
        for c in range(1, (ws.max_column or 0) + 1):
            v = ws.cell(2, c).value
            if isinstance(v, str) and v.strip() == "歳":
                return {c}
    return set()

def _sheet_values(ws):
    exempt = _exempt_cols(ws)
    c = Counter()
    for row in ws.iter_rows(values_only=True):
        for idx, v in enumerate(row, start=1):
            if idx in exempt:
                continue
            if v is not None and not (isinstance(v, str) and v.strip() == ""):
                c[str(v)] += 1
    return c

def snapshot_values(wb):
    return {ws.title: _sheet_values(ws) for ws in wb.worksheets}

def verify_no_value_change(before, wb):
    problems = []
    for ws in wb.worksheets:
        base = before.get(ws.title)
        if base is None:
            problems.append(f"シート『{ws.title}』が元ファイルに無い(追加されている)")
            continue
        extra = _sheet_values(ws) - base
        if extra:
            problems.append(f"シート『{ws.title}』で値の変更/追加を検知: {list(extra.items())[:5]}")
    return problems

def process_checked(wb, log):
    before = snapshot_values(wb)
    process_wb(wb, log)
    problems = verify_no_value_change(before, wb)
    if problems:
        raise ValueError("安全チェック失敗（セル値の変更を検知したため中止）: " + " / ".join(problems))
    log("  ✔ 安全チェック合格：セルの値は不変（削除・書式のみ）")
    return wb

def process(path, log):
    return process_checked(openpyxl.load_workbook(path), log)

def resolve_keireki():
    if "--delete-keireki" in sys.argv: return True
    if "--keep-keireki" in sys.argv:   return False
    try:
        ans = input("「経歴」シートを削除しますか? [y=削除 / N=保持] : ").strip().lower()
    except EOFError:
        ans = ""
    return ans in ("y", "yes")

def processed_name(name):
    base, ext = os.path.splitext(name)
    return base + "_processed" + (ext or ".xlsx")

def run(inputs, delete_keireki, progress=None):
    global DELETE_KEIREKI
    DELETE_KEIREKI = delete_keireki
    def emit(s):
        (progress or (lambda x: print(x, flush=True)))(s)
    files = []; base = None
    for p in inputs:
        if os.path.isdir(p):
            base = p
            files += sorted(glob.glob(os.path.join(p, "*.xlsx")))
        elif p.lower().endswith(".xlsx"):
            files.append(p)
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    if not files:
        emit("対象の .xlsx が見つかりませんでした。")
        return None, []
    if base is None:
        base = os.path.dirname(os.path.abspath(files[0]))
    out_dir = os.path.join(base, "_processed")
    logs_dir = os.path.join(base, "_logs")
    os.makedirs(out_dir, exist_ok=True); os.makedirs(logs_dir, exist_ok=True)
    emit("経歴シート: " + ("削除" if delete_keireki else "保持"))
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    results = []
    for f in files:
        name = os.path.basename(f)
        if os.path.dirname(os.path.abspath(f)) == os.path.abspath(out_dir):
            continue
        lines = []
        def log(s, lines=lines): lines.append(s); emit(s)
        emit(f"\n===== {name} =====")
        try:
            wb = process(f, log)
            out_name = processed_name(name)
            wb.save(os.path.join(out_dir, out_name))
            log(f"保存 -> _processed/{out_name}")
            results.append((name, "OK"))
        except Exception as e:
            log(f"!! エラー: {e}")
            results.append((name, f"ERROR: {e}"))
        with open(os.path.join(logs_dir, f"{name}.{stamp}.log"), "w", encoding="utf-8") as fp:
            fp.write("\n".join(lines))
    ok = sum(1 for _, s in results if s == "OK")
    emit(f"\n完了: {ok}/{len(results)} ファイル処理。出力先: {out_dir}")
    return out_dir, results

def main():
    inputs = [a for a in sys.argv[1:] if not a.startswith("-")]
    if not inputs:
        inputs = [BASE]
    run(inputs, resolve_keireki())

if __name__ == "__main__":
    main()
