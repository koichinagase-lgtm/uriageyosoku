#!/usr/bin/env python3
"""売上予測くん：完売率カーブ再生成スクリプト

JR東日本クロスステーション提供の「日別ショップ別時間帯別売上」Excelを読み込み、
曜日別（月〜日）＋祝日の完売率カーブを作って index.html の DB / HOLIDAYS を書き換える。

使い方:
  pip install pandas openpyxl jpholiday
  python3 scripts/build_curves.py                      # 既定フォルダ ~/Documents/claude/sales
  python3 scripts/build_curves.py --src <フォルダ>      # 元データの場所を指定
  python3 scripts/build_curves.py --from 2026-06-01     # 期間を絞る
  python3 scripts/build_curves.py --dry-run             # 集計結果を表示するだけ

元データ: ファイル名に「日別ショップ別時間帯別売上」を含む .xlsx をすべて読む（期間が重なる日は後のファイルを優先）。
"""
import argparse, datetime as dt, glob, json, os, re, sys
import pandas as pd
import jpholiday

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
HTML = os.path.join(ROOT, "index.html")
DAY = "月火水木金土日"  # Python weekday() 順

# ---- 機会損失率（上限＝下限×(1+率)） ----------------------------------------
# 出典: Artifact「グランスタ店 機会損失分析」（2026/4/15–8/31・139日、商品×時刻の欠品推定）
#   曜日別の1日あたり機会損失額(税抜) ÷ 同曜日の1日あたり実売上(税抜)
# 祝日は個別集計がないため日曜と同じ値を暫定使用（5/3・8/11 など祝日の損失は大きめ）。
# 機会損失分析をやり直したらここを更新する。
OPP_LOSS_PER_DAY = {"月": 81042, "火": 38862, "水": 43950, "木": 53544,
                    "金": 69881, "土": 47865, "日": 80154}
OPP_PERIOD = ("2026-04-15", "2026-08-31")
HOLIDAY_OPP_FROM = "日"


def load_file(path):
    d = pd.read_excel(path, header=None)
    hdr_row = next(i for i in range(min(15, len(d)))
                   if any(re.match(r"\d\d:00～", str(v)) for v in d.iloc[i]))
    cols = {i: int(str(v)[:2]) for i, v in enumerate(d.iloc[hdr_row])
            if re.match(r"\d\d:00～", str(v))}
    rows = []
    for _, r in d.iloc[hdr_row + 1:].iterrows():
        if not isinstance(r[0], (pd.Timestamp, dt.datetime)):
            continue
        rec = {"date": pd.Timestamp(r[0]).date()}
        for i, h in cols.items():
            v = r[i]
            rec[h] = float(v) if pd.notna(v) else 0.0
        rows.append(rec)
    return pd.DataFrame(rows)


def load_all(src):
    files = sorted(glob.glob(os.path.join(src, "*日別ショップ別時間帯別売上*.xlsx")))
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    if not files:
        sys.exit(f"元データが見つかりません: {src}")
    df = pd.concat([load_file(f) for f in files], ignore_index=True).fillna(0.0)
    df = df.drop_duplicates("date", keep="last").sort_values("date").reset_index(drop=True)
    return df, files


def day_type(d):
    return "祝" if jpholiday.is_holiday(d) else DAY[d.weekday()]


def build(df):
    hours = sorted(c for c in df.columns if isinstance(c, int))
    df["total"] = df[hours].sum(axis=1)
    df = df[df.total > 0].copy()
    # 売上が一度も立たない時間帯は端から落とす（開店前・閉店後）
    active = [h for h in hours if df[h].sum() > 0]
    h0, h1 = active[0], active[-1] + 1
    slots = list(range(h0, h1))
    cum = df[slots].cumsum(axis=1).div(df.total, axis=0) * 100
    df["type"] = [day_type(d) for d in df.date]

    wd_sales = df.groupby(df.date.map(lambda d: DAY[d.weekday()])).total.mean() / 1.08
    curves = {}
    for t in list(DAY) + ["祝"]:
        mask = (df.type == t).values
        n = int(mask.sum())
        if n == 0:
            continue
        mean = cum[mask].mean()
        pts = [{"t": f"{h0:02d}:00", "r": 0}]
        for h in slots:
            pts.append({"t": f"{h + 1:02d}:00", "r": round(float(mean[h]), 1)})
        base = HOLIDAY_OPP_FROM if t == "祝" else t
        opp = OPP_LOSS_PER_DAY[base] / wd_sales[base]
        curves[t] = {"days": n, "points": pts, "oppRatio": round(float(opp), 3)}
    meta = {"start": str(df.date.min()), "end": str(df.date.max()), "days": int(len(df)),
            "built": dt.date.today().isoformat(),
            "oppSource": f"機会損失分析（{OPP_PERIOD[0]}〜{OPP_PERIOD[1]}）"}
    return {"meta": meta, "curves": curves}


def holiday_list(start_year, years=3):
    out = []
    for y in range(start_year, start_year + years):
        out += [d.isoformat() for d, _ in jpholiday.year_holidays(y)]
    return sorted(out)


def write_html(db, hol):
    with open(HTML, encoding="utf-8") as f:
        s = f.read()
    db_js = "const DB=" + json.dumps(db, ensure_ascii=False, separators=(",", ":")) + ";"
    hol_js = "const HOLIDAYS=new Set(" + json.dumps(hol, separators=(",", ":")) + ");"
    s, n1 = re.subn(r"^const DB=.*;$", lambda m: db_js, s, count=1, flags=re.M)
    if re.search(r"^const HOLIDAYS=.*;$", s, flags=re.M):
        s = re.sub(r"^const HOLIDAYS=.*;$", lambda m: hol_js, s, count=1, flags=re.M)
    else:
        s = s.replace(db_js, db_js + "\n" + hol_js, 1)
    if n1 != 1:
        sys.exit("index.html に 'const DB=' の行が見つかりません")
    with open(HTML, "w", encoding="utf-8") as f:
        f.write(s)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=os.path.expanduser("~/Documents/claude/sales"))
    ap.add_argument("--from", dest="date_from")
    ap.add_argument("--to", dest="date_to")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()

    df, files = load_all(a.src)
    if a.date_from:
        df = df[df.date >= dt.date.fromisoformat(a.date_from)]
    if a.date_to:
        df = df[df.date <= dt.date.fromisoformat(a.date_to)]
    db = build(df)

    print("読込:", *[os.path.basename(f) for f in files], sep="\n  ")
    m = db["meta"]
    print(f"期間 {m['start']}〜{m['end']}（{m['days']}日）")
    for t, c in db["curves"].items():
        r = {p["t"]: p["r"] for p in c["points"]}
        print(f"  {t}: {c['days']:>2}日  12時{r.get('12:00', 0):5.1f}%  17時{r.get('17:00', 0):5.1f}%"
              f"  機会損失率{c['oppRatio'] * 100:4.1f}%")
    if a.dry_run:
        return
    write_html(db, holiday_list(int(m["end"][:4])))
    print("index.html を更新しました")


if __name__ == "__main__":
    main()
