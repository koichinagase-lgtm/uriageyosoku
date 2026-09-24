# 売上予測くん（ONI & Co. グランスタ東京店）

現時点の売上（税込・レジ別）から、終日の売上（税抜）を「下限〜上限」のレンジで予測する1ページのWebアプリ。Vercelで公開（noindex）。

## 予測の仕組み
- **下限** ＝ 現在売上(税抜) ÷ その時刻の完売率（曜日別または祝日の平均カーブ）
- **上限** ＝ 下限 ×（1＋機会損失率）
- 祝日は日付から自動で「祝日カーブ」に切り替わる。年末年始・連休中の平日などは画面のチェックで手動ON

## データ更新の手順
1. JR東日本クロスステーションから届く「日別ショップ別時間帯別売上」の .xlsx を `~/Documents/claude/sales` に置く（既存ファイルと期間が重なってもOK。後のファイルが優先）
2. リポジトリで実行
   ```
   pip install pandas openpyxl jpholiday   # 初回のみ
   python3 scripts/build_curves.py --dry-run   # 集計結果の確認
   python3 scripts/build_curves.py             # index.html を更新
   ```
   期間を絞るときは `--from 2026-06-01` など
3. commit & push → Vercel に自動反映

## 補足
- 機会損失率は Artifact「グランスタ店 機会損失分析」（2026/4/15–8/31）の曜日別損失額 ÷ 曜日別売上。`scripts/build_curves.py` の `OPP_LOSS_PER_DAY` で管理。祝日は日曜の値を暫定使用
- 祝日リストはデータ最終年から3年分を `jpholiday` で生成して index.html に埋め込む（年が変わったら再実行）
