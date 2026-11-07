#!/bin/zsh
# 再評価を 1 つの命令で（docs/再評価の計画.md）。nice 10・重い処理は一度に一つ・止めても同じ命令で続きから。Office は動かさない。
#   run_freeze2.sh stage1              固定版（eval-freeze-1）の道具を採点 v2 で、同じ世界の一部（1,079 件）に
#   run_freeze2.sh stage2              凍結 2（eval-freeze-2。本人がタグを打つ）の道具を、新しい標本に
#   run_freeze2.sh estimate stage1|stage2    数だけを出し直す
#   run_freeze2.sh cleanup stage1|stage2     確かめと数のあとで、書類を消す
#   run_freeze2.sh jpforms             未見の日本の様式（無作為でない別の集まり。第 2 段の道具で。世界の推定に足さない）
#   run_freeze2.sh dryrun              すでに見た書類の数件で通しの試運転（mock の印つき・別のフォルダ）
# 置き場：FREEZE2_OUT（既定は下の S/freeze2）。動画の world_counts.json・ポータルの data.json には書かない。本物に入れるのは本人の判断のあと。
# nice を重ねない：nice 20 は v1 の大量の失敗の条件。すでに nice 10 以上なら下げない。
setopt NO_BG_NICE
set -u
SCRIPT_DIR="${0:A:h}"
ROOT="${SCRIPT_DIR:h:h:h}"
S="${FREEZE2_SCRATCH:?作業のフォルダを FREEZE2_SCRATCH に入れてください}"
BASE="${FREEZE2_OUT:-$S/freeze2}"
SUBSET1="${FREEZE2_SUBSET1:-$S/final/world/subset.json}"
export TODOKU_BASELINE_PY="${TODOKU_BASELINE_PY:-$S/m/ev/baseline_venv/bin/python}"
export TODOKU_PY="${TODOKU_PY:-$ROOT/.venv/bin/python}"
PKR="${FREEZE2_PKR:-$S/final/tools/pkr}"
PY="$(command -v python3)"
RUN=("$PY" -I "$SCRIPT_DIR/runner.py")
BUDGET2="${FREEZE2_BUDGET_HOURS:-6}"  # 比べの上限。索引と取る時間（約 1.7 時間）と合わせて 8 時間に収める
DOCS_FROM=(); [[ -n "${FREEZE2_DOCS_FROM:-}" ]] && DOCS_FROM=(--docs-from "$FREEZE2_DOCS_FROM")
BUSY=(); [[ -n "${FREEZE2_IGNORE_BUSY:-}" ]] && BUSY=(--ignore-busy)  # 待っても終わらないときだけ。記録に書く
EXCL_DIR=(); [[ -n "${FREEZE2_EXCLUDE_DIR:-}" ]] && EXCL_DIR=(--exclude-dir "$FREEZE2_EXCLUDE_DIR")

cur=$(ps -o ni= -p $$ | tr -d ' ')
if (( cur < 10 )); then renice -n $((10 - cur)) -p $$ >/dev/null 2>&1; fi

clean_scorer() {  # 採点のコードが commit と違うまま本番を動かさない（記録の commit と中身が合うように）
  local d; d="$(git -C "$ROOT" status --porcelain scripts/eval/v2)"
  [[ -z "$d" ]] || { print -u2 "採点のコードに commit していない変更があります：\n$d"; exit 2; }
}

jpforms() {
  local O2="$BASE/stage2" J="$BASE/jp_forms"
  [[ -d "$O2/tool/src" ]] || { print -u2 "第 2 段の道具がありません（タグ eval-freeze-2 のあとで）"; return 2; }
  export EVAL_LOCK="$J/.heavy.lock"; mkdir -p "$J"
  [[ -f "$J/files.txt" ]] || "$PY" -I "$SCRIPT_DIR/jp_forms.py" fetch --out "$J" --subset "$SUBSET1" || return $?
  local C=(--files "$J/files.txt" --baseline-python "$TODOKU_BASELINE_PY" --todoku-python "$TODOKU_PY" --todoku-src "$O2/tool/src" --pkr "$PKR" --purge --label "未見の日本の様式（無作為でない別の集まり）")
  "$PY" -I "$SCRIPT_DIR/read_compare.py" --out "$J/read" $C > "$J/read.log" 2>&1 || return $?
  "$PY" -I "$SCRIPT_DIR/write_compare.py" --out "$J/write" --no-signed $C > "$J/write.log" 2>&1 || return $?
  print "日本の様式：$J/read/summary.md・$J/write/summary.md（世界の推定に足さない）"
}

step() { "$@"; local rc=$?; if (( rc )); then print -u2 "止まった（rc=$rc）：$*。直してから同じ命令で続きから動きます"; exit $rc; fi; }

case "${1:-}" in
  stage1)
    clean_scorer
    O="$BASE/stage1"; mkdir -p "$O"
    step $RUN tool --out "$O" --tag eval-freeze-1
    step $RUN preflight --out "$O" --pkr "$PKR"
    step $RUN fetch --out "$O" --subset "$SUBSET1" $DOCS_FROM $BUSY
    step $RUN score --out "$O" --pkr "$PKR" $BUSY
    step $RUN verify --out "$O"
    step $RUN estimate --out "$O" --subset "$SUBSET1" --label "第 1 段：eval-freeze-1 × 採点 v2"
    ;;
  stage2)
    clean_scorer
    O="$BASE/stage2"; mkdir -p "$O"
    git -C "$ROOT" rev-parse -q --verify "eval-freeze-2^{commit}" >/dev/null || { print -u2 "タグ eval-freeze-2 がありません（本人が main に打つ）"; exit 2; }
    step $RUN tool --out "$O" --tag eval-freeze-2
    step $RUN preflight --out "$O" --pkr "$PKR"
    if [[ ! -f "$O/subset.json" ]]; then
      step "$PY" -I "$SCRIPT_DIR/stage2_select.py" select --out "$O" --subset "$SUBSET1" $EXCL_DIR
    fi
    step $RUN fetch --out "$O" --subset "$O/subset.json"
    $RUN score --out "$O" --pkr "$PKR" --budget-hours "$BUDGET2"; rc=$?
    if (( rc == 74 )); then
      print "時間の上限で止めた。無作為な順の、ここまでの書類で数を出す"
      "$PY" -I "$SCRIPT_DIR/stage2_trim.py" "$O" || exit 2
    elif (( rc )); then
      print -u2 "止まった（rc=$rc）。同じ命令で続きから動きます"; exit $rc
    fi
    step $RUN verify --out "$O"
    step $RUN estimate --out "$O" --subset "$O/subset.json" --label "第 2 段：eval-freeze-2 × 採点 v2 × 新しい標本"
    if (( SECONDS < 26280 )); then jpforms || print -u2 "日本の様式は止まった（rc=$?）。第 2 段の数には関わらない"; else print "時間がないので日本の様式は測らなかった（$SECONDS 秒）"; fi
    ;;
  jpforms) jpforms; exit $?;;
  estimate)
    which="${2:?stage1 か stage2}"; O="$BASE/$which"
    [[ "$which" == stage1 ]] && SUB="$SUBSET1" || SUB="$O/subset.json"
    step $RUN verify --out "$O"
    step $RUN estimate --out "$O" --subset "$SUB" --label "$which"
    ;;
  cleanup)
    which="${2:?stage1 か stage2}"; step $RUN cleanup --out "$BASE/$which"
    ;;
  dryrun)
    O="$BASE/mock_dryrun"; mkdir -p "$O"
    ONLY="${FREEZE2_DRY_IDS:-AD-pdf-w019,AE-sheet-002,AM-text-w014,AR-csv-001,CH-csv-001,SE-csv-w018,IL-pdf-014,TZ-pdf-w014}"
    step $RUN tool --out "$O" --tag eval-freeze-1
    step $RUN preflight --out "$O" --pkr "$PKR"
    step $RUN fetch --out "$O" --subset "$SUBSET1" --only "$ONLY" --ignore-busy $DOCS_FROM
    step $RUN score --out "$O" --pkr "$PKR" --ignore-busy
    step $RUN verify --out "$O"
    step $RUN estimate --out "$O" --subset "$SUBSET1" --label "試運転（mock）：すでに見た数件" --mock
    step "$PY" -I "$SCRIPT_DIR/check_shapes.py" "$O/publish_candidate" "$ROOT/world/portal/data.json" "$S/film/src/v3/maps.py"
    step $RUN cleanup --out "$O"
    ;;
  *) print -u2 "使い方：run_freeze2.sh stage1|stage2|jpforms|estimate <段>|cleanup <段>|dryrun"; exit 2;;
esac
print "終わり：$BASE"
