#!/bin/zsh
# 最終の評価を一度に動かす。
#   run_all.sh --frozen-tag TAG --out 出力 --dev 開発の様式 --tama とっておいた様式 --world とっておいた様式 --baseline-venv 比べる道具の venv
# 試運転（開発の書類だけ。とっておいた書類には触れない）:
#   run_all.sh --frozen-tag TAG --out 出力 --dev 開発の様式 --baseline-venv V --dry-run [--dev-list 一覧] [--dry-forms 3]
#
# - 作業ツリーが汚れていたら動かない。HEAD がタグでなければ動かない。コミットと道具の版を provenance.json に残す
# - nice 10。重い処理は一度に一つ（全部の処理が同じ錠を使う。記入のエージェントは待つだけなら軽いので、いくつか並べてよい）
# - 途中で止めても、同じ命令でもう一度動かせば続きから（.done/ の印と、書類ごとの結果）
# - 結果は --out の下にだけ書く。リポジトリの中には書かない
# - 処理の順：読む・書く・処理（dev → tama → world）→ 記入 → 調査の付け直し → 新しく届く → 数 → まとめ
#   --skip-fill・--skip-world で一部を飛ばせる（飛ばしたものは report.md に「結果がない」と出る）

set -u
SCRIPT_DIR="${0:A:h}"
ROOT="${SCRIPT_DIR:h:h}"
TAG=""; OUT=""; DEV=""; TAMA=""; WORLD=""; VENV=""; TODOKU_PY=""; PDFKIT_DIR=""; DRY=0; DEV_LIST=""; DRY_FORMS=3; AGENTS=3; MODEL="opus"; SKIP_FILL=0; SKIP_WORLD=0; PER_STRATUM=2

while (( $# )); do
  case "$1" in
    --frozen-tag) TAG="$2"; shift 2;;
    --out) OUT="$2"; shift 2;;
    --dev) DEV="$2"; shift 2;;
    --dev-list) DEV_LIST="$2"; shift 2;;
    --tama) TAMA="$2"; shift 2;;
    --world) WORLD="$2"; shift 2;;
    --baseline-venv) VENV="$2"; shift 2;;
    --todoku-python) TODOKU_PY="$2"; shift 2;;
    --pdfkit-dir) PDFKIT_DIR="$2"; shift 2;;
    --fill-agents) AGENTS="$2"; shift 2;;
    --model) MODEL="$2"; shift 2;;
    --per-stratum) PER_STRATUM="$2"; shift 2;;
    --dry-forms) DRY_FORMS="$2"; shift 2;;
    --dry-run) DRY=1; shift;;
    --skip-fill) SKIP_FILL=1; shift;;
    --skip-world) SKIP_WORLD=1; shift;;
    *) print -u2 "知らない引数：$1"; exit 2;;
  esac
done
[[ -n "$TAG" && -n "$OUT" && -n "$VENV" ]] || { print -u2 "--frozen-tag・--out・--baseline-venv が要ります"; exit 2; }
[[ -n "$DEV" || -n "$DEV_LIST" ]] || { print -u2 "--dev か --dev-list が要ります"; exit 2; }
if (( ! DRY )); then
  [[ -n "$TAMA" && -n "$WORLD" ]] || { print -u2 "--tama と --world が要ります（試運転なら --dry-run）"; exit 2; }
else
  export EVAL_DEV_ONLY=1   # 試運転：とっておいた書類のパスは開く前に止まる
fi

# ---------------------------------------------------------------------- 凍結の確かめ
cd "$ROOT" || exit 2
DIRTY="$(git status --porcelain)"
if [[ -n "$DIRTY" ]]; then
  print -u2 "作業ツリーが汚れているので動きません："; print -u2 "$DIRTY"; exit 2
fi
HEAD_SHA="$(git rev-parse HEAD)"
TAG_SHA="$(git rev-parse "${TAG}^{commit}" 2>/dev/null)" || { print -u2 "タグが見つかりません：$TAG"; exit 2; }
if [[ "$HEAD_SHA" != "$TAG_SHA" ]]; then
  print -u2 "HEAD（$HEAD_SHA）がタグ $TAG（$TAG_SHA）ではありません。タグを取り出してから動かしてください"; exit 2
fi
OUT="${OUT:A}"
case "$OUT/" in
  "$ROOT"/*) print -u2 "結果の置き場はリポジトリの外にしてください：$OUT"; exit 2;;
esac
mkdir -p "$OUT/.done" "$OUT/lists" "$OUT/tools" "$OUT/logs"

# ---------------------------------------------------------------------- 二重に動かさない
LOCKD="$OUT/.running"
if ! mkdir "$LOCKD" 2>/dev/null; then
  OLD="$(cat "$LOCKD/pid" 2>/dev/null)"
  if [[ -n "$OLD" ]] && kill -0 "$OLD" 2>/dev/null; then
    print -u2 "すでに動いています（pid $OLD）"; exit 2
  fi
  print "前の実行は止まっています（pid ${OLD:-不明}）。続きから動かします"
fi
print $$ > "$LOCKD/pid"
trap 'rm -rf "$LOCKD"' EXIT
renice -n 10 -p $$ >/dev/null 2>&1 || renice 10 -p $$ >/dev/null 2>&1
export EVAL_LOCK="$OUT/.heavy.lock"

# ---------------------------------------------------------------------- 道具
BPY="$VENV/bin/python"
[[ -x "$BPY" ]] || zsh "$SCRIPT_DIR/setup_baselines.sh" "$VENV" || { print -u2 "比べる道具の venv を作れません"; exit 2; }
if [[ -z "$TODOKU_PY" ]]; then
  if [[ -x "$ROOT/.venv/bin/python" ]]; then TODOKU_PY="$ROOT/.venv/bin/python"; else TODOKU_PY="$(command -v python3)"; fi
fi
PY="$(command -v python3)"

# PDFKit の道具：リポジトリの中か --pdfkit-dir に pkr.swift があれば作る。なければ、ないまま進む
PKR=""
for d in "$PDFKIT_DIR" "$ROOT/tests/audit_findings/pdfkit"; do
  if [[ -n "$d" && -f "$d/pkr.swift" ]]; then
    if [[ ! -x "$OUT/tools/pkr" ]]; then swiftc -O "$d/pkr.swift" -o "$OUT/tools/pkr" >"$OUT/logs/pkr.log" 2>&1 || print "PDFKit の道具を作れない（なしで進む）"; fi
    break
  fi
done
[[ -x "$OUT/tools/pkr" ]] && PKR="$OUT/tools/pkr"
PKR_ARG=(); [[ -n "$PKR" ]] && PKR_ARG=(--pkr "$PKR")

DRY_ARG=(); (( DRY )) && DRY_ARG=(--dry-run)
"$PY" "$SCRIPT_DIR/provenance.py" --out "$OUT" --tag "$TAG" --commit "$HEAD_SHA" --baseline-python "$BPY" --todoku-python "$TODOKU_PY" $PKR_ARG $DRY_ARG || exit 2

step() {  # step 名前 命令...：印があれば飛ばす。失敗したら止める（もう一度動かせば、その段階の続きから）
  local name="$1"; shift
  [[ -f "$OUT/.done/$name" ]] && { print "== $name 済み"; return 0; }
  print "== $name $(date '+%H:%M:%S')"
  local t0=$SECONDS
  "$@" 2>&1 | tee -a "$OUT/logs/$name.log" | tail -n 3
  local rc=${pipestatus[1]}
  if (( rc != 0 )); then
    print -u2 "段階 $name が失敗（rc=$rc）。記録：$OUT/logs/$name.log。直してから同じ命令で続きから動きます"
    exit $rc
  fi
  print "$((SECONDS - t0))" > "$OUT/.done/$name"
}

# ---------------------------------------------------------------------- 一覧（ここで初めて、とっておいた書類のフォルダを一覧する）
make_list() {  # make_list 名前 フォルダ
  local name="$1" dir="$2" f="$OUT/lists/$1.txt"
  [[ -s "$f" ]] && return 0
  "$PY" -c "import sys; sys.path.insert(0, sys.argv[1]); from common import walk_files; print('\n'.join(walk_files(sys.argv[2])))" "$SCRIPT_DIR" "$dir" > "$f" || exit 2
}
if [[ -n "$DEV_LIST" ]]; then cp -n "$DEV_LIST" "$OUT/lists/dev.txt" 2>/dev/null; else make_list dev "$DEV"; fi
SETS=(dev)
if (( ! DRY )); then
  make_list tama "$TAMA"; make_list world "$WORLD"
  SETS=(dev tama world)
fi

# ---------------------------------------------------------------------- 読む・書く・処理
COMMON=(--baseline-python "$BPY" --todoku-python "$TODOKU_PY")
for s in $SETS; do
  L="$OUT/lists/$s.txt"
  step "read_$s" "$PY" "$SCRIPT_DIR/read_compare.py" --files "$L" --out "$OUT/read/$s" $COMMON --label "$s" $PKR_ARG
  step "write_$s" "$PY" "$SCRIPT_DIR/write_compare.py" --files "$L" --out "$OUT/write/$s" $COMMON --label "$s" $PKR_ARG
  step "processing_$s" "$PY" "$SCRIPT_DIR/processing_eval.py" --files "$L" --out "$OUT/processing/$s" $COMMON --label "$s"
  if (( ! SKIP_FILL )); then  # 計算の説明の確かめ（エージェントを使う）
    step "explain_$s" "$PY" "$SCRIPT_DIR/explain_check.py" --processing "$OUT/processing/$s" --out "$OUT/explain/$s" $COMMON --n "$( (( DRY )) && print 1 || print 5 )" --model "$MODEL"
  fi
done

# ---------------------------------------------------------------------- 記入（見ないで）
if (( ! SKIP_FILL )); then
  FILL=(--out "$OUT/fill" $COMMON --agents "$AGENTS" --model "$MODEL" $PKR_ARG)
  if (( DRY )); then
    step fill_dev "$PY" "$SCRIPT_DIR/fill_runner.py" run --set dev --root "$DEV" --limit "$DRY_FORMS" $FILL
  else
    step fill_dev "$PY" "$SCRIPT_DIR/fill_runner.py" run --set dev --root "$DEV" $FILL
    step fill_tama "$PY" "$SCRIPT_DIR/fill_runner.py" run --set tama --root "$TAMA" $FILL
    step fill_world "$PY" "$SCRIPT_DIR/fill_runner.py" run --set world --root "$WORLD" $FILL
  fi
fi

# ---------------------------------------------------------------------- 世界の数
if (( ! SKIP_WORLD )); then
  mkdir -p "$OUT/world"
  if (( DRY )); then
    # 試運転：調査の書類を取り直さない（開発の書類だけで試す約束）。リポジトリの記録（凍結の前の版の段階が混ざる）で計算だけ確かめる
    [[ -f "$OUT/world/results.jsonl" ]] || cp "$ROOT/world/results.jsonl" "$OUT/world/results.jsonl"
    step world_pick "$PY" "$SCRIPT_DIR/world_counts.py" pick --results "$OUT/world/results.jsonl" --out "$OUT/world" --per-stratum "$PER_STRATUM"
    step world_mock "$PY" "$SCRIPT_DIR/world_counts.py" mock --out "$OUT/world" --read-scores "$OUT/read/dev/scores" --write-scores "$OUT/write/dev/scores"
    step world_estimate "$PY" "$SCRIPT_DIR/world_counts.py" estimate --results "$OUT/world/results.jsonl" --out "$OUT/world" --label "試運転：凍結の前の記録"
  else
    [[ -f "$OUT/world/results.jsonl" ]] || cp "$ROOT/world/results.jsonl" "$OUT/world/results.jsonl"
    step world_regrade "$PY" "$SCRIPT_DIR/world_counts.py" regrade --results "$OUT/world/results.jsonl" --graders 1
    step world_pick "$PY" "$SCRIPT_DIR/world_counts.py" pick --results "$OUT/world/results.jsonl" --out "$OUT/world" --per-stratum "$PER_STRATUM"
    step world_compare "$PY" "$SCRIPT_DIR/world_counts.py" compare --out "$OUT/world" $COMMON $PKR_ARG
    step world_estimate "$PY" "$SCRIPT_DIR/world_counts.py" estimate --results "$OUT/world/results.jsonl" --out "$OUT/world" --label "凍結の版 $TAG"
  fi
fi

# ---------------------------------------------------------------------- まとめ
rm -f "$OUT/.done/report"
step report "$PY" "$SCRIPT_DIR/report.py" --out "$OUT" $DRY_ARG
[[ -z "$(git status --porcelain)" ]] || print "注意：評価のあいだに作業ツリーが変わった：$(git status --porcelain | head -5)"
print "終わり：$OUT/report.md"
