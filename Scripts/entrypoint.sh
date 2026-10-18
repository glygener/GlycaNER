#!/bin/bash
#
# GlycaNER container entrypoint.
#
# Usage (arguments after the image name in `docker run`):
#   all                         BERT detection, then post-processing (default)
#   detection                   only BERT detection        -> /data/output/bert_detection
#   post_processing             only post-processing       -> /data/output/post_processing
#   from <script>               resume at <script> and run everything after it
#   list                        print every script in order
#
# Key results are copied to /data/output/final.

set -u

OUT="${GLYCANER_OUTPUT:-/data/output}"
export GLYCANER_STAGE1_DIR="${GLYCANER_STAGE1_DIR:-$OUT/bert_detection}"
export GLYCANER_STAGE2_DIR="${GLYCANER_STAGE2_DIR:-$OUT/post_processing}"
export GLYCANER_STAGE3_DIR="${GLYCANER_STAGE3_DIR:-$GLYCANER_STAGE2_DIR}"
export GLYCANER_STAGE3_INPUT="${GLYCANER_STAGE3_INPUT:-$GLYCANER_STAGE2_DIR/gst_punct_removed_cleaned.xlsx}"
FINAL_DIR="$OUT/final"
LOG_ROOT="$OUT/logs"

DET=/app/Scripts/BERT_Detection
PP=/app/Scripts/Post_Processing

DETECTION_STEPS=(1_pubtator_api 2_text_file 3_get_sentence 1_gst 2_gst update_gst 4_gst)
mapfile -t PP_STEPS < <(sed -n '/^SCRIPTS=(/,/^)/p' "$PP/run_pipeline.sh" | sed '1d;$d' | tr -d ' ')

log() { echo "[glycaner] $*"; }

run_py() {  # run_py <step> [args...]
    local step=$1; shift
    local logfile="$LOG_ROOT/${step}.log"
    log "Running $step.py"
    python "$DET/$step.py" "$@" 2>&1 | tee "$logfile"
    local status=${PIPESTATUS[0]}
    if [ "$status" -ne 0 ]; then
        log "FAILED: $step.py (exit $status). Log: $logfile"
        log "Resume after fixing with:  from $step"
        exit "$status"
    fi
}

detection_step() {
    case "$1" in
        1_pubtator_api)
            if [ ! -s "$GLYCANER_PMID_FILE" ]; then
                log "No PMID list mounted at $GLYCANER_PMID_FILE; using the bundled example list."
                export GLYCANER_PMID_FILE=/app/examples/pmids_example.txt
            fi
            log "PMID list: $GLYCANER_PMID_FILE ($(grep -c . "$GLYCANER_PMID_FILE") PMIDs)"
            run_py 1_pubtator_api ;;
        update_gst)
            run_py update_gst \
                --input "$GLYCANER_STAGE1_DIR/output_predictions1.csv" \
                --output "$GLYCANER_STAGE1_DIR/final_gst1.xlsx" ;;
        *)  run_py "$1" ;;
    esac
}

run_detection() {  # run_detection [start_step]
    local start="${1:-}" go=0
    [ -z "$start" ] && go=1
    log "=== BERT detection ==="
    for s in "${DETECTION_STEPS[@]}"; do
        [ "$s" == "$start" ] && go=1
        [ $go -eq 1 ] && detection_step "$s"
    done
}

run_post_processing() {  # run_post_processing [start_script]
    log "=== Post-processing and normalization ==="
    [ -f "$GLYCANER_STAGE1_DIR/final_gst1.xlsx" ] || { log "Missing $GLYCANER_STAGE1_DIR/final_gst1.xlsx (run detection first)"; exit 1; }
    PYTHON=python SCRIPT_DIR="$PP" LOG_DIR="$LOG_ROOT" "$PP/run_pipeline.sh" ${1:-} || exit $?
}

collect_final() {
    mkdir -p "$FINAL_DIR"
    for f in "$GLYCANER_STAGE3_DIR/9c_normalized_form.xlsx" \
             "$GLYCANER_STAGE3_DIR/8_final_gst_string_classified.xlsx" \
             "$GLYCANER_STAGE1_DIR/final_gst1.xlsx" \
             "$GLYCANER_STAGE1_DIR/Highlighted_ML_Predictions.html"; do
        [ -f "$f" ] && cp "$f" "$FINAL_DIR/"
    done
    log "Final results in $FINAL_DIR:"
    ls -1 "$FINAL_DIR" | sed 's/^/    /'
}

contains() { local x=$1; shift; for e in "$@"; do [ "$e" == "$x" ] && return 0; done; return 1; }

list_steps() {
    echo "BERT detection:"; printf '    %s\n' "${DETECTION_STEPS[@]}"
    echo "Post-processing:"; printf '    %s\n' "${PP_STEPS[@]}"
}

mkdir -p "$GLYCANER_STAGE1_DIR" "$GLYCANER_STAGE2_DIR" "$GLYCANER_STAGE3_DIR" "$LOG_ROOT"

CMD="${1:-all}"
START=$(date +%s)
case "$CMD" in
    all)             run_detection; run_post_processing; collect_final ;;
    detection)       run_detection ;;
    post_processing) run_post_processing; collect_final ;;
    from)
        step="${2:-}"; step="${step%.py}"
        if contains "$step" "${DETECTION_STEPS[@]}"; then
            run_detection "$step"; run_post_processing; collect_final
        elif contains "$step" "${PP_STEPS[@]}"; then
            run_post_processing "$step"; collect_final
        else
            echo "Unknown script '$step'. Valid scripts:"; list_steps; exit 1
        fi ;;
    list)   list_steps; exit 0 ;;
    help|-h|--help) sed -n '2,12p' "$0" | sed 's/^# \{0,1\}//'; exit 0 ;;
    *)      echo "Unknown command '$CMD'. Try: all | detection | post_processing | from <script> | list"; exit 1 ;;
esac
log "Done in $(( $(date +%s) - START ))s"
