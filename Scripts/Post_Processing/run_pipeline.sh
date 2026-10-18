#!/bin/bash
#
# Runs the GlycaNER post-processing scripts in order, starting from the
# BioBERT terms in final_gst1.xlsx.
#
# Usage:
#   ./run_pipeline.sh                 run every script
#   ./run_pipeline.sh 3_merge_np      start from this script and run to the end
#
# Behaviour:
#   Stops at the first script that fails, so a broken stage does not silently
#   feed bad input to the next one.
#   Writes a timestamped log for every script into ./logs

set -u

PYTHON="${PYTHON:-python3}"
SCRIPT_DIR="${SCRIPT_DIR:-$(cd "$(dirname "$0")" && pwd)}"
LOG_DIR="${LOG_DIR:-$SCRIPT_DIR/logs}"

SCRIPTS=(
    before_11
    11_merge
    13_remove_punct
    before_1_modify_bert_output
    before_1_fix_bert
    1_base_np
    1b_colon_slash_handle
    2_remove_ending
    2b_merge_connector_base_np
    2c_same_base_np
    2d_extra_bert_update
    2e_nested_base_np_update
    2f_conj_additional_base_np
    2g_bert_adjectival_merge
    3_merge_np
    4_merged_gst
    5_extend_bert
    6a_remove_parenthetical_part
    6b_conjunction_parentheses_resolve
    7_cleanup
    8_categorize_terms
    9a_normalization1
    9b_normalization2
    9c_normalization3_modify
)

mkdir -p "$LOG_DIR"

# Optional starting point
START_FROM="${1:-}"
STARTED=1
if [ -n "$START_FROM" ]; then
    STARTED=0
    FOUND=0
    for s in "${SCRIPTS[@]}"; do
        if [ "$s" == "$START_FROM" ] || [ "$s" == "${START_FROM%.py}" ]; then
            FOUND=1
            break
        fi
    done
    if [ $FOUND -eq 0 ]; then
        echo "ERROR: '$START_FROM' is not in the pipeline. Exiting."
        exit 1
    fi
    echo "Starting from: $START_FROM"
fi

# Check every script exists before running anything
MISSING=0
for s in "${SCRIPTS[@]}"; do
    if [ ! -f "$SCRIPT_DIR/$s.py" ]; then
        echo "ERROR: missing $SCRIPT_DIR/$s.py"
        MISSING=1
    fi
done
if [ $MISSING -eq 1 ]; then
    echo "One or more scripts are missing. Exiting before running anything."
    exit 1
fi

TOTAL=${#SCRIPTS[@]}
RUN_START=$(date +%s)
echo "Pipeline started at $(date)"
echo "Working directory: $SCRIPT_DIR"
echo "Logs: $LOG_DIR"
echo "----------------------------------------------------------------------"

INDEX=0
for s in "${SCRIPTS[@]}"; do
    INDEX=$((INDEX + 1))

    if [ $STARTED -eq 0 ]; then
        if [ "$s" == "${START_FROM%.py}" ]; then
            STARTED=1
        else
            echo "[$INDEX/$TOTAL] Skipping $s.py"
            continue
        fi
    fi

    LOG_FILE="$LOG_DIR/${s}.log"
    STEP_START=$(date +%s)
    echo "[$INDEX/$TOTAL] Running $s.py ..."

    "$PYTHON" "$SCRIPT_DIR/$s.py" 2>&1 | tee "$LOG_FILE"
    STATUS=${PIPESTATUS[0]}
    STEP_END=$(date +%s)

    if [ $STATUS -ne 0 ]; then
        echo "----------------------------------------------------------------------"
        echo "FAILED: $s.py exited with status $STATUS after $((STEP_END - STEP_START))s"
        echo "See $LOG_FILE for details."
        echo "After fixing it, resume with:  ./run_pipeline.sh $s"
        exit $STATUS
    fi

    echo "[$INDEX/$TOTAL] Finished $s.py in $((STEP_END - STEP_START))s"
    echo "----------------------------------------------------------------------"
done

RUN_END=$(date +%s)
echo "All $TOTAL scripts completed successfully in $((RUN_END - RUN_START))s"
echo "Pipeline finished at $(date)"