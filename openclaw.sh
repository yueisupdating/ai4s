#!/bin/bash
set -euo pipefail

INSTITUTE_NAMES=(
  "中科院动物所"
  "中科院植物所"
  "中科院上海药物所"
  "中科院地质与地球物理所"
  "中科院大气物理所"
  "中科院半导体所"
  "中科院微电子所"
  "中科院数学与系统科学研究院"
  "中科院物理所"
  "中科院化学研究所"
  "中科院国家空间科学中心"
)

SOURCE_QUESTIONNAIRE_DOCX="院内科研单位AI4S“就绪度”调查问卷.docx"
WORKDIR="/home/updating/.openclaw/workspace/ai4sci"
RAW_DOCX="问卷.docx"
FORMATTED_DOCX="问卷_格式整理.docx"
SINGLE_CHOICE_JSON="单选题答案.json"
RESULT_DIR="result"
OPENCLAW_AGENT_TIMEOUT_SECONDS="${OPENCLAW_AGENT_TIMEOUT_SECONDS:-21600}"

HELPER_DIR=".openclaw_helpers"
DOCX_FORMATTER_SOURCE="$HELPER_DIR/fix_docx_format.py"
VALIDATOR_SCRIPT="$HELPER_DIR/validate_outputs.py"
SEARCH_GUIDANCE_TEMPLATE="$HELPER_DIR/search_guidance.md"
ARXIV_RESEARCH_PROMPT_TEMPLATE="$HELPER_DIR/arxiv_research_prompt.md"
BUILD_QUESTIONNAIRE_PROMPT_TEMPLATE="$HELPER_DIR/build_questionnaire_prompt.md"
FINALIZE_QUESTIONNAIRE_PROMPT_TEMPLATE="$HELPER_DIR/finalize_questionnaire_prompt.md"

render_template() {
  local path="$1"
  local template
  local current_dir

  current_dir="$(pwd)"
  template="$(<"$path")"
  template="${template//\{\{INSTITUTE_NAME\}\}/$INSTITUTE_NAME}"
  template="${template//\{\{SOURCE_QUESTIONNAIRE_DOCX\}\}/$SOURCE_QUESTIONNAIRE_DOCX}"
  template="${template//\{\{RAW_DOCX\}\}/$RAW_DOCX}"
  template="${template//\{\{SINGLE_CHOICE_JSON\}\}/$SINGLE_CHOICE_JSON}"
  template="${template//\{\{CURRENT_DIR\}\}/$current_dir}"
  printf '%s' "$template"
}

wait_for_file() {
  local path="$1"
  local timeout_seconds="${2:-3600}"
  local elapsed=0

  while [[ ! -s "$path" ]]; do
    if ((elapsed >= timeout_seconds)); then
      echo "ERROR: waited ${timeout_seconds}s but ${path} was not created." >&2
      exit 1
    fi
    sleep 5
    elapsed=$((elapsed + 5))
  done
}

find_named_output_file() {
  local name
  local path

  for name in "$@"; do
    if [[ -s "$name" ]]; then
      printf '%s\n' "$name"
      return
    fi
  done

  shopt -s nullglob
  for name in "$@"; do
    for path in ./*/"$name"; do
      if [[ -s "$path" ]]; then
        printf '%s\n' "$path"
        shopt -u nullglob
        return
      fi
    done
  done
  shopt -u nullglob
}

wait_for_named_output_file() {
  local timeout_seconds="$1"
  shift
  local elapsed=0
  local path

  while true; do
    path="$(find_named_output_file "$@")"
    if [[ -n "$path" ]]; then
      printf '%s\n' "$path"
      return
    fi

    if ((elapsed >= timeout_seconds)); then
      echo "ERROR: waited ${timeout_seconds}s but none of these files were created in . or one-level subdirectories: $*" >&2
      exit 1
    fi
    sleep 5
    elapsed=$((elapsed + 5))
  done
}

cleanup_workspace() {
  local keep_names=(
    "fix_docx_format.py"
    "README.md"
    "$SOURCE_QUESTIONNAIRE_DOCX"
    "openclaw.sh"
    "logs"
    "skills"
    "$RESULT_DIR"
  )
  local path name keep_name keep

  shopt -s nullglob
  for path in ./*; do
    name="${path#./}"
    keep=0
    for keep_name in "${keep_names[@]}"; do
      if [[ "$name" == "$keep_name" ]]; then
        keep=1
        break
      fi
    done
    if [[ "$name" == openclaw.sh.bak-* ]]; then
      keep=1
    fi

    if ((keep == 0)); then
      echo "cleanup: removing $name"
      rm -rf -- "$path"
    fi
  done
  shopt -u nullglob
}

safe_name() {
  local name="$1"
  name="${name//\//_}"
  name="${name// /_}"
  printf '%s\n' "$name"
}

archive_failure_artifacts() {
  local institute_name="$1"
  local safe_institute timestamp archive_dir
  local artifacts=(
    "$RAW_DOCX"
    "$FORMATTED_DOCX"
    "$SINGLE_CHOICE_JSON"
    "问卷.md"
    "问卷v2.md"
    "问卷v2.docx"
    "lost.md"
    "plan.md"
    "arxiv.md"
  )
  local path

  safe_institute="$(safe_name "$institute_name")"
  timestamp="$(date +%Y%m%d-%H%M%S)"
  archive_dir="logs/failed-${safe_institute}-${timestamp}"
  mkdir -p "$archive_dir"

  shopt -s nullglob
  for path in "${artifacts[@]}"; do
    if [[ -e "$path" ]]; then
      cp -a -- "$path" "$archive_dir/"
    fi
  done
  shopt -u nullglob

  echo "==== ${institute_name}: archived failure artifacts to ${archive_dir} ====" >&2
}

validate_score_json() {
  python3 "$VALIDATOR_SCRIPT" score-json "$1"
}

validate_docx_question_coverage() {
  python3 "$VALIDATOR_SCRIPT" docx-coverage "$1"
}

validate_docx_open_question_dimensions() {
  python3 "$VALIDATOR_SCRIPT" docx-q47 "$1"
}

validate_docx_source_text() {
  python3 "$VALIDATOR_SCRIPT" docx-source-text "$SOURCE_QUESTIONNAIRE_DOCX" "$1"
}

restore_docx_formatter() {
  if [[ ! -s "$DOCX_FORMATTER_SOURCE" ]]; then
    echo "ERROR: missing canonical formatter: $DOCX_FORMATTER_SOURCE" >&2
    exit 1
  fi
  cp -f -- "$DOCX_FORMATTER_SOURCE" fix_docx_format.py
  chmod +x fix_docx_format.py
}

require_helper_files() {
  local required=(
    "$DOCX_FORMATTER_SOURCE"
    "$VALIDATOR_SCRIPT"
    "$SEARCH_GUIDANCE_TEMPLATE"
    "$ARXIV_RESEARCH_PROMPT_TEMPLATE"
    "$BUILD_QUESTIONNAIRE_PROMPT_TEMPLATE"
    "$FINALIZE_QUESTIONNAIRE_PROMPT_TEMPLATE"
  )
  local path

  for path in "${required[@]}"; do
    if [[ ! -s "$path" ]]; then
      echo "ERROR: missing required helper file: $path" >&2
      exit 1
    fi
  done
}

run_for_institute() {
  local INSTITUTE_NAME="$1"
  local SEARCH_GUIDANCE
  local safe_institute
  local output_path
  local json_output_path
  local docx_output_path
  local single_choice_output_path

  SEARCH_GUIDANCE="$(render_template "$SEARCH_GUIDANCE_TEMPLATE")"
  safe_institute="${INSTITUTE_NAME//\//_}"
  safe_institute="${safe_institute// /_}"
  output_path="$RESULT_DIR/${FORMATTED_DOCX%.docx}_${safe_institute}.docx"
  json_output_path="$RESULT_DIR/${SINGLE_CHOICE_JSON%.json}_${safe_institute}.json"

  if [[ -s "$output_path" && -s "$json_output_path" ]]; then
    if validate_docx_question_coverage "$output_path" && validate_docx_source_text "$output_path" && validate_docx_open_question_dimensions "$output_path" && validate_score_json "$json_output_path"; then
      echo "==== ${INSTITUTE_NAME}: already done; skipping ===="
      return
    fi
    echo "==== ${INSTITUTE_NAME}: existing output failed validation; regenerating ===="
    rm -f -- "$output_path" "$json_output_path"
  fi

  echo "==== ${INSTITUTE_NAME}: start ===="
  cleanup_workspace
  restore_docx_formatter
  mkdir -p "$RESULT_DIR"
  rm -f -- "$RAW_DOCX" "$FORMATTED_DOCX" "$SINGLE_CHOICE_JSON"

  openclaw agent \
    --agent main \
    --local \
    --timeout "$OPENCLAW_AGENT_TIMEOUT_SECONDS" \
    --message "$(printf '%s\n\n%s' "$SEARCH_GUIDANCE" "$(render_template "$ARXIV_RESEARCH_PROMPT_TEMPLATE")")"

  openclaw agent \
    --agent main \
    --local \
    --timeout "$OPENCLAW_AGENT_TIMEOUT_SECONDS" \
    --message "$(printf '%s\n\n%s' "$SEARCH_GUIDANCE" "$(render_template "$BUILD_QUESTIONNAIRE_PROMPT_TEMPLATE")")"

  openclaw agent \
    --agent main \
    --local \
    --timeout "$OPENCLAW_AGENT_TIMEOUT_SECONDS" \
    --message "$(printf '%s\n\n%s' "$SEARCH_GUIDANCE" "$(render_template "$FINALIZE_QUESTIONNAIRE_PROMPT_TEMPLATE")")"

  docx_output_path="$(wait_for_named_output_file 36000 "$RAW_DOCX" "${RAW_DOCX%.docx}v2.docx")"
  single_choice_output_path="$(wait_for_named_output_file 36000 "$SINGLE_CHOICE_JSON")"
  validate_score_json "$single_choice_output_path"
  restore_docx_formatter
  INSTITUTE_NAME="$INSTITUTE_NAME" python3 fix_docx_format.py "$docx_output_path" -o "$FORMATTED_DOCX"
  wait_for_file "$FORMATTED_DOCX" 6000
  validate_docx_question_coverage "$FORMATTED_DOCX"
  validate_docx_source_text "$FORMATTED_DOCX"
  validate_docx_open_question_dimensions "$FORMATTED_DOCX"
  mv -f -- "$FORMATTED_DOCX" "$output_path"
  mv -f -- "$single_choice_output_path" "$json_output_path"
  cleanup_workspace
  echo "==== ${INSTITUTE_NAME}: wrote ${output_path} and ${json_output_path} ===="
}

cd ~
source ai4sci/bin/activate
cd "$WORKDIR"
mkdir -p "$RESULT_DIR"
require_helper_files

if [[ "${1:-}" == "--one-institute" ]]; then
  if [[ $# -ne 2 ]]; then
    echo "ERROR: --one-institute requires exactly one institute name." >&2
    exit 2
  fi
  run_for_institute "$2"
  exit
fi

batch_failed=0
for institute_name in "${INSTITUTE_NAMES[@]}"; do
  if bash "$0" --one-institute "$institute_name"; then
    :
  else
    status=$?
    batch_failed=1
    archive_failure_artifacts "$institute_name"
    echo "==== ${institute_name}: failed with status ${status}; continuing to next institute ====" >&2
  fi
done

if ((batch_failed)); then
  echo "==== batch finished with one or more failed institutes; see log above ====" >&2
  exit 1
fi

echo "==== batch finished successfully ===="
