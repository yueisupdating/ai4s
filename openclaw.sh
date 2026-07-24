#!/bin/bash
set -euo pipefail

INSTITUTE_NAMES=(
  "金属研究所"
  "东北地理与农业生态研究所"
  "上海技术物理研究所"
  "分子细胞科学卓越创新中心"
  "上海营养与健康研究所"
  "福建物质结构研究所"
  "南京地质古生物研究所"
  "紫金山天文台"
  "苏州生物医学工程技术研究所"
  "水生生物研究所"
  "南海海洋研究所"
  "广州地球化学研究所"
  "东莞材料科学与技术研究所"
  "成都生物研究所"
  "重庆绿色智能技术研究院"
  "西双版纳热带植物园"
  "国家授时中心"
  "近代物理研究所"
  "青海盐湖研究所"
  "新疆生态与地理研究所"
  "古脊椎动物与古人类研究所"
  "生物物理研究所"
  "北京基因组研究所（国家生物信息中心）"
  "自动化研究所"
  "科技战略咨询研究院"
  "空间应用工程与技术中心"
)

SOURCE_QUESTIONNAIRE_DOCX="院内科研单位AI4S“就绪度”调查问卷.docx"
WORKDIR="/home/updating/.openclaw/workspace/ai4sci"
RAW_DOCX="问卷.docx"
FORMATTED_DOCX="问卷_格式整理.docx"
SINGLE_CHOICE_JSON="单选题答案.json"
RESULT_DIR="result"
SCORE_TABLE_DIR="${SCORE_TABLE_DIR:-score_tables}"
QUESTIONNAIRE_TEMPLATE="${QUESTIONNAIRE_TEMPLATE:-问卷图表设计20260715.xlsx}"
SUMMARY_OUTPUT="${SUMMARY_OUTPUT:-问卷图表设计20260715_汇总.xlsx}"
OPENCLAW_AGENT_TIMEOUT_SECONDS="${OPENCLAW_AGENT_TIMEOUT_SECONDS:-21600}"

HELPER_DIR=".openclaw_helpers"
DOCX_FORMATTER_SOURCE="$HELPER_DIR/fix_docx_format.py"
VALIDATOR_SCRIPT="$HELPER_DIR/validate_outputs.py"
SCORE_TABLE_EXPORT_SCRIPT="$HELPER_DIR/export_result_scores.py"
DOCX_FIELDS_EXPORT_SCRIPT="$HELPER_DIR/add_docx_fields_to_score_tables.py"
QUESTIONNAIRE_SUMMARY_SCRIPT="$HELPER_DIR/build_questionnaire_summary.py"
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
    "README.md"
    "$SOURCE_QUESTIONNAIRE_DOCX"
    "openclaw.sh"
    "assert"
    "logs"
    "skills"
    "$RESULT_DIR"
    "$SCORE_TABLE_DIR"
    "$QUESTIONNAIRE_TEMPLATE"
    "$SUMMARY_OUTPUT"
    "requirements.txt"
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

result_docx_path_for() {
  local safe_institute
  safe_institute="$(safe_name "$1")"
  printf '%s/%s_%s.docx\n' "$RESULT_DIR" "${FORMATTED_DOCX%.docx}" "$safe_institute"
}

result_json_path_for() {
  local safe_institute
  safe_institute="$(safe_name "$1")"
  printf '%s/%s_%s.json\n' "$RESULT_DIR" "${SINGLE_CHOICE_JSON%.json}" "$safe_institute"
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

validate_result_for_institute() {
  local institute_name="$1"
  local output_path json_output_path

  output_path="$(result_docx_path_for "$institute_name")"
  json_output_path="$(result_json_path_for "$institute_name")"
  if [[ ! -s "$output_path" || ! -s "$json_output_path" ]]; then
    echo "ERROR: missing result files for ${institute_name}: ${output_path}, ${json_output_path}" >&2
    return 1
  fi

  validate_docx_question_coverage "$output_path" &&
    validate_docx_source_text "$output_path" &&
    validate_docx_open_question_dimensions "$output_path" &&
    validate_score_json "$json_output_path"
}

all_expected_results_valid() {
  local institute_name

  for institute_name in "${INSTITUTE_NAMES[@]}"; do
    validate_result_for_institute "$institute_name" || return 1
  done
}

require_helper_files() {
  local required=(
    "$DOCX_FORMATTER_SOURCE"
    "$VALIDATOR_SCRIPT"
    "$SCORE_TABLE_EXPORT_SCRIPT"
    "$DOCX_FIELDS_EXPORT_SCRIPT"
    "$QUESTIONNAIRE_SUMMARY_SCRIPT"
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

export_score_tables() {
  if [[ ! -s "$SCORE_TABLE_EXPORT_SCRIPT" ]]; then
    echo "ERROR: missing score table export script: $SCORE_TABLE_EXPORT_SCRIPT" >&2
    exit 1
  fi

  echo "==== exporting score tables to ${SCORE_TABLE_DIR} ===="
  rm -rf -- "$SCORE_TABLE_DIR"
  python3 "$SCORE_TABLE_EXPORT_SCRIPT" \
    --input-dir "$RESULT_DIR" \
    --output-dir "$SCORE_TABLE_DIR" \
    --format xlsx

  echo "==== adding DOCX/JSON fields to ${SCORE_TABLE_DIR} ===="
  python3 "$DOCX_FIELDS_EXPORT_SCRIPT" \
    --docx-dir "$RESULT_DIR" \
    --json-dir "$RESULT_DIR" \
    --xlsx-dir "$SCORE_TABLE_DIR"
}

build_total_summary() {
  if [[ ! -s "$QUESTIONNAIRE_TEMPLATE" ]]; then
    echo "ERROR: missing questionnaire summary template: $QUESTIONNAIRE_TEMPLATE" >&2
    exit 1
  fi
  if [[ ! -d "$SCORE_TABLE_DIR" ]]; then
    echo "ERROR: missing score table directory: $SCORE_TABLE_DIR" >&2
    exit 1
  fi

  echo "==== building questionnaire summary workbook ${SUMMARY_OUTPUT} ===="
  python3 "$QUESTIONNAIRE_SUMMARY_SCRIPT" \
    --score-dir "$SCORE_TABLE_DIR" \
    --template "$QUESTIONNAIRE_TEMPLATE" \
    --output "$SUMMARY_OUTPUT"
}

run_for_institute() {
  local INSTITUTE_NAME="$1"
  local SEARCH_GUIDANCE
  local output_path
  local json_output_path
  local docx_output_path
  local single_choice_output_path

  SEARCH_GUIDANCE="$(render_template "$SEARCH_GUIDANCE_TEMPLATE")"
  output_path="$(result_docx_path_for "$INSTITUTE_NAME")"
  json_output_path="$(result_json_path_for "$INSTITUTE_NAME")"

  if [[ -s "$output_path" && -s "$json_output_path" ]]; then
    if validate_result_for_institute "$INSTITUTE_NAME"; then
      echo "==== ${INSTITUTE_NAME}: already done; skipping ===="
      return
    fi
    echo "==== ${INSTITUTE_NAME}: existing output failed validation; regenerating ===="
    rm -f -- "$output_path" "$json_output_path"
  fi

  echo "==== ${INSTITUTE_NAME}: start ===="
  cleanup_workspace
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
  INSTITUTE_NAME="$INSTITUTE_NAME" python3 "$DOCX_FORMATTER_SOURCE" "$docx_output_path" -o "$FORMATTED_DOCX"
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

if [[ "${1:-}" == "--summary-only" ]]; then
  build_total_summary
  exit
fi

if [[ "${1:-}" == "--export-summary-only" ]]; then
  export_score_tables
  build_total_summary
  exit
fi

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
  if all_expected_results_valid; then
    echo "==== batch had failed steps, but all expected result files validate; exporting score tables ===="
    export_score_tables
    build_total_summary
    echo "==== batch finished successfully ===="
    exit
  fi
  echo "==== batch finished with one or more failed institutes; see log above ====" >&2
  exit 1
fi

export_score_tables
build_total_summary
echo "==== batch finished successfully ===="
