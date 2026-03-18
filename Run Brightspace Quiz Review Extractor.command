#!/bin/zsh

set -u

APP_NAME="Brightspace Quiz Review Extractor"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
EXTRACTOR_SCRIPT="$SCRIPT_DIR/brightspace_quiz_review_extractor_v2.py"
REQUIREMENTS_FILE="$SCRIPT_DIR/requirements.txt"
VENV_DIR="$SCRIPT_DIR/.venv"
VENV_PYTHON="$VENV_DIR/bin/python"

show_alert() {
  local message="$1"
  /usr/bin/osascript <<APPLESCRIPT
display alert "${APP_NAME}" message "${message}"
APPLESCRIPT
}

show_dialog() {
  local message="$1"
  /usr/bin/osascript <<APPLESCRIPT
display dialog "${message}" buttons {"OK"} default button "OK"
APPLESCRIPT
}

choose_input_mode() {
  /usr/bin/osascript <<'APPLESCRIPT'
button returned of (display dialog "Select the type of Brightspace export you want to review." buttons {"Cancel", "Folder", "ZIP"} default button "Folder")
APPLESCRIPT
}

choose_input_path() {
  local mode="$1"
  if [[ "$mode" == "ZIP" ]]; then
    /usr/bin/osascript <<'APPLESCRIPT'
POSIX path of (choose file with prompt "Select the Brightspace ZIP export." of type {"zip"})
APPLESCRIPT
  else
    /usr/bin/osascript <<'APPLESCRIPT'
POSIX path of (choose folder with prompt "Select the unpacked Brightspace export folder.")
APPLESCRIPT
  fi
}

choose_output_parent() {
  /usr/bin/osascript <<'APPLESCRIPT'
POSIX path of (choose folder with prompt "Select where the review output folder should be created.")
APPLESCRIPT
}

choose_image_mode() {
  /usr/bin/osascript <<'APPLESCRIPT'
button returned of (display dialog "How should question images be linked in the review output?" buttons {"Cancel", "Leave In Export", "Copy Images"} default button "Copy Images")
APPLESCRIPT
}

ensure_python3() {
  if command -v python3 >/dev/null 2>&1; then
    return 0
  fi
  show_alert "Python 3 was not found on this Mac. Install Python 3, then run this launcher again."
  return 1
}

bootstrap_venv() {
  if [[ -x "$VENV_PYTHON" ]]; then
    return 0
  fi

  ensure_python3 || return 1
  show_dialog "First run setup: this launcher will create a local .venv and install the required Python package. This may take a minute and needs internet access."

  echo "Creating local virtual environment in $VENV_DIR"
  if ! python3 -m venv "$VENV_DIR"; then
    show_alert "Could not create the local virtual environment."
    return 1
  fi

  echo "Installing runtime dependency from requirements.txt"
  if ! "$VENV_PYTHON" -m pip install -r "$REQUIREMENTS_FILE"; then
    show_alert "Could not install the required Python package. Check your internet connection and try again."
    return 1
  fi
  return 0
}

main() {
  if [[ ! -f "$EXTRACTOR_SCRIPT" || ! -f "$REQUIREMENTS_FILE" ]]; then
    show_alert "This launcher must stay in the same folder as brightspace_quiz_review_extractor_v2.py and requirements.txt."
    exit 1
  fi

  local input_mode
  input_mode="$(choose_input_mode)" || exit 0

  local input_path
  input_path="$(choose_input_path "$input_mode")" || exit 0

  local output_parent
  output_parent="$(choose_output_parent)" || exit 0

  local image_mode
  image_mode="$(choose_image_mode)" || exit 0

  bootstrap_venv || exit 1

  local timestamp
  timestamp="$(date +"%Y%m%d_%H%M%S")"
  local output_dir="${output_parent%/}/quiz_review_out_${timestamp}"

  mkdir -p "$output_dir"

  echo "Running extractor..."
  echo "Input:  $input_path"
  echo "Output: $output_dir"

  local -a cmd
  cmd=("$VENV_PYTHON" "$EXTRACTOR_SCRIPT" "$input_path" --out "$output_dir")
  if [[ "$image_mode" == "Copy Images" ]]; then
    cmd+=(--copy-images-to-assets)
  fi

  if ! "${cmd[@]}"; then
    show_alert "The extractor did not complete successfully. Review the Terminal window for details."
    exit 1
  fi

  /usr/bin/open "$output_dir"
  show_alert "Review files were created successfully in:\n$output_dir"
}

main "$@"
