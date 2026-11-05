#!/usr/bin/env bash
#
# pages2docx.sh - Recursively convert Apple Pages (.pages) documents to Word (.docx)
#
# Requires macOS with Apple Pages installed (bundle id com.apple.Pages).
# Conversion is done by Pages itself via AppleScript, so the exported .docx
# preserves the document's layout as well as Pages can.
#
# Notes:
#   * Works with modern single-file .pages documents and older package bundles.
#   * Each document is opened through LaunchServices (the same mechanism as
#     double-clicking it in Finder), exported, then closed again.
#   * Documents you already have open in Pages are left untouched.
#   * Dismiss any Pages dialogs before running the script.
#
# Usage: ./pages2docx.sh [options] <directory>

set -uo pipefail

SCRIPT_NAME=${0##*/}
PAGES_BUNDLE_ID="com.apple.Pages"

usage() {
  cat <<EOF
Usage: $SCRIPT_NAME [options] <directory>

Recursively finds every .pages document under <directory> and exports it as
.docx using the installed Apple Pages application.

Options:
  -o, --output DIR   Write .docx files to DIR, mirroring the source folder
                     structure. Default: write each .docx next to its .pages file.
  -f, --force        Overwrite existing .docx files (default: skip them).
  -n, --dry-run      Show what would be converted without converting anything.
  -q, --quiet        Only print errors and the final summary.
  -h, --help         Show this help.

Examples:
  $SCRIPT_NAME .
  $SCRIPT_NAME -o converted "My Documents"
  $SCRIPT_NAME --force ~/Projects
EOF
}

die() {
  printf 'error: %s\n' "$*" >&2
  exit 1
}

# Find the Pages application without launching it.
find_pages_app() {
  local candidate
  for candidate in /Applications/Pages*.app "$HOME"/Applications/Pages*.app /System/Applications/Pages*.app; do
    [ -d "$candidate" ] || continue
    if [ "$(/usr/libexec/PlistBuddy -c 'Print :CFBundleIdentifier' \
          "$candidate/Contents/Info.plist" 2>/dev/null)" = "$PAGES_BUNDLE_ID" ]; then
      printf '%s\n' "$candidate"
      return 0
    fi
  done

  # Fallback for non-standard install locations (Spotlight lookup).
  candidate=$(mdfind "kMDItemCFBundleIdentifier == '$PAGES_BUNDLE_ID'" 2>/dev/null | head -n 1)
  [ -n "$candidate" ] && printf '%s\n' "$candidate"
}

# Export one document.
#
# AppleScript's `open` command is unreliable for launching/opening documents
# (it can return `missing value` or fail with -1700), so the file is opened
# with `open -a` instead. Before opening, remember which documents are already
# in Pages; afterwards, export the one whose ID is new. This avoids touching
# documents the user may already have open, even if they have the same name.
convert_one() {
  local src=$1
  local dst=$2
  local before_ids
  local tmp_docx
  local status

  before_ids=$(osascript -e 'tell application "Pages" to get id of every document' 2>/dev/null | tr -d ' \n')

  if ! open -a "$app_path" "$src" >/dev/null 2>&1; then
    printf 'could not ask Pages to open the document\n' >&2
    return 1
  fi

  # Hidden sibling file, still ending in .docx so Pages keeps the extension.
  tmp_docx="$(dirname "$dst")/.pages2docx.$$.$3.docx"

  osascript - "$before_ids" "$tmp_docx" <<'APPLESCRIPT'
on run argv
  set beforeText to item 1 of argv
  set dstPath to item 2 of argv

  set beforeIDs to {}
  if beforeText is not "" then
    set oldDelims to AppleScript's text item delimiters
    set AppleScript's text item delimiters to ","
    set beforeIDs to text items of beforeText
    set AppleScript's text item delimiters to oldDelims
  end if

  tell application "Pages"
    set deadline to (current date) + 120
    set theDoc to missing value

    -- Wait for Pages to finish opening the new document.
    repeat
      repeat with d in documents
        set did to (id of d) as text
        if beforeIDs does not contain did then
          set theDoc to d
          exit repeat
        end if
      end repeat
      if theDoc is not missing value then exit repeat
      if (current date) > deadline then
        error "document did not open in Pages (is it already open? close it and retry)"
      end if
      delay 0.5
    end repeat

    try
      with timeout of 600 seconds
        export theDoc to (POSIX file dstPath) as Microsoft Word
      end timeout
    on error errMsg number errNum
      try
        close theDoc saving no
      end try
      error errMsg number errNum
    end try

    close theDoc saving no
  end tell
end run
APPLESCRIPT

  status=$?
  if [ "$status" -eq 0 ] && [ -s "$tmp_docx" ] && mv -f "$tmp_docx" "$dst"; then
    return 0
  fi

  rm -f "$tmp_docx"
  return 1
}

# --- Parse arguments ---------------------------------------------------------

output_dir=""
force=0
dry_run=0
quiet=0

while [ $# -gt 0 ]; do
  case "$1" in
    -o|--output)
      [ $# -ge 2 ] || die "option $1 requires a directory argument"
      output_dir=$2
      shift 2
      ;;
    -f|--force)   force=1; shift ;;
    -n|--dry-run) dry_run=1; shift ;;
    -q|--quiet)   quiet=1; shift ;;
    -h|--help)    usage; exit 0 ;;
    --)           shift; break ;;
    -*)           die "unknown option: $1 (try --help)" ;;
    *)            break ;;
  esac
done

[ $# -le 1 ] || die "too many arguments (try --help)"

root=${1:-.}
[ -d "$root" ] || die "not a directory: $root"
root=$(cd "$root" && pwd) || die "cannot access directory: $root"

if [ -n "$output_dir" ]; then
  mkdir -p "$output_dir" || die "cannot create output directory: $output_dir"
  output_dir=$(cd "$output_dir" && pwd) || die "cannot access output directory: $output_dir"
fi

app_path=$(find_pages_app)
if [ -z "$app_path" ]; then
  die "Apple Pages was not found (bundle id $PAGES_BUNDLE_ID). Install Pages from the App Store."
fi

if [ "$dry_run" -eq 0 ] && [ "$quiet" -eq 0 ]; then
  printf 'Using %s\n' "$app_path"
fi

# --- Convert -----------------------------------------------------------------

log() {
  [ "$quiet" -eq 1 ] || printf '%s\n' "$*"
}

total=0
converted=0
skipped=0
failed=0

while IFS= read -r -d '' src; do
  base=${src##*/}

  # Ignore hidden files/packages left behind by the OS or editors.
  case "$base" in
    .*) continue ;;
  esac

  total=$((total + 1))
  stem=${base%.[Pp][Aa][Gg][Ee][Ss]}

  if [ -n "$output_dir" ]; then
    rel=${src#"$root"/}
    [ "$rel" = "$src" ] && rel=$base
    rel_dir=$(dirname "$rel")
    if [ "$rel_dir" = "." ]; then
      dst_dir=$output_dir
    else
      dst_dir="$output_dir/$rel_dir"
    fi
    dst="$dst_dir/$stem.docx"
  else
    dst="${src%.[Pp][Aa][Gg][Ee][Ss]}.docx"
  fi

  if [ -e "$dst" ] && [ "$force" -eq 0 ]; then
    log "skip (exists): ${src#"$root"/}"
    skipped=$((skipped + 1))
    continue
  fi

  if [ "$dry_run" -eq 1 ]; then
    printf 'would convert: %s -> %s\n' "${src#"$root"/}" "$dst"
    converted=$((converted + 1))
    continue
  fi

  if [ -n "$output_dir" ]; then
    mkdir -p "$dst_dir" || {
      failed=$((failed + 1))
      printf 'FAILED (cannot create %s): %s\n' "$dst_dir" "$src" >&2
      continue
    }
  fi

  log "converting: ${src#"$root"/}"

  if err=$(convert_one "$src" "$dst" "$total" 2>&1); then
    converted=$((converted + 1))
  else
    failed=$((failed + 1))
    printf 'FAILED: %s\n' "$src" >&2
    [ -n "$err" ] && printf '        %s\n' "$err" >&2
    rm -f "$dst"
  fi
done < <(find "$root" -mindepth 1 -name '.*' -prune -o \( -type f -o -type d \) -iname '*.pages' -print0)

# --- Summary -----------------------------------------------------------------

printf '\nSummary: %d found, %d converted, %d skipped, %d failed\n' \
  "$total" "$converted" "$skipped" "$failed"

[ "$failed" -eq 0 ]
