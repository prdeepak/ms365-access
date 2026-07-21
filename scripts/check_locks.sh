#!/usr/bin/env bash
#
# Verify each hash-pinned lock is in sync with the requirements .txt it compiles from.
#
# Why this exists: Dependabot edits the .txt files but knows nothing about the
# uv-generated .lock files. Every image and CI job installs from the locks
# (--require-hashes), so a .txt-only bump is inert AND leaves the two files
# disagreeing. That shipped as #65 and #69 before this check existed.
#
# How it works — the important part: `uv pip compile` treats versions already
# pinned in its OUTPUT file as preferences, changing only what the .txt (or a
# moved uv.toml exclude-newer floor) forces. So we copy the committed lock to a
# temp file and recompile ONTO that copy. A package that the .txt moved shows up
# as a diff; everything else keeps its pin.
#
# Recompiling to an empty path instead would resolve from scratch and pick up
# every newer release allowed by the floor — reporting failure constantly and
# for the wrong reason. Do not "simplify" this by dropping the copy.
#
# Usage:  scripts/check_locks.sh          (also: make check-locks)
# Exits 0 when every lock is in sync, 1 otherwise.

set -euo pipefail

cd "$(dirname "$0")/.."   # repo root, so uv discovers uv.toml (exclude-newer floor)

if ! command -v uv >/dev/null 2>&1; then
    echo "error: uv is not installed (needed to recompile the locks)" >&2
    exit 2
fi

PYTHON_VERSION=3.11
failed=0

# Strip only the top-level header comments (which record the -o path and would
# always differ for a temp file). Indented "# via ..." annotations start with
# whitespace, so they survive and are still compared.
strip_header() { grep -v '^#' "$1"; }

check_lock() {
    local txt=$1 lock=$2
    shift 2
    local extra=("$@")

    printf '  %-32s ' "$(basename "$lock")"

    local tmp
    tmp=$(mktemp)
    cp "$lock" "$tmp"          # seed with committed pins so they are preferred

    if ! uv pip compile "$txt" \
            --generate-hashes \
            --python-version "$PYTHON_VERSION" \
            "${extra[@]}" \
            -o "$tmp" --quiet 2>/dev/null; then
        echo "ERROR (uv pip compile failed)"
        rm -f "$tmp"
        failed=1
        return
    fi

    if diff -q <(strip_header "$lock") <(strip_header "$tmp") >/dev/null; then
        echo "in sync"
    else
        echo "OUT OF SYNC"
        echo
        echo "  $lock does not match $txt. Package-level differences"
        echo "  (committed lock '<' vs freshly compiled '>'):"
        # diff prefixes each line with "< " / "> " — the space matters here.
        diff <(strip_header "$lock") <(strip_header "$tmp") \
            | grep -E '^[<>] [a-zA-Z0-9._-]+==' | sed 's/^/    /' || true
        local cmd="uv pip compile $txt --generate-hashes --python-version $PYTHON_VERSION"
        if [ "${#extra[@]}" -gt 0 ]; then
            cmd="$cmd ${extra[*]}"
        fi
        echo
        echo "  Regenerate with:"
        echo "    $cmd -o $lock"
        echo
        failed=1
    fi

    rm -f "$tmp"
}

echo "Checking locks are in sync with their requirements files..."
check_lock backend/requirements.txt      backend/requirements.lock
check_lock backend/requirements-mcp.txt  backend/requirements-mcp.lock
check_lock backend/requirements-dev.txt  backend/requirements-dev.lock -c backend/requirements.lock

if [ "$failed" -ne 0 ]; then
    echo "Lock check FAILED — regenerate the lock(s) above and commit the result."
    exit 1
fi

echo "All locks in sync."
