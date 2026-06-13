# Security Hardening — ms365-access

**Audience:** an agent launched in this repo to execute the fixes below
**Status:** actionable punch-list (ms365-access is already partly hardened — do NOT re-do what already exists)
**Date:** 2026-06-13

---

## Baseline — what ALREADY exists (do not touch these)

| Item | State | Evidence |
|------|-------|----------|
| `backend/requirements.lock` | Hash-pinned (904 `--hash=` entries) | `uv pip compile ... --generate-hashes` |
| `backend/requirements-mcp.lock` | Hash-pinned (433 `--hash=` entries) | Same |
| `Dockerfile` | `--require-hashes` enforced | `pip install --no-cache-dir --require-hashes -r requirements.lock` |
| `Dockerfile.mcp` | `--require-hashes` enforced | Same pattern |
| `.github/workflows/security.yml` | CVE scanning via `pip-audit` | Runs against both lock files on push/PR/weekly |

**Gaps to fill (in order):**
1. No `.github/dependabot.yml` at all
2. No `.pre-commit-config.yaml`
3. No seasoning/cooldown layer (install-time or PR-time)
4. `security.yml` is missing `gitleaks` and uses `pip install pip-audit` rather than the pinned action
5. Docker base images not pinned to a digest (optional — noted at end)

---

## Item 1 — Add `.github/dependabot.yml`

**Why:** Without Dependabot, known-vulnerable deps get no automated PR. Also keeps the GitHub Actions themselves patched (Actions are a supply-chain surface).

**File to create:** `.github/dependabot.yml`

```yaml
# Dependabot keeps dependencies patched and flags known-vulnerable versions.
# Free on private repos. PRs land for you to review; read the changelog before merging,
# then regenerate the lock:
#   uv pip compile backend/requirements.txt --generate-hashes --python-version 3.11 -o backend/requirements.lock
#   uv pip compile backend/requirements-mcp.txt --generate-hashes --python-version 3.11 -o backend/requirements-mcp.lock
version: 2

updates:
  # Python API deps — Dependabot reads requirements.txt (the source ranges).
  - package-ecosystem: "pip"
    directory: "/backend"
    schedule:
      interval: "weekly"
      day: "monday"
    open-pull-requests-limit: 10
    labels:
      - "dependencies"
    # SEASONING: hold non-security version bumps for 30 days after release.
    # Security/CVE advisories still land immediately regardless of this setting.
    # Reference: https://github.blog/changelog/2025-07-01-dependabot-supports-configuration-of-a-minimum-package-age/
    cooldown:
      default-days: 30
      semver-patch-days: 7

  # Keep the GitHub Actions used in CI patched.
  - package-ecosystem: "github-actions"
    directory: "/"
    schedule:
      interval: "weekly"
      day: "monday"
    labels:
      - "dependencies"
      - "ci"
    cooldown:
      default-days: 14
```

**Note on cooldown semantics:** `cooldown` governs Dependabot PRs only — it delays PRs for packages younger than N days after their release date. It does NOT apply to security/CVE-advisory-triggered updates (those land immediately) and does NOT gate manual `pip install` or `uv` runs. The next item (Item 3b) closes that gap.

---

## Item 2 — Add `.pre-commit-config.yaml`

**Why:** Blocks secrets before they ever reach a commit — the last local gate before push protection.

**File to create:** `.pre-commit-config.yaml` at repo root.

```yaml
# Local pre-commit gates — block secrets before they ever reach a commit.
# One-time setup (run from repo root):
#   pip install pre-commit        (or: brew install pre-commit)
#   pre-commit autoupdate         # corrects rev: pins to current valid release tags
#   pre-commit install            # installs the git hook
#
# `autoupdate` is important — run it before the first install so hook revs are current.
repos:
  - repo: https://github.com/gitleaks/gitleaks
    rev: v8.30.0
    hooks:
      - id: gitleaks

  - repo: https://github.com/pre-commit/pre-commit-hooks
    rev: v6.0.0
    hooks:
      - id: detect-private-key            # catches PEM/SSH keys
      - id: check-added-large-files       # a 5 MB blob is often an accidental dump
        args: ["--maxkb=1024"]
      - id: end-of-file-fixer
      - id: trailing-whitespace
```

**After creating the file, run:**
```bash
pre-commit autoupdate        # corrects rev: pins to their current valid release tags
pre-commit install           # installs the git hook into .git/hooks/pre-commit
pre-commit run --all-files   # smoke-test on the existing tree (expect clean)
```

---

## Item 3 — Seasoning / cooldown layer (BOTH sub-items required)

Dependabot cooldown (Item 1) gates PRs only. A fresh `uv pip install` or a venv rebuild bypasses it entirely. You need an install-time gate too.

### 3a — Dependabot cooldown (already included in Item 1)

See `cooldown:` block in the dependabot.yml above. No extra action needed here.

### 3b — Install-time `--exclude-newer` gate via `uv.toml`

**Why:** `uv --exclude-newer` tells the resolver to pretend packages uploaded after a given date don't exist, so a fresh venv rebuild or `uv pip sync` never silently pulls a package younger than the window — covering manual/agent installs that Dependabot's cooldown does not.

> ⚠️ **uv version caveat:** uv 0.10.x (`uv --version` here: 0.10.4) accepts only a **fixed date** for `exclude-newer`, **not** a `"30 days"` duration (newer-uv only). The config value is a static *floor* you bump monthly; the true rolling window comes from passing `--exclude-newer "$(date -v-30d +%F)"` on the compile command.

**Reference:** https://docs.astral.sh/uv/concepts/resolution/

**File to create:** `uv.toml` at repo root (top-level key — the `[tool.uv]` header is for `pyproject.toml`, NOT `uv.toml`).

```toml
# Seasoning: refuse to resolve packages uploaded AFTER this date (safety floor
# for ad-hoc `uv pip install`). FIXED DATE only on uv 0.10.x — bump ~monthly.
# The real rolling window is set on the compile command below.
exclude-newer = "2026-05-14"
```

**After creating `uv.toml`, regenerate both locks under a rolling 30-day window:**
```bash
SEASON="$(date -v-30d +%F)"   # macOS; Linux: $(date -d '30 days ago' +%F)
uv pip compile backend/requirements.txt \
  --generate-hashes --python-version 3.11 --exclude-newer "$SEASON" \
  -o backend/requirements.lock

uv pip compile backend/requirements-mcp.txt \
  --generate-hashes --python-version 3.11 --exclude-newer "$SEASON" \
  -o backend/requirements-mcp.lock
```

If a package in `requirements.txt` is younger than 30 days, `uv` will error and tell you which one. Either pin an older version or explicitly override `exclude-newer` for that one package (or bump the window temporarily).

**npm side:** ms365-access has no `client/package.json` (the `client/` dir is Python-only). No npm action needed.

---

## Item 4 — Fix `.github/workflows/security.yml`

The existing `security.yml` has two gaps versus the MARVIN reference:

| Gap | Current | Should be |
|-----|---------|-----------|
| pip-audit install method | `run: pip install pip-audit` (floating version) | Use pinned `pypa/gh-action-pip-audit@v1.1.0` action |
| gitleaks | Missing entirely | Add `gitleaks/gitleaks-action@v2` job |
| `workflow_dispatch` | Missing | Add — allows manual trigger for ad-hoc CVE checks |
| `permissions` block | Missing | Add `permissions: contents: read` (least-privilege) |

**Replace** `.github/workflows/security.yml` with:

```yaml
name: security

# Supply-chain + secret gates. Runs on every push/PR and weekly (to catch
# CVEs newly disclosed against already-pinned deps).
on:
  push:
    branches: [main]
  pull_request:
  schedule:
    - cron: "0 6 * * 1"   # Mondays 06:00 UTC
  workflow_dispatch:

permissions:
  contents: read

jobs:
  pip-audit:
    name: pip-audit (known-vuln scan)
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: "3.11"
      # Audit API lock against Python advisory DB.
      - name: Audit API dependencies
        uses: pypa/gh-action-pip-audit@v1.1.0
        with:
          inputs: backend/requirements.lock
      # Audit MCP lock separately.
      - name: Audit MCP dependencies
        uses: pypa/gh-action-pip-audit@v1.1.0
        with:
          inputs: backend/requirements-mcp.lock

  gitleaks:
    name: gitleaks (secret scan)
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
        with:
          fetch-depth: 0   # full history, so a secret in any past commit is caught
      - uses: gitleaks/gitleaks-action@v2
        env:
          GITHUB_TOKEN: ${{ secrets.GITHUB_TOKEN }}
          # Free for personal repos (prdeepak/*).
          # Org-owned repos need GITLEAKS_LICENSE.
```

**Why the action over `pip install pip-audit`:** The pinned action (`@v1.1.0`) bundles a known-good pip-audit version; `pip install pip-audit` floats to whatever PyPI serves at CI time, which is a supply-chain risk in itself.

**Note on ignoring future CVEs:** If `pip-audit` fails on a CVE you've assessed and accepted (like MARVIN's PYSEC-2026-161 starlette BadHost), add an `ignore-vulns` key:
```yaml
      - name: Audit API dependencies
        uses: pypa/gh-action-pip-audit@v1.1.0
        with:
          inputs: backend/requirements.lock
          ignore-vulns: |
            PYSEC-XXXX-YYY
```
Document the rationale in a comment inline, with date and reassessment trigger (e.g. "remove when FastAPI ≥ 0.136 co-upgrade lands").

---

## Item 5 — Enable GitHub secret scanning + push protection (repo setting)

**Not a file change — do this in the GitHub UI.**

GitHub renamed this page: it is now **Settings → (left sidebar, under "Security and quality") → "Advanced Security"** — the URL is still `https://github.com/prdeepak/ms365-access/settings/security_analysis`.

1. Open that page.
2. Scroll **below the Dependabot block** to the **"Secret scanning"** section and enable it (free for all repos as of 2024).
3. **Push protection** appears as a sub-toggle once secret scanning is on — enable it.

This blocks secrets at `git push` time even if the pre-commit hook was skipped (e.g. `--no-verify`, new machine, CI pushes).

**Verify from the CLI:**
```bash
gh api repos/prdeepak/ms365-access --jq '.security_and_analysis'
# expect secret_scanning, secret_scanning_push_protection, dependabot_security_updates = "enabled"
```

> Status (2026-06-13): confirmed enabled — `secret_scanning`, `secret_scanning_push_protection`, `dependabot_security_updates` all `"enabled"`. The optional `secret_scanning_non_provider_patterns` (generic/custom token formats) and `secret_scanning_validity_checks` (probe found tokens for liveness) remain disabled — nice-to-haves, enable if desired.

---

## Item 6 — Require the security checks before merge (branch protection)

**Not a file change — do this in the GitHub UI / `gh`.** Without this, the `security` workflow runs but a red PR can still merge to `main`. This turns the CVE + secret scan from advisory into an enforced merge gate.

**Settings → Branches** (classic branch protection) or **Settings → Rules → Rulesets** (newer) for `main`:

1. ✅ **Require status checks to pass before merging**
2. ✅ **Require branches to be up to date before merging** (`strict` — checks run against the post-merge state)
3. In the check search box, select **`pip-audit (known-vuln scan)`** and **`gitleaks (secret scan)`**.
4. Optionally ✅ **Require a pull request before merging** — note this blocks direct pushes to `main`; everything must go through a PR.

> ⚠️ **Timing gotcha:** the individual check names only appear in the selection box **after the workflow has reported at least once** on a branch/PR. If you enable "require status checks" before that, the required-checks list is empty and nothing is actually gated. Let the workflow run once (e.g. on the implementing PR), then add the checks.

**Add the checks from the CLI (additive — won't disturb other protection settings):**
```bash
gh api -X POST repos/prdeepak/ms365-access/branches/main/protection/required_status_checks/contexts \
  -f "contexts[]=pip-audit (known-vuln scan)" \
  -f "contexts[]=gitleaks (secret scan)"
# verify:
gh api repos/prdeepak/ms365-access/branches/main/protection/required_status_checks --jq '{strict,contexts}'
```

> Status (2026-06-13): enforced — `main` requires both `pip-audit (known-vuln scan)` and `gitleaks (secret scan)`, `strict: true`.

---

## Item 7 — Pin Docker base images to digests ✅ DONE (2026-06-13)

**Was:** both Dockerfiles used `python:3.11-slim` (floating tag — a new image can be pulled silently).

**Now:** both `Dockerfile` and `Dockerfile.mcp` pin the multi-arch (OCI image index) digest:
```dockerfile
FROM python:3.11-slim@sha256:f9fa7f851e38bfb19c9de3afbc4b86ae7176ea7aaf94535c31df5458d5849457
```
This makes `docker pull` deterministic and tamper-evident. The digest is the manifest-list digest, so cross-platform builds (amd64/arm64/…) still resolve correctly.

**Re-pin when** you intentionally upgrade Python or want base-image security patches (the pinned tag will NOT float them in for you):
```bash
docker buildx imagetools inspect python:3.11-slim --format '{{.Manifest.Digest}}'
# paste the new sha256 into both Dockerfiles
```

---

## Execution order for the agent

Run these steps **in order**; each is independently verifiable:

1. `mkdir -p .github` (already exists) and write `.github/dependabot.yml` (Item 1)
2. Write `.pre-commit-config.yaml` at repo root (Item 2), then `pre-commit autoupdate && pre-commit install && pre-commit run --all-files`
3. Write `uv.toml` at repo root (Item 3b), then regenerate both lock files and verify they compile cleanly
4. Replace `.github/workflows/security.yml` with the fixed version (Item 4)
5. Commit all four file changes in a single commit: `feat(security): add dependabot, pre-commit, uv seasoning, fix security.yml`
6. Push the commit
7. (Manual) Enable secret scanning + push protection on the **Advanced Security** settings page (Item 5)
8. (Manual) After the `security` workflow runs once, require `pip-audit` + `gitleaks` as status checks on `main` (Item 6)

**Verify after push:**
- GitHub Actions → security workflow should pass green on the new commit
- `pre-commit run --all-files` should be clean locally
- Dependabot should appear under Insights → Dependency graph → Dependabot within ~10 minutes of the dependabot.yml push
- `gh api repos/prdeepak/ms365-access --jq '.security_and_analysis'` shows secret scanning + push protection enabled (Item 5)
- `gh api repos/prdeepak/ms365-access/branches/main/protection/required_status_checks --jq '{strict,contexts}'` lists both checks (Item 6)
