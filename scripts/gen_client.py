#!/usr/bin/env python3
"""Generate Python client and README API docs from OpenAPI spec.

Reads the OpenAPI spec from a running FastAPI server and generates:
  1. client/<class_name>.py — Python HTTP client (stdlib only)
  2. Updates README.md API Reference section (between markers)

Usage:
    python scripts/gen_client.py              # generate from running server
    python scripts/gen_client.py --check      # verify files are up-to-date (for CI)
    python scripts/gen_client.py spec.json    # use saved spec file
"""

import json
import re
import sys
import textwrap
from pathlib import Path
from urllib.request import urlopen

# ── Project Config ─────────────────────────────────────────────────────────

CLASS_NAME = "Ms365Client"
MODULE_NAME = "ms365_client"
BASE_URL = "http://localhost:8365"
SPEC_URL = "http://localhost:8365/openapi.json"
CLIENT_FILE = "client/ms365_client.py"
HAS_ACCOUNT = False  # ms365 is single-account

# Endpoints returning raw bytes (substring match on path)
RAW_PATHS = ["/content"]

# Skip browser-flow endpoints (not useful in a programmatic client)
SKIP_PREFIXES = ["/auth/login", "/auth/callback"]

# Group display order and labels
GROUP_ORDER = [
    "health", "mail", "calendar", "files", "sharepoint",
]

MODULE_DOC = f'''\
"""Python client for ms365-access ({BASE_URL}).

Auto-generated from OpenAPI spec — do not edit manually.
Regenerate with: make gen-client

Zero dependencies beyond stdlib.

Usage:
    from ms365_client import Ms365Client
    client = Ms365Client()  # defaults to {BASE_URL}
"""
'''

# ── End Config ─────────────────────────────────────────────────────────────


def fetch_spec(source=None):
    """Fetch OpenAPI spec from file or URL."""
    if source and Path(source).exists():
        return json.loads(Path(source).read_text())
    url = source or SPEC_URL
    return json.loads(urlopen(url).read().decode())


def resolve_ref(spec, ref):
    """Resolve a JSON $ref pointer."""
    parts = ref.lstrip("#/").split("/")
    obj = spec
    for part in parts:
        obj = obj[part]
    return obj


def derive_method_name(op_id, path, method):
    """Extract original function name from FastAPI's auto-generated operationId.

    FastAPI generates: {func_name}{sanitized_path}_{method}
    where sanitized_path replaces all non-word chars with underscore.
    """
    sanitized_path = re.sub(r"\W", "_", path)
    suffix = f"{sanitized_path}_{method}"
    if op_id.endswith(suffix):
        name = op_id[: -len(suffix)]
        if name:
            return name
    # Fallback: strip just the method suffix
    if op_id.endswith(f"_{method}"):
        return op_id[: -(len(method) + 1)]
    return op_id


def is_raw_response(path):
    return any(raw in path for raw in RAW_PATHS)


def should_skip(path):
    return any(path.startswith(s) for s in SKIP_PREFIXES)


def get_py_default(schema):
    """Get Python default value string from OpenAPI schema."""
    if "default" not in schema:
        return None
    d = schema["default"]
    if isinstance(d, bool):
        return repr(d)
    if isinstance(d, str):
        return repr(d)
    if isinstance(d, (int, float)):
        return repr(d)
    if isinstance(d, list):
        return "None"  # avoid mutable defaults
    return repr(d)


def parse_endpoints(spec):
    """Parse OpenAPI paths into structured endpoint list."""
    endpoints = []
    for path, methods in spec["paths"].items():
        for method, detail in methods.items():
            if should_skip(path):
                continue
            op_id = detail.get("operationId", "")
            name = derive_method_name(op_id, path, method)
            summary = detail.get("summary", "")
            description = detail.get("description", "")

            path_params = []
            query_params = []
            for param in detail.get("parameters", []):
                info = {
                    "name": param["name"],
                    "in": param["in"],
                    "required": param.get("required", False),
                    "schema": param.get("schema", {}),
                }
                if info["in"] == "path":
                    path_params.append(info)
                elif info["in"] == "query":
                    query_params.append(info)

            body_fields = []
            has_request_body = "requestBody" in detail
            if has_request_body:
                rb = detail["requestBody"]
                content = rb.get("content", {})
                json_schema = content.get("application/json", {}).get("schema", {})
                if "$ref" in json_schema:
                    json_schema = resolve_ref(spec, json_schema["$ref"])
                required_fields = set(json_schema.get("required", []))
                for fname, fschema in json_schema.get("properties", {}).items():
                    if "$ref" in fschema:
                        fschema = resolve_ref(spec, fschema["$ref"])
                    body_fields.append({
                        "name": fname,
                        "required": fname in required_fields,
                        "schema": fschema,
                    })

            endpoints.append({
                "path": path,
                "method": method,
                "name": name,
                "summary": summary or description or "",
                "path_params": path_params,
                "query_params": query_params,
                "body_fields": body_fields,
                "has_request_body": has_request_body,
                "raw_response": is_raw_response(path),
            })

    endpoints.sort(key=lambda e: (e["path"], e["method"]))

    # Deduplicate method names by prefixing with group name
    name_counts = {}
    for ep in endpoints:
        name_counts[ep["name"]] = name_counts.get(ep["name"], 0) + 1
    for ep in endpoints:
        if name_counts[ep["name"]] > 1:
            prefix = ep["path"].strip("/").split("/")[0].replace("-", "_")
            ep["name"] = f"{prefix}_{ep['name']}"

    return endpoints


def group_by_prefix(endpoints):
    """Group endpoints by first path segment."""
    groups = {}
    for ep in endpoints:
        parts = ep["path"].strip("/").split("/")
        prefix = parts[0] if parts else "root"
        groups.setdefault(prefix, []).append(ep)
    # Sort groups by configured order
    ordered = {}
    for g in GROUP_ORDER:
        if g in groups:
            ordered[g] = groups.pop(g)
    for g in sorted(groups):
        ordered[g] = groups[g]
    return ordered


# ── Client code generation ─────────────────────────────────────────────────

def gen_method(ep):
    """Generate a single client method."""
    lines = []
    name = ep["name"]
    method = ep["method"]

    # --- Build parameter list ---
    args = ["self"]

    # Path params (positional, required)
    for p in ep["path_params"]:
        args.append(p["name"])

    # Request body fields: required first, then optional
    generic_body = ep.get("has_request_body", False) and not ep["body_fields"]
    req_body = [f for f in ep["body_fields"] if f["required"]]
    opt_body = [f for f in ep["body_fields"] if not f["required"]]
    for f in req_body:
        args.append(f["name"])
    for f in opt_body:
        d = get_py_default(f["schema"]) or "None"
        args.append(f"{f['name']}={d}")
    if generic_body:
        args.append("data=None")

    # Query params (skip account if HAS_ACCOUNT)
    qp = [p for p in ep["query_params"] if not (HAS_ACCOUNT and p["name"] == "account")]
    for p in sorted(qp, key=lambda p: not p["required"]):
        d = get_py_default(p["schema"])
        if p["required"] and d is None:
            args.append(p["name"])
        else:
            args.append(f"{p['name']}={d or 'None'}")

    # Account param at end
    has_acct = HAS_ACCOUNT and any(p["name"] == "account" for p in ep["query_params"])
    if has_acct:
        args.append("account=None")

    # --- Signature ---
    sig = ", ".join(args)
    max_line = 88
    if len(f"    def {name}({sig}):") <= max_line:
        lines.append(f"    def {name}({sig}):")
    else:
        lines.append(f"    def {name}(")
        for i, a in enumerate(args):
            comma = "," if i < len(args) - 1 else ","
            lines.append(f"            {a}{comma}")
        lines.append("    ):")

    # --- Docstring ---
    doc = ep["summary"] or name.replace("_", " ").capitalize()
    # Truncate long docstrings to one line
    doc = doc.split("\n")[0][:120]
    lines.append(f'        """{doc}"""')

    # --- Body: build data dict ---
    if ep["body_fields"]:
        lines.append("        data = {}")
        for f in ep["body_fields"]:
            if f["required"]:
                lines.append(f'        data["{f["name"]}"] = {f["name"]}')
            else:
                lines.append(f'        if {f["name"]} is not None:')
                lines.append(f'            data["{f["name"]}"] = {f["name"]}')

    # --- Body: build params ---
    if has_acct:
        if qp:
            kwargs = ", ".join(f"{p['name']}={p['name']}" for p in qp)
            params_expr = f"self._params(account, {kwargs})"
        else:
            params_expr = "self._params(account)"
    elif qp:
        items = ", ".join(f'"{p["name"]}": {p["name"]}' for p in qp)
        lines.append(
            f"        params = {{k: v for k, v in {{{items}}}.items() if v is not None}}"
        )
        params_expr = "params"
    else:
        params_expr = None

    # --- Body: HTTP call ---
    py_path = re.sub(r"\{(\w+)\}", r"{\1}", ep["path"])
    has_body = bool(ep["body_fields"]) or generic_body

    if method == "get":
        helper = "_get_raw" if ep["raw_response"] else "_get_json"
        if params_expr:
            lines.append(f'        return self.{helper}(f"{py_path}", {params_expr})')
        else:
            lines.append(f'        return self.{helper}(f"{py_path}")')
    elif method in ("post", "put", "patch"):
        helper = f"_{method}_json"
        if has_body and params_expr:
            lines.append(f'        return self.{helper}(f"{py_path}", data, {params_expr})')
        elif has_body:
            lines.append(f'        return self.{helper}(f"{py_path}", data)')
        elif params_expr:
            lines.append(f'        return self.{helper}(f"{py_path}", params={params_expr})')
        else:
            lines.append(f'        return self.{helper}(f"{py_path}")')
    elif method == "delete":
        if params_expr:
            lines.append(f'        return self._delete_json(f"{py_path}", {params_expr})')
        else:
            lines.append(f'        return self._delete_json(f"{py_path}")')

    return "\n".join(lines)


HTTP_HELPERS = '''\
    def _get_json(self, path, params=None, timeout=30):
        url = f"{self.base_url}{path}"
        if params:
            url += ("&" if "?" in path else "?") + urlencode(params)
        try:
            with urlopen(Request(url), timeout=timeout) as resp:
                return json.loads(resp.read().decode())
        except (URLError, HTTPError, json.JSONDecodeError) as e:
            log.warning(f"GET {path} failed: {e}")
            return None

    def _get_raw(self, path, params=None, timeout=30):
        url = f"{self.base_url}{path}"
        if params:
            url += ("&" if "?" in path else "?") + urlencode(params)
        try:
            with urlopen(Request(url), timeout=timeout) as resp:
                return resp.read()
        except (URLError, HTTPError) as e:
            log.warning(f"GET (raw) {path} failed: {e}")
            return None

    def _post_json(self, path, data=None, params=None, timeout=30):
        url = f"{self.base_url}{path}"
        if params:
            url += ("&" if "?" in path else "?") + urlencode(params)
        try:
            body = json.dumps(data).encode() if data else b""
            req = Request(url, data=body, method="POST")
            req.add_header("Content-Type", "application/json")
            with urlopen(req, timeout=timeout) as resp:
                return json.loads(resp.read().decode())
        except (URLError, HTTPError, json.JSONDecodeError) as e:
            log.warning(f"POST {path} failed: {e}")
            return None

    def _put_json(self, path, data=None, params=None, timeout=30):
        url = f"{self.base_url}{path}"
        if params:
            url += ("&" if "?" in path else "?") + urlencode(params)
        try:
            body = json.dumps(data).encode() if data else b""
            req = Request(url, data=body, method="PUT")
            req.add_header("Content-Type", "application/json")
            with urlopen(req, timeout=timeout) as resp:
                return json.loads(resp.read().decode())
        except (URLError, HTTPError, json.JSONDecodeError) as e:
            log.warning(f"PUT {path} failed: {e}")
            return None

    def _patch_json(self, path, data=None, params=None, timeout=30):
        url = f"{self.base_url}{path}"
        if params:
            url += ("&" if "?" in path else "?") + urlencode(params)
        try:
            body = json.dumps(data).encode() if data else b""
            req = Request(url, data=body, method="PATCH")
            req.add_header("Content-Type", "application/json")
            with urlopen(req, timeout=timeout) as resp:
                return json.loads(resp.read().decode())
        except (URLError, HTTPError, json.JSONDecodeError) as e:
            log.warning(f"PATCH {path} failed: {e}")
            return None

    def _delete_json(self, path, params=None, timeout=30):
        url = f"{self.base_url}{path}"
        if params:
            url += ("&" if "?" in path else "?") + urlencode(params)
        try:
            req = Request(url, method="DELETE")
            with urlopen(req, timeout=timeout) as resp:
                return json.loads(resp.read().decode())
        except (URLError, HTTPError, json.JSONDecodeError) as e:
            log.warning(f"DELETE {path} failed: {e}")
            return None
'''


def generate_client(endpoints):
    """Generate the full Python client file."""
    groups = group_by_prefix(endpoints)

    parts = [MODULE_DOC]
    parts.append("import json")
    parts.append("import logging")
    parts.append("from urllib.error import HTTPError, URLError")
    parts.append("from urllib.parse import urlencode")
    parts.append("from urllib.request import Request, urlopen")
    parts.append("")
    parts.append(f'log = logging.getLogger("{MODULE_NAME}")')
    parts.append("")
    parts.append("")
    parts.append(f"class {CLASS_NAME}:")
    parts.append(f'    """Client for {CLASS_NAME.replace("Client", " Access")} HTTP API."""')
    parts.append("")

    # Constructor
    if HAS_ACCOUNT:
        parts.append(f'    def __init__(self, base_url="{BASE_URL}", account=None):')
        parts.append('        self.base_url = base_url.rstrip("/")')
        parts.append("        self.default_account = account")
    else:
        parts.append(f'    def __init__(self, base_url="{BASE_URL}"):')
        parts.append('        self.base_url = base_url.rstrip("/")')

    # Low-level helpers
    parts.append("")
    parts.append("    # " + "-" * 66)
    parts.append("    # Low-level helpers")
    parts.append("    # " + "-" * 66)
    parts.append("")

    if HAS_ACCOUNT:
        parts.append('    def _params(self, account=None, **kwargs):')
        parts.append('        """Build query params dict, injecting account."""')
        parts.append("        p = {}")
        parts.append("        acct = account or self.default_account")
        parts.append("        if acct:")
        parts.append('            p["account"] = acct')
        parts.append("        for k, v in kwargs.items():")
        parts.append("            if v is not None:")
        parts.append("                p[k] = v")
        parts.append("        return p")
        parts.append("")

    parts.append(HTTP_HELPERS)

    # Methods grouped by prefix
    for prefix, eps in groups.items():
        parts.append("    # " + "-" * 66)
        parts.append(f"    # {prefix.replace('-', ' ').title()}")
        parts.append("    # " + "-" * 66)
        for ep in eps:
            parts.append("")
            parts.append(gen_method(ep))

    return "\n".join(parts) + "\n"


# ── README generation ──────────────────────────────────────────────────────

README_START = "<!-- GEN:API_START -->"
README_END = "<!-- GEN:API_END -->"


def generate_readme_section(endpoints):
    """Generate markdown API reference section."""
    groups = group_by_prefix(endpoints)
    lines = [README_START, "", "## API Reference", ""]
    lines.append("> Auto-generated from OpenAPI spec. Do not edit manually.")
    lines.append("> Regenerate with: `make gen-client`")
    lines.append("")

    for prefix, eps in groups.items():
        title = prefix.replace("-", " ").title()
        lines.append(f"### {title}")
        lines.append("")
        lines.append("| Method | Endpoint | Description |")
        lines.append("|--------|----------|-------------|")
        for ep in eps:
            m = ep["method"].upper()
            path = ep["path"]
            # Add query param hints
            qp = [p["name"] for p in ep["query_params"]]
            if qp:
                path += "?" + "&".join(f"{q}=..." for q in qp[:3])
                if len(qp) > 3:
                    path += "&..."
            desc = ep["summary"].split("\n")[0][:80] if ep["summary"] else ""
            lines.append(f"| {m} | `{path}` | {desc} |")
        lines.append("")

    lines.append(README_END)
    return "\n".join(lines)


def update_readme(section_text):
    """Update README.md between markers, or append if markers not found."""
    readme_path = Path("README.md")
    if not readme_path.exists():
        print("  README.md not found, skipping")
        return False

    content = readme_path.read_text()
    if README_START in content and README_END in content:
        before = content[: content.index(README_START)]
        after = content[content.index(README_END) + len(README_END) :]
        new_content = before + section_text + after
    else:
        # Append with markers
        new_content = content.rstrip() + "\n\n" + section_text + "\n"

    readme_path.write_text(new_content)
    return True


# ── Main ───────────────────────────────────────────────────────────────────

SPEC_FILE = "openapi.json"  # committed snapshot of the spec


def fetch_spec_with_fallback(source=None):
    """Fetch spec from source, server, or committed snapshot (in that order)."""
    # Explicit source (file or URL)
    if source:
        print(f"Fetching OpenAPI spec from {source}...")
        return fetch_spec(source)

    # Try live server first
    try:
        print(f"Fetching OpenAPI spec from {SPEC_URL}...")
        return fetch_spec(SPEC_URL)
    except Exception:
        pass

    # Fall back to committed snapshot
    if Path(SPEC_FILE).exists():
        print(f"Server not reachable, using committed {SPEC_FILE}...")
        return fetch_spec(SPEC_FILE)

    print(f"ERROR: Cannot fetch spec from {SPEC_URL} and no {SPEC_FILE} found.")
    print(f"  Start the server or run 'make gen-client' first to create {SPEC_FILE}.")
    sys.exit(1)


def main():
    check_mode = "--check" in sys.argv
    source = None
    for arg in sys.argv[1:]:
        if not arg.startswith("--"):
            source = arg

    spec = fetch_spec_with_fallback(source)
    print(f"  Found {len(spec['paths'])} paths")

    endpoints = parse_endpoints(spec)
    print(f"  Parsed {len(endpoints)} endpoints (after filtering)")

    # Generate client
    client_code = generate_client(endpoints)
    client_path = Path(CLIENT_FILE)

    if check_mode:
        if client_path.exists() and client_path.read_text() == client_code:
            print(f"  {CLIENT_FILE} is up-to-date")
        else:
            print(f"  {CLIENT_FILE} is STALE — run 'make gen-client' to update")
            sys.exit(1)
    else:
        client_path.parent.mkdir(parents=True, exist_ok=True)
        client_path.write_text(client_code)
        print(f"  Generated {CLIENT_FILE} ({len(endpoints)} methods)")

        # Save spec snapshot for offline checks
        Path(SPEC_FILE).write_text(json.dumps(spec, indent=2) + "\n")
        print(f"  Saved {SPEC_FILE} (for offline check-client)")

    # Generate README section
    readme_section = generate_readme_section(endpoints)

    if check_mode:
        readme_path = Path("README.md")
        if readme_path.exists():
            content = readme_path.read_text()
            if README_START in content:
                current = content[
                    content.index(README_START) : content.index(README_END)
                    + len(README_END)
                ]
                if current == readme_section:
                    print("  README.md API section is up-to-date")
                else:
                    print(
                        "  README.md API section is STALE — run 'make gen-client' to update"
                    )
                    sys.exit(1)
    else:
        if update_readme(readme_section):
            print("  Updated README.md API Reference section")

    print("Done.")


if __name__ == "__main__":
    main()
