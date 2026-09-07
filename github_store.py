"""
Reads and writes the master workbook straight from/to a GitHub repo via the
Contents API, so it survives Streamlit Cloud app restarts/redeploys (local
disk on Streamlit Cloud is not reliably persistent).

Needs three values, normally supplied via st.secrets:
  GITHUB_TOKEN     - a fine-grained Personal Access Token with
                     Contents: Read and write on this one repo
  GITHUB_REPO      - "your-username/your-repo-name"
  GITHUB_FILE_PATH - path inside the repo, e.g. "data/master.xlsx"
"""

import base64
import requests
from urllib.parse import quote

API_ROOT = "https://api.github.com"


def _headers(token):
    token = str(token).strip()
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "manpower-reconciliation-app",
    }


def get_file(token, repo, path, branch="main"):
    """Returns (content_bytes, sha). Raises on error."""
    repo = str(repo).strip().strip("/")
    path = str(path).strip().lstrip("/")
    branch = str(branch or "main").strip()
    if "/" not in repo:
        raise ValueError("GITHUB_REPO must be in the form owner/repository")
    if not path:
        raise ValueError("GITHUB_FILE_PATH cannot be empty")

    # Check repository access separately so a 404 can be diagnosed correctly.
    repo_url = f"{API_ROOT}/repos/{repo}"
    repo_resp = requests.get(repo_url, headers=_headers(token), timeout=30)
    if repo_resp.status_code == 401:
        raise RuntimeError("GitHub returned 401 Unauthorized. Your token is invalid, expired, or not being passed correctly.")
    if repo_resp.status_code == 403:
        raise RuntimeError("GitHub returned 403 Forbidden. Check token permissions and repository access.")
    if repo_resp.status_code == 404:
        raise RuntimeError(f"GitHub cannot access repository '{repo}'. Check GITHUB_REPO and make sure the token has access to this repository.")
    repo_resp.raise_for_status()

    encoded_path = quote(path, safe="/")
    url = f"{API_ROOT}/repos/{repo}/contents/{encoded_path}"
    resp = requests.get(url, headers=_headers(token), params={"ref": branch}, timeout=30)
    if resp.status_code == 401:
        raise RuntimeError("GitHub returned 401 Unauthorized while reading the file. Check the token.")
    if resp.status_code == 403:
        raise RuntimeError("GitHub returned 403 Forbidden while reading the file. Check token Contents permission.")
    if resp.status_code == 404:
        raise RuntimeError(f"GitHub cannot find '{path}' on branch '{branch}' in repository '{repo}'. Check GITHUB_FILE_PATH and GITHUB_BRANCH.")
    resp.raise_for_status()
    data = resp.json()
    if data.get("type") != "file":
        raise RuntimeError(f"GitHub path '{path}' is not a file (type={data.get('type')!r}).")
    if "content" not in data:
        raise RuntimeError("GitHub returned the file metadata but no file content. Check the file size/API response.")
    content = base64.b64decode(data["content"])
    return content, data["sha"]


def put_file(token, repo, path, content_bytes, sha, message, branch="main"):
    """Commits new content to the given path (sha = the current file's sha,
    from get_file, required so GitHub knows you're updating, not creating)."""
    url = f"{API_ROOT}/repos/{repo}/contents/{path}"
    payload = {
        "message": message,
        "content": base64.b64encode(content_bytes).decode("utf-8"),
        "sha": sha,
        "branch": branch,
    }
    resp = requests.put(url, headers=_headers(token), json=payload, timeout=30)
    resp.raise_for_status()
    return resp.json()
