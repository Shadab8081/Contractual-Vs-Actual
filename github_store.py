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

API_ROOT = "https://api.github.com"


def _headers(token):
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def get_file(token, repo, path, branch="main"):
    """Returns (content_bytes, sha). Raises on error."""
    url = f"{API_ROOT}/repos/{repo}/contents/{path}"
    resp = requests.get(url, headers=_headers(token), params={"ref": branch}, timeout=30)
    resp.raise_for_status()
    data = resp.json()
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
