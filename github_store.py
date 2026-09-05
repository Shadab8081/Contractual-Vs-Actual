"""GitHub Contents API helpers for the master workbook."""

import base64
import requests
from urllib.parse import quote

API_ROOT = "https://api.github.com"


class GitHubStoreError(RuntimeError):
    """Friendly, actionable GitHub API error."""


def _headers(token):
    return {
        "Authorization": f"Bearer {token.strip()}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "manpower-reconciliation-app",
    }


def _raise_for_github(resp, action, repo, path=None):
    if resp.ok:
        return
    try:
        body = resp.json()
        message = body.get("message", "")
        documentation = body.get("documentation_url", "")
    except ValueError:
        message = resp.text[:300]
        documentation = ""

    status = resp.status_code
    if status == 401:
        detail = (
            "GitHub rejected the token (401 Unauthorized). Create a valid GitHub "
            "Personal Access Token and give it access to this repository. For a "
            "fine-grained token, select this repository and enable Repository "
            "permissions → Contents → Read and write. Then update GITHUB_TOKEN "
            "in Streamlit Cloud → App settings → Secrets and redeploy."
        )
    elif status == 403:
        detail = (
            "GitHub understood the token but denied access (403 Forbidden). Check "
            "that the token has Contents permission and that the selected repository "
            "is Shadab8081/Contractual-Vs-Actual. If your organization uses SSO, "
            "authorize the token for the organization as required."
        )
    elif status == 404:
        detail = (
            "GitHub could not find the repository or file (404 Not Found). Check "
            f"GITHUB_REPO='{repo}', GITHUB_FILE_PATH='{path or ''}', and the branch."
        )
    else:
        detail = f"GitHub returned HTTP {status}."

    extra = f" GitHub message: {message}" if message else ""
    if documentation:
        extra += f" Documentation: {documentation}"
    raise GitHubStoreError(f"{action} failed. {detail}{extra}")


def validate_token(token, repo):
    """Validate token authentication and repository access before file operations."""
    if not token or not token.strip():
        raise GitHubStoreError("GITHUB_TOKEN is empty.")
    if not repo or "/" not in repo:
        raise GitHubStoreError(
            "GITHUB_REPO must look like 'Shadab8081/Contractual-Vs-Actual'."
        )

    resp = requests.get(
        f"{API_ROOT}/repos/{quote(repo, safe='/')}",
        headers=_headers(token),
        timeout=30,
    )
    _raise_for_github(resp, "GitHub authentication/repository access check", repo)
    data = resp.json()
    return {
        "full_name": data.get("full_name"),
        "default_branch": data.get("default_branch"),
        "private": bool(data.get("private")),
    }


def get_file(token, repo, path, branch="main"):
    """Return (content_bytes, sha) for a repository file."""
    if not token:
        raise GitHubStoreError("GITHUB_TOKEN is missing.")
    if not repo:
        raise GitHubStoreError("GITHUB_REPO is missing.")

    url = f"{API_ROOT}/repos/{quote(repo, safe='/')}/contents/{quote(path, safe='/')}"
    resp = requests.get(
        url,
        headers=_headers(token),
        params={"ref": branch},
        timeout=30,
    )
    _raise_for_github(resp, "Reading the master workbook", repo, path)

    data = resp.json()
    if data.get("type") != "file":
        raise GitHubStoreError(
            f"The GitHub path '{path}' is not a file (type={data.get('type')!r})."
        )
    if not data.get("content"):
        raise GitHubStoreError(
            f"GitHub returned no file content for '{path}'. Check that the file is "
            "committed to the selected branch and is within the Contents API limits."
        )

    try:
        content = base64.b64decode(data["content"], validate=False)
    except Exception as exc:
        raise GitHubStoreError(f"Could not decode '{path}' from GitHub: {exc}") from exc

    return content, data["sha"]


def put_file(token, repo, path, content_bytes, sha, message, branch="main"):
    """Commit updated file content to GitHub."""
    if not token:
        raise GitHubStoreError("GITHUB_TOKEN is missing.")
    if not repo:
        raise GitHubStoreError("GITHUB_REPO is missing.")

    url = f"{API_ROOT}/repos/{quote(repo, safe='/')}/contents/{quote(path, safe='/')}"
    payload = {
        "message": message,
        "content": base64.b64encode(content_bytes).decode("utf-8"),
        "sha": sha,
        "branch": branch,
    }
    resp = requests.put(
        url,
        headers=_headers(token),
        json=payload,
        timeout=30,
    )
    _raise_for_github(resp, "Saving the new master workbook", repo, path)
    return resp.json()
