"""Jira REST API v2 client cho Jira Server / Data Center.

Xac thuc bang Personal Access Token (Bearer) + header X-AUSERNAME.
"""
from __future__ import annotations

import json
from typing import Any, Optional

import requests


class JiraError(Exception):
    """Loi tra ve tu Jira API."""

    def __init__(self, message: str, status_code: Optional[int] = None, payload: Any = None):
        super().__init__(message)
        self.status_code = status_code
        self.payload = payload


class JiraClient:
    def __init__(
        self,
        base_url: str,
        token: str,
        username: str = "",
        cookie: str = "",
        verify_ssl: bool = True,
        timeout: int = 30,
    ):
        if not base_url:
            raise ValueError("Thieu JIRA_URL")
        if not token:
            raise ValueError("Thieu JIRA_TOKEN")

        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self.session = requests.Session()
        self.session.verify = verify_ssl

        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }
        if username:
            headers["X-AUSERNAME"] = username
        if cookie:
            headers["Cookie"] = cookie
        self.session.headers.update(headers)

    def _url(self, path: str) -> str:
        return f"{self.base_url}/rest/api/2/{path.lstrip('/')}"

    def _agile_url(self, path: str) -> str:
        return f"{self.base_url}/rest/agile/1.0/{path.lstrip('/')}"

    def _request(self, method: str, url: str, **kwargs) -> Any:
        try:
            resp = self.session.request(method, url, timeout=self.timeout, **kwargs)
        except requests.exceptions.SSLError as exc:
            raise JiraError(
                f"Loi SSL khi ket noi {url}. Neu server dung self-signed cert, "
                f"dat JIRA_VERIFY_SSL=false trong .env. Chi tiet: {exc}"
            ) from exc
        except requests.exceptions.RequestException as exc:
            raise JiraError(f"Loi ket noi {url}: {exc}") from exc

        if resp.status_code >= 400:
            payload: Any
            try:
                payload = resp.json()
            except ValueError:
                payload = resp.text
            raise JiraError(
                f"Jira tra ve {resp.status_code}: {_format_error(payload)}",
                status_code=resp.status_code,
                payload=payload,
            )

        if resp.status_code == 204 or not resp.content:
            return None
        try:
            return resp.json()
        except ValueError:
            return resp.text

    def myself(self) -> dict:
        return self._request("GET", self._url("myself"))

    def get_project(self, key: str) -> dict:
        return self._request("GET", self._url(f"project/{key}"))

    def get_issue(self, key: str, fields: str = "") -> dict:
        params = {"fields": fields} if fields else None
        return self._request("GET", self._url(f"issue/{key}"), params=params)

    def get_sprint_id_by_issue(self, issue_key: str, sprint_field: str = "") -> Optional[str]:
        """Lay sprint ID tu issue - giong get_sprint_id_by_issue trong du an."""
        # Thu agile API truoc (Jira Software)
        try:
            agile_issue = self._request("GET", self._agile_url(f"issue/{issue_key}"))
            sprint_id = _extract_sprint_id_from_fields(agile_issue.get("fields", {}), sprint_field)
            if sprint_id:
                return sprint_id
        except JiraError:
            pass

        fields_to_fetch = sprint_field if sprint_field else "labels,*navigable"
        issue = self.get_issue(issue_key, fields=fields_to_fetch)
        return _extract_sprint_id_from_fields(issue.get("fields", {}), sprint_field)

    def check_upcode_label_and_get_sprint_id(
        self, issue_key: str, sprint_field: str = ""
    ) -> tuple[bool, Optional[str]]:
        """Giong ham check_upcode_label_and_get_sprint_id trong du an."""
        issue = self.get_issue(issue_key, fields="labels")
        labels = issue.get("fields", {}).get("labels", [])
        has_upcode = "Upcode" in labels
        sprint_id = self.get_sprint_id_by_issue(issue_key, sprint_field=sprint_field)
        return has_upcode, sprint_id

    def create_issue(self, fields: dict) -> dict:
        return self._request("POST", self._url("issue"), data=json.dumps({"fields": fields}))

    def update_issue(self, issue_key: str, fields: dict) -> Any:
        return self._request(
            "PUT",
            self._url(f"issue/{issue_key}"),
            data=json.dumps({"fields": fields}),
        )

    def add_issue_to_sprint(self, sprint_id: str | int, issue_key: str) -> Any:
        return self._request(
            "POST",
            self._agile_url(f"sprint/{sprint_id}/issue"),
            data=json.dumps({"issues": [issue_key]}),
        )

    def delete_issue(self, issue_key: str) -> Any:
        return self._request("DELETE", self._url(f"issue/{issue_key}"))

    def get_sprint_issues(self, sprint_id: str | int, max_results: int = 100) -> list[dict]:
        """Lay tat ca issue trong sprint (co phan trang)."""
        issues: list[dict] = []
        start_at = 0
        while True:
            data = self._request(
                "GET",
                self._agile_url(f"sprint/{sprint_id}/issue"),
                params={"startAt": start_at, "maxResults": max_results},
            )
            batch = data.get("issues", [])
            issues.extend(batch)
            if start_at + len(batch) >= data.get("total", 0) or not batch:
                break
            start_at += len(batch)
        return issues


def _extract_sprint_id_from_fields(fields: dict, sprint_field: str = "") -> Optional[str]:
    candidates: list[Any] = []
    if sprint_field and sprint_field in fields:
        candidates.append(fields[sprint_field])

    for value in fields.values():
        if isinstance(value, dict) and value.get("id") and ("state" in value or "originBoardId" in value):
            candidates.append(value)
        elif isinstance(value, list):
            for item in value:
                if isinstance(item, dict) and item.get("id"):
                    candidates.append(item)

    for item in reversed(candidates):
        if isinstance(item, dict) and item.get("id"):
            state = item.get("state")
            if state in (None, "ACTIVE", "FUTURE", "active", "future"):
                return str(item["id"])
    for item in reversed(candidates):
        if isinstance(item, dict) and item.get("id"):
            return str(item["id"])
    return None


def _format_error(payload: Any) -> str:
    if isinstance(payload, dict):
        parts = []
        msgs = payload.get("errorMessages")
        if msgs:
            parts.extend(msgs)
        errs = payload.get("errors")
        if isinstance(errs, dict):
            parts.extend(f"{k}: {v}" for k, v in errs.items())
        if parts:
            return "; ".join(parts)
    return str(payload)
