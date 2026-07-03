"""Lay va hien thi cay task/subtask trong sprint - giong get_tasks_tree_by_sprint_id."""
from __future__ import annotations

import os
from typing import Any, Optional

from jira_client import JiraClient


def _field_value(fields: dict, env_name: str) -> Any:
    key = os.getenv(env_name, "").strip()
    if not key:
        return None
    return fields.get(key)


def _normalize_issue(issue: dict) -> dict:
    fields = issue.get("fields", {})
    assignee = fields.get("assignee") or {}
    reporter = fields.get("reporter") or {}
    status = fields.get("status") or {}
    priority = fields.get("priority") or {}
    issuetype = fields.get("issuetype") or {}
    parent = fields.get("parent") or {}

    return {
        "key": issue.get("key", ""),
        "summary": fields.get("summary", ""),
        "issuetype": issuetype.get("name", ""),
        "status": status.get("name", ""),
        "status_category": (status.get("statusCategory") or {}).get("key", ""),
        "assignee": assignee.get("name", ""),
        "assignee_name": assignee.get("displayName", ""),
        "reporter": reporter.get("name", ""),
        "reporter_name": reporter.get("displayName", ""),
        "priority": priority.get("name", ""),
        "service": _field_value(fields, "JIRA_FIELD_SERVICE"),
        "duedate": fields.get("duedate", ""),
        "parent_key": parent.get("key", ""),
    }


def _match_filter(value: str, needle: str) -> bool:
    if not needle:
        return True
    return needle.lower() in (value or "").lower()


def _passes_filters(
    node: dict,
    hide_closed: bool,
    service_filter: str,
    create_by_filter: str,
    priority_filter: str,
    issuetype_filter: str,
) -> bool:
    if hide_closed and node["status_category"] == "done":
        return False
    if service_filter and not _match_filter(str(node.get("service") or ""), service_filter):
        return False
    if create_by_filter and not (
        _match_filter(node.get("reporter_name", ""), create_by_filter)
        or _match_filter(node.get("reporter", ""), create_by_filter)
    ):
        return False
    if priority_filter and not _match_filter(node.get("priority", ""), priority_filter):
        return False
    if issuetype_filter and not _match_filter(node.get("issuetype", ""), issuetype_filter):
        return False
    return True


def get_tasks_tree_by_sprint_id(
    client: JiraClient,
    sprint_id: str | int,
    hide_closed_tasks: bool = True,
    service_filter: str = "",
    create_by_filter: str = "",
    priority_filter: str = "",
    issuetype: str = "",
) -> list[dict]:
    """Lay cay task/subtask trong sprint, giong logic du an."""
    raw_issues = client.get_sprint_issues(sprint_id)
    nodes = [_normalize_issue(issue) for issue in raw_issues]

    tasks: dict[str, dict] = {}
    subtasks_by_parent: dict[str, list[dict]] = {}

    for node in nodes:
        if node["parent_key"]:
            subtasks_by_parent.setdefault(node["parent_key"], []).append(node)
        else:
            tasks[node["key"]] = {**node, "children": []}

    # Subtask co parent nam ngoai sprint van hien duoi parent neu parent co trong sprint
    for parent_key, children in subtasks_by_parent.items():
        if parent_key not in tasks:
            continue
        for child in children:
            if _passes_filters(
                child, hide_closed_tasks, service_filter, create_by_filter, priority_filter, issuetype
            ):
                tasks[parent_key]["children"].append(child)

    tree: list[dict] = []
    for task in tasks.values():
        if not _passes_filters(
            task, hide_closed_tasks, service_filter, create_by_filter, priority_filter, issuetype
        ):
            if not task["children"]:
                continue
        tree.append(task)

    tree.sort(key=lambda x: x["key"])
    return tree


def find_task_by_summary(tree: list[dict], summary: str) -> Optional[dict]:
    needle = summary.strip().lower()
    for task in tree:
        if task["summary"].strip().lower() == needle:
            return task
        for child in task.get("children", []):
            if child["summary"].strip().lower() == needle:
                return child
    return None


def find_parent_key(tree: list[dict], summary: str) -> str:
    """Tim key task cha theo tieu de - dung khi dien cot Task cha trong Excel."""
    node = find_task_by_summary(tree, summary)
    if not node:
        return ""
    if node.get("parent_key"):
        return node["parent_key"]
    return node["key"]


def print_tasks_tree(tree: list[dict]) -> None:
    if not tree:
        print("Khong co task nao trong sprint (hoac bi loc het).")
        return

    print(f"Tong {len(tree)} task cha:")
    for task in tree:
        assignee = task.get("assignee_name") or task.get("assignee") or "-"
        service = task.get("service") or "-"
        print(
            f"\n[{task['key']}] {task['summary']}\n"
            f"  Loai: {task['issuetype']} | Trang thai: {task['status']} | "
            f"Nguoi xu ly: {assignee} | Dich vu: {service}"
        )
        for child in task.get("children", []):
            c_assignee = child.get("assignee_name") or child.get("assignee") or "-"
            print(f"  └─ [{child['key']}] {child['summary']} ({child['issuetype']}, {child['status']}, {c_assignee})")
