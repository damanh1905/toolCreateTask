"""Tool import Task/Sub-task len Jira tu file Excel.

Cach dung:
    python main.py template                 # sinh file Excel mau (tasks_template.xlsx)
    python main.py validate tasks.xlsx      # kiem tra file + ket noi (khong tao gi)
    python main.py import tasks.xlsx        # tao issue that len Jira
    python main.py import tasks.xlsx --dry-run   # chi in ra, khong goi API
"""
from __future__ import annotations

import argparse
import os
import sys
from typing import Optional

from dotenv import load_dotenv

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

import excel_io
from excel_io import IssueRow
from jira_client import JiraClient, JiraError
from sprint_tasks import find_parent_key, get_tasks_tree_by_sprint_id, print_tasks_tree


def _load_client() -> JiraClient:
    load_dotenv()
    url = os.getenv("JIRA_URL", "").strip()
    token = os.getenv("JIRA_TOKEN", "").strip()
    username = os.getenv("JIRA_USERNAME", "").strip()
    cookie = os.getenv("JIRA_COOKIE", "").strip()
    verify = os.getenv("JIRA_VERIFY_SSL", "true").strip().lower() not in ("false", "0", "no")

    if not verify:
        import urllib3

        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    return JiraClient(url, token, username=username, cookie=cookie, verify_ssl=verify)


def _derive_project_from_key(key: str) -> str:
    return key.split("-")[0] if "-" in key else ""


def _validate_rows(rows: list[IssueRow]) -> list[str]:
    errors: list[str] = []
    local_ids = {r.local_id for r in rows if r.local_id}

    for r in rows:
        prefix = f"Dong {r.row_number}"
        if not r.summary:
            errors.append(f"{prefix}: thieu cot 'Tiêu đề'")
        if not r.assignee:
            errors.append(
                f"{prefix}: thieu cot 'Người xử lý' (hoac dat JIRA_USERNAME trong .env)"
            )
        if not r.issue_type:
            errors.append(f"{prefix}: thieu cot 'Loại'")

        if r.is_subtask:
            if not r.parent_key:
                errors.append(f"{prefix}: Sub-task can co cot 'Task cha'")
            else:
                # project phai xac dinh duoc: tu ProjectKey, hoac tu parent key that, hoac tu task cha trong file
                has_project = bool(r.project_key)
                is_local_parent = r.parent_key in local_ids
                is_real_key = "-" in r.parent_key
                if not (has_project or is_local_parent or is_real_key):
                    errors.append(
                        f"{prefix}: khong xac dinh duoc project cho subtask "
                        f"(kiem tra JIRA_DEFAULT_PROJECT_KEY hoac dung 'Task cha' la key Jira hop le)"
                    )
        else:
            if not r.project_key:
                errors.append(f"{prefix}: thieu project (dien JIRA_DEFAULT_PROJECT_KEY trong .env)")

        if r.custom_fields:
            import json

            try:
                json.loads(r.custom_fields)
            except json.JSONDecodeError as exc:
                errors.append(f"{prefix}: CustomFields khong phai JSON hop le ({exc})")

    return errors


def cmd_sprint_tree(args) -> int:
    """Hien thi cay task/subtask trong sprint de nhan dien Task cha."""
    try:
        client = _load_client()
        sprint_id = ""
        if args.sprint_id:
            sprint_id = args.sprint_id.strip()
        elif args.issue_key:
            _, sprint_id = client.check_upcode_label_and_get_sprint_id(
                args.issue_key, os.getenv("JIRA_FIELD_SPRINT", "").strip()
            )
        elif os.getenv("JIRA_REFERENCE_ISSUE", "").strip():
            _, sprint_id = client.check_upcode_label_and_get_sprint_id(
                os.getenv("JIRA_REFERENCE_ISSUE", "").strip(),
                os.getenv("JIRA_FIELD_SPRINT", "").strip(),
            )
        else:
            sprint_id = os.getenv("JIRA_SPRINT_ID", "").strip()

        if not sprint_id:
            print("Khong lay duoc sprint. Truyen --issue-key hoac --sprint-id")
            return 1
        tree = get_tasks_tree_by_sprint_id(
            client,
            sprint_id,
            hide_closed_tasks=not args.show_closed,
            service_filter=args.service or "",
            create_by_filter=args.creator or "",
            priority_filter=args.priority or "",
            issuetype=args.issuetype or "",
        )
        print(f"Sprint ID: {sprint_id}")
        print_tasks_tree(tree)

        if args.lookup:
            key = find_parent_key(tree, args.lookup)
            if key:
                print(f"\nTim thay '{args.lookup}' -> key: {key}")
            else:
                print(f"\nKhong tim thay task/subtask co tieu de: {args.lookup}")
                return 1
    except (JiraError, ValueError) as exc:
        print(f"Loi: {exc}")
        return 1
    return 0


def cmd_test(args) -> int:
    """Kiem tra ket noi Jira (giong curl GET issue)."""
    try:
        client = _load_client()
        me = client.myself()
        name = me.get("name") or me.get("displayName") or "?"
        print(f"Ket noi OK. User: {name}")

        if args.issue:
            issue = client.get_issue(args.issue)
            fields = issue.get("fields", {})
            summary = fields.get("summary", "?")
            itype = (fields.get("issuetype") or {}).get("name", "?")
            project = (fields.get("project") or {}).get("key", "?")
            print(f"Issue {args.issue}: [{itype}] {summary} (project: {project})")
    except (JiraError, ValueError) as exc:
        print(f"Loi: {exc}")
        return 1
    return 0


def cmd_template(args) -> int:
    path = args.output
    excel_io.write_template(path, with_sample=not args.no_sample)
    print(f"Da tao file mau: {os.path.abspath(path)}")
    print("Mo file, dien du lieu vao sheet 'Issues' roi chay: python main.py import <file>")
    return 0


def cmd_validate(args) -> int:
    rows = excel_io.read_issues(args.file)
    if not rows:
        print("File khong co du lieu.")
        return 1

    errors = _validate_rows(rows)
    tasks = [r for r in rows if not r.is_subtask]
    subtasks = [r for r in rows if r.is_subtask]
    print(f"Doc duoc {len(rows)} dong: {len(tasks)} task/issue, {len(subtasks)} sub-task.")

    if errors:
        print(f"\nCo {len(errors)} loi can sua:")
        for e in errors:
            print(f"  - {e}")
        return 1

    print("Du lieu hop le.")

    # Kiem tra ket noi Jira
    try:
        client = _load_client()
        me = client.myself()
        name = me.get("name") or me.get("displayName") or "?"
        print(f"Ket noi Jira OK. Dang nhap voi user: {name}")
    except (JiraError, ValueError) as exc:
        print(f"Canh bao: khong ket noi duoc Jira: {exc}")
        return 1
    return 0


def _is_jira_key(value: str) -> bool:
    if not value or value.startswith("DRYRUN-"):
        return False
    if "-" not in value:
        return False
    project, num = value.rsplit("-", 1)
    return bool(project) and num.isdigit()


def _find_reference_issue(rows: list[IssueRow], cli_issue_key: str = "") -> str:
    """Tim issue tham chieu de lay sprint - giong project_id trong du an."""
    ref = (cli_issue_key or os.getenv("JIRA_REFERENCE_ISSUE", "")).strip()
    if ref:
        return ref
    for row in rows:
        if row.parent_key and _is_jira_key(row.parent_key):
            return row.parent_key
    return ""


def _resolve_sprint_id(
    client: JiraClient,
    parent_key: str = "",
    reference_issue: str = "",
    cli_sprint_id: str = "",
) -> str:
    """Lay sprint ID tu issue (uu tien task cha / issue tham chieu), khong can truyen tay."""
    sprint_field = os.getenv("JIRA_FIELD_SPRINT", "").strip()

    for issue_key in (parent_key, reference_issue):
        if issue_key and _is_jira_key(issue_key):
            sprint_id = client.get_sprint_id_by_issue(issue_key, sprint_field=sprint_field)
            if sprint_id:
                return sprint_id

    return (cli_sprint_id or os.getenv("JIRA_SPRINT_ID", "")).strip()


def cmd_import(args) -> int:
    rows = excel_io.read_issues(args.file)
    if not rows:
        print("File khong co du lieu.")
        return 1

    errors = _validate_rows(rows)
    if errors:
        print(f"Co {len(errors)} loi trong file, dung lai:")
        for e in errors:
            print(f"  - {e}")
        return 1

    client: Optional[JiraClient] = None
    try:
        client = _load_client()
        if not args.dry_run:
            client.myself()
    except (JiraError, ValueError) as exc:
        print(f"Khong ket noi duoc Jira: {exc}")
        return 1

    assert client is not None
    reference_issue = _find_reference_issue(rows, getattr(args, "issue_key", "") or "")
    default_sprint_id = _resolve_sprint_id(
        client,
        reference_issue=reference_issue,
        cli_sprint_id=getattr(args, "sprint_id", "") or "",
    )

    if not default_sprint_id:
        print(
            "Khong lay duoc Sprint ID. Hay cung cap mot issue trong sprint:\n"
            "  - Dien JIRA_REFERENCE_ISSUE=MYTVB2C-xxxxx trong .env, hoac\n"
            "  - Truyen --issue-key MYTVB2C-xxxxx, hoac\n"
            "  - Dien cot 'Task cha' la key Jira co san trong file Excel"
        )
        return 1

    if reference_issue:
        has_upcode, sprint_from_ref = client.check_upcode_label_and_get_sprint_id(
            reference_issue, os.getenv("JIRA_FIELD_SPRINT", "").strip()
        )
        print(f"Issue tham chieu: {reference_issue}")
        print(f"  Label Upcode: {'co' if has_upcode else 'khong'}")
        if sprint_from_ref:
            print(f"  Sprint ID tu issue: {sprint_from_ref}")

    print(f"Sprint ID: {default_sprint_id} (issue moi se duoc add vao sprint nay)")

    tasks = [r for r in rows if not r.is_subtask]
    subtasks = [r for r in rows if r.is_subtask]

    local_to_key: dict[str, str] = {}
    created = 0
    failed = 0

    def handle_create(row: IssueRow, fields: dict, parent_key: str = "") -> Optional[str]:
        nonlocal created, failed
        note_update = excel_io.build_note_update(row)
        sprint_id = (
            _resolve_sprint_id(
                client,
                parent_key=parent_key,
                reference_issue=reference_issue,
                cli_sprint_id=getattr(args, "sprint_id", "") or "",
            )
            or default_sprint_id
        )

        if args.dry_run:
            print(f"  [DRY-RUN] Dong {row.row_number} ({fields['issuetype']['name']}): {row.summary}")
            import json

            print(f"           fields = {json.dumps(fields, ensure_ascii=False)}")
            if note_update:
                print(f"           note_update = {json.dumps(note_update, ensure_ascii=False)}")
            print(f"           add_to_sprint = {sprint_id}")
            created += 1
            return f"DRYRUN-{row.row_number}"
        try:
            assert client is not None
            result = client.create_issue(fields)
            key = result.get("key", "?")

            if note_update:
                client.update_issue(key, note_update)

            try:
                client.add_issue_to_sprint(sprint_id, key)
            except JiraError as exc:
                print(f"  LOI Dong {row.row_number}: add sprint that bai, xoa issue {key}")
                try:
                    client.delete_issue(key)
                except JiraError:
                    pass
                raise exc

            print(f"  OK  Dong {row.row_number}: tao {key} + add sprint {sprint_id} - {row.summary}")
            created += 1
            return key
        except JiraError as exc:
            print(f"  LOI Dong {row.row_number}: {exc}")
            failed += 1
            if not args.continue_on_error:
                raise
            return None

    print("=== Buoc 1: tao Task/Issue cha ===")
    try:
        for row in tasks:
            fields = excel_io.build_fields(row, row.project_key, parent_key=None)
            key = handle_create(row, fields, parent_key="")
            if key and row.local_id:
                local_to_key[row.local_id] = key
    except JiraError:
        print("\nDung lai do gap loi (them --continue-on-error de bo qua dong loi).")
        return 1

    print("\n=== Buoc 2: tao Sub-task ===")
    try:
        for row in subtasks:
            parent_ref = row.parent_key
            parent_key = local_to_key.get(parent_ref, parent_ref)
            project_key = row.project_key or _derive_project_from_key(parent_key)
            if not project_key:
                print(f"  LOI Dong {row.row_number}: khong xac dinh duoc ProjectKey")
                failed += 1
                if not args.continue_on_error:
                    return 1
                continue
            fields = excel_io.build_fields(row, project_key, parent_key=parent_key)
            handle_create(row, fields, parent_key=parent_key)
    except JiraError:
        print("\nDung lai do gap loi (them --continue-on-error de bo qua dong loi).")
        return 1

    print(f"\n=== Ket qua: tao thanh cong {created}, that bai {failed} ===")
    return 0 if failed == 0 else 2


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Import Task/Sub-task len Jira tu file Excel."
    )
    sub = parser.add_subparsers(dest="command", required=True)

    p_tree = sub.add_parser("sprint-tree", help="Xem cay task/subtask trong sprint de nhan dien Task cha")
    p_tree.add_argument("-s", "--sprint-id", help="Sprint ID (neu da biet)")
    p_tree.add_argument("-k", "--issue-key", help="Issue tham chieu de lay sprint (VD: MYTVB2C-50784)")
    p_tree.add_argument("--service", help="Loc theo dich vu")
    p_tree.add_argument("--creator", help="Loc theo nguoi tao")
    p_tree.add_argument("--priority", help="Loc theo priority")
    p_tree.add_argument("--issuetype", help="Loc theo loai issue (Task, Sub-task...)")
    p_tree.add_argument("--show-closed", action="store_true", help="Hien ca task da dong")
    p_tree.add_argument("--lookup", help="Tim key theo tieu de task (de dien cot Task cha)")
    p_tree.set_defaults(func=cmd_sprint_tree)

    p_test = sub.add_parser("test", help="Kiem tra ket noi Jira (giong curl GET issue)")
    p_test.add_argument(
        "-i", "--issue", default="MYTVB2C-51653", help="Key issue de test GET (mac dinh: MYTVB2C-51653)"
    )
    p_test.set_defaults(func=cmd_test)

    p_tpl = sub.add_parser("template", help="Sinh file Excel mau")
    p_tpl.add_argument("-o", "--output", default="tasks_template.xlsx", help="Ten file xuat ra")
    p_tpl.add_argument("--no-sample", action="store_true", help="Khong them dong vi du")
    p_tpl.set_defaults(func=cmd_template)

    p_val = sub.add_parser("validate", help="Kiem tra file va ket noi (khong tao issue)")
    p_val.add_argument("file", help="Duong dan file Excel")
    p_val.set_defaults(func=cmd_validate)

    p_imp = sub.add_parser("import", help="Tao issue len Jira tu file Excel")
    p_imp.add_argument("file", help="Duong dan file Excel")
    p_imp.add_argument(
        "-k", "--issue-key",
        help="Issue tham chieu de lay sprint (VD: MYTVB2C-50784). Mac dinh lay tu JIRA_REFERENCE_ISSUE",
    )
    p_imp.add_argument("-s", "--sprint-id", help="Sprint ID (chi dung khi khong lay duoc tu issue)")
    p_imp.add_argument("--dry-run", action="store_true", help="Chi in ra, khong goi API")
    p_imp.add_argument(
        "--continue-on-error", action="store_true", help="Bo qua dong loi va tiep tuc"
    )
    p_imp.set_defaults(func=cmd_import)

    return parser


def main(argv: Optional[list[str]] = None) -> int:
    load_dotenv()
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except FileNotFoundError as exc:
        print(f"Khong tim thay file: {exc}")
        return 1
    except KeyboardInterrupt:
        print("\nDa huy.")
        return 130


if __name__ == "__main__":
    sys.exit(main())
