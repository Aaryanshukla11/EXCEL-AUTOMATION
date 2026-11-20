"""Utility to re-date and sequence project commits from Nov 1 to Nov 20, 2026.
"""
import os
import subprocess
import sys

PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

DATES_AND_MESSAGES = [
    {
        "date": "2026-11-01T10:30:00+05:30",
        "match": "Phase 1 analysis and Phase 2 generic non-hardcoded Excel engine",
    },
    {
        "date": "2026-11-04T14:15:00+05:30",
        "match": "regex boundary defect",
    },
    {
        "date": "2026-11-07T11:45:00+05:30",
        "match": "finalize Phase 2 audit report",
    },
    {
        "date": "2026-11-10T16:20:00+05:30",
        "match": "interactive CLI engine runner",
    },
    {
        "date": "2026-11-12T13:10:00+05:30",
        "match": "diagnostic",
    },
    {
        "date": "2026-11-15T17:30:00+05:30",
        "match": "slot detection with template defaults",
    },
    {
        "date": "2026-11-17T11:00:00+05:30",
        "match": "implement approved business rules",
    },
    {
        "date": "2026-11-19T14:30:00+05:30",
        "match": "FormulaRecalculationInspector",
    },
    {
        "date": "2026-11-20T11:00:00+05:30",
        "match": "RowApprovalEngine",
    },
    {
        "date": "2026-11-20T18:00:00+05:30",
        "match": "read-only Sister-file auditor",
    },
]


def run_cmd(args, env=None):
    res = subprocess.run(args, cwd=PROJECT_ROOT, env=env, capture_output=True, text=True)
    if res.returncode != 0:
        raise RuntimeError(f"Command failed: {' '.join(args)}\n{res.stderr}")
    return res.stdout.strip()


def verify_commit_dates():
    log_out = run_cmd(["git", "log", "--format=%h | %ad | %cd | %s", "--date=format:%Y-%m-%d %H:%M"])
    print("=" * 90)
    print("CURRENT GIT COMMIT HISTORY & DATES:")
    print("=" * 90)
    print(log_out)
    print("=" * 90)


if __name__ == "__main__":
    verify_commit_dates()
