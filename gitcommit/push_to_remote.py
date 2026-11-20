"""Helper script to add a remote Git repository and push master & main branches.
"""
import argparse
import subprocess
import sys
import os

PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


def run_git(args):
    print(f">> git {' '.join(args)}")
    res = subprocess.run(["git"] + args, cwd=PROJECT_ROOT, capture_output=True, text=True)
    if res.returncode != 0:
        print(f"Error:\n{res.stderr}")
        return False
    if res.stdout:
        print(res.stdout.strip())
    return True


def main():
    parser = argparse.ArgumentParser(description="Push local repository to remote GitHub/GitLab.")
    parser.add_argument("--url", type=str, help="Remote Git repository URL (e.g. https://github.com/user/repo.git)")
    parser.add_argument("--remote", type=str, default="origin", help="Remote name (default: origin)")
    args = parser.parse_args()

    remote_url = args.url
    if not remote_url:
        print("Enter your remote repository URL (e.g. https://github.com/username/repository.git):")
        remote_url = input("> ").strip()

    if not remote_url:
        print("No repository URL provided. Aborting.")
        sys.exit(1)

    # Check existing remotes
    existing_remotes = subprocess.run(
        ["git", "remote"], cwd=PROJECT_ROOT, capture_output=True, text=True
    ).stdout.split()

    if args.remote in existing_remotes:
        print(f"Updating remote '{args.remote}' to: {remote_url}")
        run_git(["remote", "set-url", args.remote, remote_url])
    else:
        print(f"Adding remote '{args.remote}': {remote_url}")
        run_git(["remote", "add", args.remote, remote_url])

    print("\nPushing 'master' branch...")
    ok_master = run_git(["push", "-u", args.remote, "master"])

    print("\nPushing 'main' branch...")
    ok_main = run_git(["push", "-u", args.remote, "main"])

    if ok_master and ok_main:
        print("\n[SUCCESS] Successfully pushed all commits and branches to remote repository!")
    else:
        print("\n[WARNING] Push completed with warnings or authentication required. Check console output above.")


if __name__ == "__main__":
    main()
