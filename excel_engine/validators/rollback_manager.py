"""Rollback and Recovery Manager.

Provides deterministic recovery to guaranteed safe pre-execution states using
authoritative backups. Never relies on reverse-cell patching.
"""
import hashlib
import os
import shutil
from typing import Tuple


class RollbackManager:
    """Manages restoration of master workbooks from authoritative backups."""

    @staticmethod
    def compute_sha256(filepath: str) -> str:
        """Compute SHA-256 hash of a file."""
        sha256 = hashlib.sha256()
        with open(filepath, "rb") as f:
            for chunk in iter(lambda: f.read(65536), b""):
                sha256.update(chunk)
        return sha256.hexdigest()

    def rollback(self, backup_filepath: str, target_working_copy_path: str, expected_hash: str) -> Tuple[bool, str]:
        """Restore target working copy from backup, verifying hash integrity."""
        if not os.path.exists(backup_filepath):
            return False, f"Backup file not found at {backup_filepath}"

        current_backup_hash = self.compute_sha256(backup_filepath)
        if current_backup_hash != expected_hash:
            return False, f"Backup hash mismatch! Expected {expected_hash}, got {current_backup_hash}"

        # Overwrite working copy with pristine backup
        shutil.copy2(backup_filepath, target_working_copy_path)

        restored_hash = self.compute_sha256(target_working_copy_path)
        if restored_hash != expected_hash:
            return False, "Failed to restore working copy accurately"

        return True, "Workbook successfully rolled back to pre-execution state"
