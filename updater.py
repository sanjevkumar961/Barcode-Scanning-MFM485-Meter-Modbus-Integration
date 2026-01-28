"""
Automatic updater for MotorNoLoadTester application.

Usage: python updater.py <new_exe_path> <current_exe_path>

This script safely updates the application by:
1. Validating input arguments and file integrity
2. Creating timestamped backups
3. Replacing the executable atomically
4. Rolling back on failure
5. Starting the updated application
"""

import sys
import os
import time
import shutil
import subprocess
import traceback
import hashlib
import logging
from datetime import datetime
from typing import Optional, Tuple
from pathlib import Path

# Constants
APP_EXIT_WAIT_TIME = 3  # seconds to allow app to exit
BACKUP_RETENTION_COUNT = 3  # keep last N backups
LOG_ROTATION_SIZE = 5 * 1024 * 1024  # 5MB
MAX_RETRIES = 3
RETRY_DELAY = 0.5  # seconds between retries

# Derived paths
def setup_paths(current_exe: str) -> Tuple[str, str, str]:
    """Setup and create necessary directories."""
    base_dir = os.path.dirname(current_exe)
    backup_dir = os.path.join(base_dir, "backup")
    log_dir = os.path.join(base_dir, "logs")
    
    try:
        os.makedirs(backup_dir, exist_ok=True)
        os.makedirs(log_dir, exist_ok=True)
        return backup_dir, log_dir, base_dir
    except OSError as e:
        print(f"ERROR: Cannot create directories: {e}")
        sys.exit(1)

def setup_logging(log_dir: str) -> logging.Logger:
    """Setup logging with rotation and proper formatting."""
    log_file = os.path.join(log_dir, "updater.log")
    
    logger = logging.getLogger("updater")
    logger.setLevel(logging.DEBUG)
    
    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    
    # File handler with rotation
    try:
        file_handler = logging.handlers.RotatingFileHandler(
            log_file,
            maxBytes=LOG_ROTATION_SIZE,
            backupCount=5
        )
        file_handler.setLevel(logging.DEBUG)
    except Exception as e:
        print(f"WARNING: Cannot setup file logging: {e}")
        file_handler = None
    
    # Formatter
    formatter = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    if file_handler:
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    
    return logger

def validate_arguments() -> Tuple[str, str]:
    """Validate and parse command line arguments."""
    if len(sys.argv) < 3:
        print("ERROR: Missing required arguments")
        print("Usage: python updater.py <new_exe_path> <current_exe_path>")
        sys.exit(1)
    
    new_exe = sys.argv[1]
    current_exe = sys.argv[2]
    
    # Validate paths
    if not os.path.isabs(new_exe):
        print(f"ERROR: new_exe path must be absolute: {new_exe}")
        sys.exit(1)
    
    if not os.path.isabs(current_exe):
        print(f"ERROR: current_exe path must be absolute: {current_exe}")
        sys.exit(1)
    
    if not os.path.exists(new_exe):
        print(f"ERROR: new_exe does not exist: {new_exe}")
        sys.exit(1)
    
    if not os.path.isfile(new_exe):
        print(f"ERROR: new_exe is not a file: {new_exe}")
        sys.exit(1)
    
    return new_exe, current_exe

def check_file_permissions(directory: str, logger: logging.Logger) -> bool:
    """Check if we have write permissions to target directory."""
    try:
        test_file = os.path.join(directory, ".permission_test")
        with open(test_file, "w") as f:
            f.write("test")
        os.remove(test_file)
        return True
    except (OSError, PermissionError) as e:
        logger.error(f"Permission denied for directory {directory}: {e}")
        return False

def calculate_file_hash(filepath: str) -> str:
    """Calculate SHA256 hash of file for integrity verification."""
    sha256_hash = hashlib.sha256()
    try:
        with open(filepath, "rb") as f:
            for byte_block in iter(lambda: f.read(4096), b""):
                sha256_hash.update(byte_block)
        return sha256_hash.hexdigest()
    except Exception as e:
        raise IOError(f"Cannot read file for hashing: {e}")

def verify_file_integrity(new_exe: str, logger: logging.Logger) -> bool:
    """Verify new executable is valid before deployment."""
    try:
        # Check file size (executable should be > 1MB)
        file_size = os.path.getsize(new_exe)
        if file_size < 1024 * 100:  # Less than 100KB
            logger.error(f"File size suspiciously small: {file_size} bytes")
            return False
        
        logger.info(f"New executable size: {file_size} bytes")
        
        # Calculate hash for logging
        file_hash = calculate_file_hash(new_exe)
        logger.info(f"New executable hash: {file_hash}")
        
        return True
    except Exception as e:
        logger.error(f"File integrity check failed: {e}")
        return False

def wait_for_app_exit(wait_time: int, logger: logging.Logger) -> None:
    """Wait for application to fully exit with logging."""
    logger.info(f"Waiting {wait_time} seconds for application to exit...")
    for i in range(wait_time, 0, -1):
        time.sleep(1)
        if i % wait_time == 0 or i == 1:
            logger.debug(f"Waiting... {i} seconds remaining")

def cleanup_old_backup(backup_dir: str, logger: logging.Logger) -> None:
    """Clean up old backups, keeping only BACKUP_RETENTION_COUNT."""
    try:
        backups = sorted(
            [f for f in os.listdir(backup_dir) if f.endswith(".exe")],
            reverse=True
        )
        
        for old_backup in backups[BACKUP_RETENTION_COUNT:]:
            backup_path = os.path.join(backup_dir, old_backup)
            try:
                os.remove(backup_path)
                logger.info(f"Removed old backup: {old_backup}")
            except OSError as e:
                logger.warning(f"Cannot remove backup {old_backup}: {e}")
    except Exception as e:
        logger.warning(f"Backup cleanup failed: {e}")

def safe_file_operation(operation, *args, max_retries: int = MAX_RETRIES, 
                        retry_delay: float = RETRY_DELAY, logger: Optional[logging.Logger] = None):
    """Safely execute file operation with retries for Windows file locking."""
    for attempt in range(max_retries):
        try:
            return operation(*args)
        except (OSError, PermissionError) as e:
            if attempt < max_retries - 1:
                if logger:
                    logger.debug(f"Retry {attempt + 1}/{max_retries} after {retry_delay}s: {e}")
                time.sleep(retry_delay)
            else:
                raise

def backup_current_exe(current_exe: str, backup_dir: str, logger: logging.Logger) -> str:
    """Create timestamped backup of current executable."""
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_exe = os.path.join(backup_dir, f"MotorNoLoadTester_backup_{timestamp}.exe")
        
        logger.info(f"Creating backup: {backup_exe}")
        safe_file_operation(shutil.copy2, current_exe, backup_exe, 
                           max_retries=MAX_RETRIES, logger=logger)
        
        logger.info("Backup created successfully")
        return backup_exe
    except Exception as e:
        logger.error(f"Failed to create backup: {e}")
        raise

def perform_update(new_exe: str, current_exe: str, logger: logging.Logger) -> bool:
    """Perform the actual update operation."""
    old_exe = current_exe + ".old"
    
    try:
        # Step 1: Cleanup from previous failed updates
        if os.path.exists(old_exe):
            try:
                logger.debug(f"Removing old temporary file: {old_exe}")
                safe_file_operation(os.remove, old_exe, logger=logger)
            except Exception as e:
                logger.warning(f"Cannot remove old temporary file: {e}")
        
        # Step 2: Rename current exe (atomic operation)
        logger.info("Renaming current executable")
        safe_file_operation(os.rename, current_exe, old_exe, logger=logger)
        
        # Step 3: Move new exe into place
        logger.info("Installing new executable")
        safe_file_operation(shutil.move, new_exe, current_exe, logger=logger)
        
        # Step 4: Verify installation
        if not os.path.exists(current_exe):
            raise IOError("New executable not found after move operation")
        
        logger.info("Update completed successfully")
        return True
        
    except Exception as e:
        logger.error(f"Update failed: {e}")
        raise

def rollback_update(old_exe: str, backup_exe: str, current_exe: str, 
                   logger: logging.Logger) -> bool:
    """Attempt to rollback to previous version."""
    rollback_success = False
    
    try:
        logger.warning("Attempting rollback to previous version")
        
        # Remove corrupted new exe if it exists
        if os.path.exists(current_exe):
            try:
                logger.debug("Removing corrupted executable")
                safe_file_operation(os.remove, current_exe, logger=logger)
            except Exception as e:
                logger.error(f"Cannot remove corrupted file: {e}")
        
        # Try to restore from old_exe first (most recent)
        if os.path.exists(old_exe):
            try:
                logger.info("Restoring from recent backup (.old file)")
                safe_file_operation(os.rename, old_exe, current_exe, logger=logger)
                rollback_success = True
            except Exception as e:
                logger.warning(f"Cannot restore from .old file: {e}")
        
        # Fallback to timestamped backup
        if not rollback_success and os.path.exists(backup_exe):
            try:
                logger.info("Restoring from timestamped backup")
                safe_file_operation(shutil.copy2, backup_exe, current_exe, logger=logger)
                rollback_success = True
            except Exception as e:
                logger.error(f"Cannot restore from backup: {e}")
        
        if rollback_success:
            logger.warning("Rollback completed successfully")
            return True
        else:
            logger.critical("Rollback failed - no backup available")
            return False
            
    except Exception as e:
        logger.critical(f"Rollback operation failed: {e}")
        return False

def start_application(current_exe: str, logger: logging.Logger) -> None:
    """Start the application after successful update/rollback."""
    try:
        logger.info(f"Starting application: {current_exe}")
        subprocess.Popen([current_exe])
    except Exception as e:
        logger.error(f"Failed to start application: {e}")
        raise

def main() -> int:
    """Main updater function."""
    # Validate arguments first
    try:
        new_exe, current_exe = validate_arguments()
    except SystemExit:
        return 1
    
    # Setup paths and logging
    backup_dir, log_dir, base_dir = setup_paths(current_exe)
    logger = setup_logging(log_dir)
    
    logger.info("=" * 60)
    logger.info("UPDATER STARTED")
    logger.info(f"Current executable: {current_exe}")
    logger.info(f"New executable: {new_exe}")
    logger.info("=" * 60)
    
    backup_exe = None
    old_exe = current_exe + ".old"
    
    try:
        # Pre-flight checks
        if not check_file_permissions(base_dir, logger):
            logger.error("Insufficient permissions to update application")
            return 1
        
        if not verify_file_integrity(new_exe, logger):
            logger.error("New executable failed integrity checks")
            return 1
        
        if not os.path.exists(current_exe):
            logger.error(f"Current executable not found: {current_exe}")
            return 1
        
        # Wait for application to exit
        wait_for_app_exit(APP_EXIT_WAIT_TIME, logger)
        
        # Backup current executable
        backup_exe = backup_current_exe(current_exe, backup_dir, logger)
        
        # Perform update
        if perform_update(new_exe, current_exe, logger):
            # Cleanup old exe after successful update
            if os.path.exists(old_exe):
                try:
                    safe_file_operation(os.remove, old_exe, logger=logger)
                except Exception as e:
                    logger.warning(f"Cannot cleanup temporary file: {e}")
            
            # Cleanup old backups
            cleanup_old_backup(backup_dir, logger)
            
            # Start updated application
            start_application(current_exe, logger)
            
            logger.info("=" * 60)
            logger.info("UPDATE COMPLETED SUCCESSFULLY")
            logger.info("=" * 60)
            return 0
        
    except Exception as e:
        logger.error(f"Update failed with exception: {e}")
        logger.error(traceback.format_exc())
        
        # Attempt rollback
        if not rollback_update(old_exe, backup_exe or "", current_exe, logger):
            logger.critical("CRITICAL: Update failed and rollback also failed!")
            logger.critical("Application may be in an inconsistent state")
            return 2
        
        # Start rolled-back application
        try:
            start_application(current_exe, logger)
            logger.warning("=" * 60)
            logger.warning("ROLLBACK COMPLETED - RUNNING PREVIOUS VERSION")
            logger.warning("=" * 60)
            return 1
        except Exception as start_error:
            logger.critical(f"Failed to start rolled-back application: {start_error}")
            return 2
    
    return 0

if __name__ == "__main__":
    import logging.handlers
    sys.exit(main())
