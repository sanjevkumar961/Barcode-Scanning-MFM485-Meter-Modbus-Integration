import sys, os, time, shutil, subprocess, traceback

new_exe = sys.argv[1]
current_exe = sys.argv[2]

base_dir = os.path.dirname(current_exe)

backup_dir = os.path.join(base_dir, "backup")
log_dir = os.path.join(base_dir, "logs")

os.makedirs(backup_dir, exist_ok=True)
os.makedirs(log_dir, exist_ok=True)

log_file = os.path.join(log_dir, "updater.log")

def log(msg):
    with open(log_file, "a") as f:
        f.write(msg + "\n")

backup_exe = os.path.join(backup_dir, "MotorNoLoadTester_backup.exe")
old_exe = current_exe + ".old"

try:
    log("Updater started")
    time.sleep(2)  # allow main app to exit

    # Cleanup from previous failed updates
    if os.path.exists(old_exe):
        os.remove(old_exe)

    # Step 1: Backup current exe
    log("Creating backup")
    shutil.copy2(current_exe, backup_exe)

    # Step 2: Rename current exe (atomic)
    log("Renaming current exe")
    os.rename(current_exe, old_exe)

    # Step 3: Move new exe into place
    log("Installing new exe")
    shutil.move(new_exe, current_exe)

    # Step 4: Start updated app
    log("Starting updated application")
    subprocess.Popen([current_exe])

    # Step 5: Cleanup old exe
    if os.path.exists(old_exe):
        os.remove(old_exe)

    log("Update successful")

except Exception as e:
    log("Update failed")
    log(str(e))
    log(traceback.format_exc())

    # Rollback logic
    try:
        log("Attempting rollback")

        if os.path.exists(current_exe):
            os.remove(current_exe)

        if os.path.exists(old_exe):
            os.rename(old_exe, current_exe)
        elif os.path.exists(backup_exe):
            shutil.copy2(backup_exe, current_exe)

        subprocess.Popen([current_exe])
        log("Rollback successful")

    except Exception as rollback_error:
        log("Rollback FAILED")
        log(str(rollback_error))
