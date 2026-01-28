import os
import sys
import json
import shutil
import subprocess
import ctypes
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
from typing import List
import win32com.client
import time
from datetime import datetime, timedelta

# --- Global Paths ---
BASE_DIR = Path(os.environ.get('APPDATA')) / 'VaultMirror'
SCRIPTS_DIR = BASE_DIR / 'scripts'
STATES_DIR = BASE_DIR / 'sync-states'
LOCKS_DIR = BASE_DIR / 'locks'
# Note: Deleted folder will be on destination drive, not C: drive

for p in [BASE_DIR, SCRIPTS_DIR, STATES_DIR, LOCKS_DIR]:
    p.mkdir(parents=True, exist_ok=True)

# --- Constants ---
DELETION_GRACE_PERIOD_DAYS = 30  # Files in deleted folder older than this will be purged

# --- Helpers ---

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def select_folder(title="Select Folder"):
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    folder_selected = filedialog.askdirectory(title=title)
    root.destroy()
    return folder_selected

def run_standalone_sync(script_path):
    p = Path(script_path)
    if not p.exists(): return
    with open(p, 'r', encoding='utf-8') as f:
        code = f.read()
    exec_globals = {'os': os, 'json': json, 'shutil': shutil, 'Path': Path, '__name__': '__main__'}
    exec(code, exec_globals)

class DriveSyncScheduler:
    def __init__(self):
        try:
            self.scheduler = win32com.client.Dispatch('Schedule.Service')
            self.scheduler.Connect()
        except Exception as e:
            print(f"COM Connection Error: {e}")
        self.config_file = BASE_DIR / 'sync-config.json'
        self.load_config()
        
    def load_config(self):
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r') as f:
                    self.config = json.load(f)
            except:
                self.config = {'sync_jobs': {}}
        else:
            self.config = {'sync_jobs': {}}
    
    def save_config(self):
        with open(self.config_file, 'w') as f:
            json.dump(self.config, f, indent=2)

    def _create_sync_script(self, case_name, source_path, dest_path, bidirectional, state_file):
        script_path = SCRIPTS_DIR / f"sync_{case_name}.py"
        lock_file = LOCKS_DIR / f"{case_name}.lock"
        
        # Choose destination drive for deleted folder (use destination drive by default)
        dest_root = Path(dest_path).drive if Path(dest_path).drive else Path(source_path).drive
        DELETED_ROOT = Path(f"{dest_root}\\VaultMirror_Deleted\\{case_name}")
        
        # Build the script template with placeholders
        script_template = '''import os
import json
import shutil
import time
from pathlib import Path
from datetime import datetime, timedelta

# Files to ignore during sync
EXCLUSIONS = [".tmp"]
DELETION_GRACE_PERIOD_DAYS = 30  # Files in deleted folder older than this will be purged

# IMPORTANT: Exclude our own deleted folder from sync
DELETED_ROOT = Path(r"DELETED_ROOT_PLACEHOLDER")
EXCLUSION_PATHS = [DELETED_ROOT]

def is_drive_accessible(path):
    """Check if a drive/path is actually accessible"""
    try:
        # Try to list one item to test accessibility
        p = Path(path)
        if not p.exists():
            # Path doesn't exist - might be disconnected drive
            return False
        # Try to read from the path
        next(p.iterdir(), None)
        return True
    except (OSError, IOError, PermissionError, WindowsError):
        return False

def is_excluded_path(file_path):
    """Check if file is in an excluded path"""
    try:
        for excluded in EXCLUSION_PATHS:
            if excluded and excluded.exists():
                # Check if file is within the excluded directory
                if Path(file_path).is_relative_to(excluded):
                    return True
    except:
        pass
    return False

def get_tree_state(path):
    """Get current state of files in path, excluding our deleted folder"""
    p = Path(path)
    state = {}
    if not p.exists(): 
        return state
    
    for file in p.rglob("*"):
        if file.is_file():
            # Skip excluded file types
            if any(file.name.lower().endswith(ext) for ext in EXCLUSIONS):
                continue
            # Skip files in our deleted folder
            if is_excluded_path(file):
                continue
            try: 
                state[str(file.relative_to(p))] = {
                    'mtime': file.stat().st_mtime,
                    'size': file.stat().st_size
                }
            except: 
                pass
    return state

def safe_delete(file_path, deleted_root, sync_id, direction):
    """Move file to deleted folder instead of permanent deletion"""
    try:
        # Create deleted folder if it doesn't exist
        deleted_root.mkdir(parents=True, exist_ok=True)
        
        # Create unique deletion timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Get relative path for organization
        try:
            # Try to get relative path from source/dest root
            if "A_to_B" in direction:
                root_path = Path(r"SOURCE_PATH_PLACEHOLDER")
            elif "B_to_A" in direction:
                root_path = Path(r"DEST_PATH_PLACEHOLDER")
            else:
                root_path = file_path.parents[-1]
            
            try:
                rel_path = str(file_path.relative_to(root_path))
            except:
                rel_path = file_path.name
        except:
            rel_path = file_path.name
        
        # Create safe filename for storage
        safe_name = rel_path.replace(os.sep, "_").replace("..", "parent")
        if len(safe_name) > 200:
            safe_name = safe_name[:100] + "..." + safe_name[-100:]
        
        # Create organized deletion folder structure
        deleted_dir = deleted_root / direction / timestamp[:8]  # YYYYMMDD
        deleted_dir.mkdir(parents=True, exist_ok=True)
        
        # Move file to deleted folder
        dest_path = deleted_dir / f"{timestamp}_{safe_name}"
        
        # If destination exists, add counter
        counter = 1
        original_dest = dest_path
        while dest_path.exists():
            dest_path = original_dest.with_stem(f"{original_dest.stem}_{counter}")
            counter += 1
            
        shutil.move(str(file_path), str(dest_path))
        
        # Create metadata file
        meta = {
            'original_path': str(file_path),
            'original_rel_path': rel_path,
            'deleted_at': timestamp,
            'sync_id': sync_id,
            'direction': direction,
            'original_size': file_path.stat().st_size if file_path.exists() else 0
        }
        
        with open(f"{dest_path}.meta.json", 'w') as f:
            json.dump(meta, f, indent=2)
            
        return True
    except Exception as e:
        print(f"Safe delete failed for {file_path}: {e}")
        return False

def purge_old_deletions(deleted_root, days_old=DELETION_GRACE_PERIOD_DAYS):
    """Purge files in deleted folder older than specified days"""
    if not deleted_root.exists():
        return 0
        
    cutoff_time = time.time() - (days_old * 24 * 60 * 60)
    purged_count = 0
    
    for meta_file in deleted_root.rglob("*.meta.json"):
        try:
            # Check meta file age
            if meta_file.stat().st_mtime < cutoff_time:
                # Find associated data file
                data_file = meta_file.with_suffix('')  # Remove .meta.json
                if data_file.exists():
                    data_file.unlink()
                meta_file.unlink()
                purged_count += 1
                
                # Try to remove empty directories
                try:
                    meta_file.parent.rmdir()
                except:
                    pass  # Directory not empty
        except:
            continue
    
    return purged_count

def sync():
    lock_path = Path(r"LOCK_FILE_PLACEHOLDER")
    if lock_path.exists(): 
        print("Sync already in progress")
        return
    
    lock_path.touch()
    
    try:
        dir_a, dir_b = Path(r"SOURCE_PATH_PLACEHOLDER"), Path(r"DEST_PATH_PLACEHOLDER")
        state_path = Path(r"STATE_FILE_PLACEHOLDER")
        sync_id = "CASE_NAME_PLACEHOLDER"
        
        # Initialize deleted folder on destination drive
        DELETED_ROOT.mkdir(parents=True, exist_ok=True)
        
        print(f"Deleted files stored at: {DELETED_ROOT}")
        
        # Check drive accessibility
        a_accessible = is_drive_accessible(dir_a)
        b_accessible = is_drive_accessible(dir_b)
        
        if not a_accessible and not b_accessible:
            print(f"ERROR: Both drives inaccessible. Skipping sync.")
            return
            
        if not a_accessible:
            print(f"WARNING: Source drive {dir_a} is inaccessible. Only copying from B to A if bidirectional.")
            
        if not b_accessible:
            print(f"WARNING: Destination drive {dir_b} is inaccessible. Only copying from A to B if bidirectional.")
        
        # Purge old deletions before sync
        purged = purge_old_deletions(DELETED_ROOT, DELETION_GRACE_PERIOD_DAYS)
        if purged > 0:
            print(f"Purged {purged} old deleted files")
        
        last_state = {}
        if state_path.exists():
            try:
                with open(state_path, "r") as f: 
                    last_state = json.load(f)
            except: 
                pass

        # Only scan accessible drives
        curr_a = get_tree_state(dir_a) if a_accessible else {}
        curr_b = get_tree_state(dir_b) if b_accessible else {}
        
        all_paths = set(curr_a.keys()) | set(curr_b.keys()) | set(last_state.keys())
        new_state = {}
        deletions = 0
        
        for rel in all_paths:
            p_a, p_b = dir_a / rel, dir_b / rel
            in_a, in_b, in_l = rel in curr_a, rel in curr_b, rel in last_state
            
            # Skip if path is in our deleted folder (shouldn't happen with exclusion, but safety check)
            if is_excluded_path(p_a) or is_excluded_path(p_b):
                continue
'''
        
        # Add the sync logic based on bidirectional flag
        if bidirectional:
            script_template += '''
            # Bi-directional deletion logic with safety checks
            if in_l and not in_a and in_b:
                # File existed before, now only in B (deleted from A)
                if b_accessible and p_b.exists():
                    if safe_delete(p_b, DELETED_ROOT, sync_id, "A_to_B"):
                        deletions += 1
                    continue
            elif in_l and not in_b and in_a:
                # File existed before, now only in A (deleted from B)
                if a_accessible and p_a.exists():
                    if safe_delete(p_a, DELETED_ROOT, sync_id, "B_to_A"):
                        deletions += 1
                    continue
            '''
        else:
            script_template += '''
            # One-way deletion: only delete from destination if source doesn't have it
            if in_l and not in_a and in_b:
                # File existed before in source, now missing from source but in destination
                if b_accessible and p_b.exists():
                    if safe_delete(p_b, DELETED_ROOT, sync_id, "one_way"):
                        deletions += 1
                    continue
            '''
        
        # Add copy logic for both directions
        script_template += '''
            # Copy from A to B if A is accessible
            if in_a and a_accessible:
                # Copy from A to B if B is accessible
                if b_accessible and (not in_b or curr_a[rel]['mtime'] > curr_b.get(rel, {}).get('mtime', 0)):
                    p_b.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(p_a, p_b)
                    new_state[rel] = {'mtime': curr_a[rel]['mtime'], 'size': curr_a[rel]['size']}
                elif not b_accessible and in_a:
                    # B not accessible, but A has file - keep in state
                    new_state[rel] = {'mtime': curr_a[rel]['mtime'], 'size': curr_a[rel]['size']}
        '''
        
        if bidirectional:
            script_template += '''
            # Bi-directional: copy from B to A if B is accessible
            elif in_b and b_accessible:
                if a_accessible and (not in_a or curr_b[rel]['mtime'] > curr_a.get(rel, {}).get('mtime', 0)):
                    p_a.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(p_b, p_a)
                    new_state[rel] = {'mtime': curr_b[rel]['mtime'], 'size': curr_b[rel]['size']}
                elif not a_accessible and in_b:
                    # A not accessible, but B has file - keep in state
                    new_state[rel] = {'mtime': curr_b[rel]['mtime'], 'size': curr_b[rel]['size']}
            '''
        
        # Add the rest of the sync function
        script_template += '''
        
        with open(state_path, "w") as f:
            json.dump(new_state, f, indent=2)
            
        if deletions > 0:
            print(f"SAFE DELETE: Moved {deletions} file(s) to {DELETED_ROOT}")
            print(f"Files will be permanently deleted after ''' + str(DELETION_GRACE_PERIOD_DAYS) + ''' days.")
            
    except Exception as e:
        print(f"Sync error: {e}")
    finally:
        if lock_path.exists(): 
            lock_path.unlink()

if __name__ == "__main__":
    sync()
'''
        
        # Replace placeholders with actual values
        script_content = script_template
        script_content = script_content.replace("DELETED_ROOT_PLACEHOLDER", str(DELETED_ROOT))
        script_content = script_content.replace("SOURCE_PATH_PLACEHOLDER", str(source_path))
        script_content = script_content.replace("DEST_PATH_PLACEHOLDER", str(dest_path))
        script_content = script_content.replace("LOCK_FILE_PLACEHOLDER", str(lock_file))
        script_content = script_content.replace("STATE_FILE_PLACEHOLDER", str(state_file))
        script_content = script_content.replace("CASE_NAME_PLACEHOLDER", case_name)
        
        with open(script_path, 'w', encoding='utf-8') as f:
            f.write(script_content)
        return script_path

    def create_sync_task(self, case_name, source_path, dest_path, interval, bidirectional):
        task_name = f"dfirvault-sync-{case_name}"
        state_file = STATES_DIR / f"state_{task_name}.json"
        sync_script = self._create_sync_script(case_name, source_path, dest_path, bidirectional, state_file)
        
        # Display warning for bidirectional sync
        if bidirectional:
            # Calculate deleted folder location
            dest_root = Path(dest_path).drive if Path(dest_path).drive else Path(source_path).drive
            deleted_location = f"{dest_root}\\VaultMirror_Deleted\\{case_name}"
            
            print("\n" + "="*60)
            print("⚠️  BIDIRECTIONAL SYNC WARNING")
            print("="*60)
            print("Bidirectional sync will propagate DELETIONS between drives.")
            print("Deleted files will be moved to the recycle folder for 30 days.")
            print(f"Recycle location: {deleted_location}")
            print("="*60)
            input("Press Enter to acknowledge and continue...")
        
        interval_map = {'1': ('MINUTE', '1', 'Every Minute'), '2': ('HOURLY', '1', 'Hourly'), 
                        '3': ('DAILY', '1', 'Daily'), '4': ('WEEKLY', '1', 'Weekly')}
        sch, mod, friendly_name = interval_map.get(interval, ('HOURLY', '1', 'Hourly'))
        
        exe_path = sys.executable
        cmd = ['schtasks', '/Create', '/TN', task_name, '/TR', f'"{exe_path}" --run-task "{sync_script}"',
               '/SC', sch, '/MO', mod, '/F']
        
        res = subprocess.run(cmd, capture_output=True, text=True, shell=True)
        if res.returncode == 0:
            # Calculate deleted folder location for display
            dest_root = Path(dest_path).drive if Path(dest_path).drive else Path(source_path).drive
            deleted_location = f"{dest_root}\\VaultMirror_Deleted\\{case_name}"
            
            self.config['sync_jobs'][task_name] = {
                'case_name': case_name, 
                'source_path': str(source_path), 
                'dest_path': str(dest_path), 
                'bidirectional': bidirectional, 
                'interval_desc': friendly_name, 
                'script_path': str(sync_script),
                'deleted_location': deleted_location
            }
            self.save_config()
            return True
        return False

    def delete_sync_task(self, task_name):
        subprocess.run(f'schtasks /Delete /TN "{task_name}" /F', shell=True, capture_output=True)
        state_file = STATES_DIR / f"state_{task_name}.json"
        if state_file.exists(): state_file.unlink()
        details = self.config['sync_jobs'].get(task_name)
        if details:
            if 'script_path' in details:
                p = Path(details['script_path'])
                if p.exists(): p.unlink()
            l = LOCKS_DIR / f"{details.get('case_name', '')}.lock"
            if l.exists(): l.unlink()
        if task_name in self.config['sync_jobs']:
            del self.config['sync_jobs'][task_name]
            self.save_config()

    def run_sync_immediately(self, task_name):
        subprocess.run(f'schtasks /Run /TN "{task_name}"', shell=True, capture_output=True)

# --- UI ---

def clear():
    os.system('cls' if os.name == 'nt' else 'clear')

def view_deleted_folder(deleted_path):
    """View contents of a specific deleted folder"""
    if not deleted_path.exists():
        print(f"\nFolder not found: {deleted_path}")
        input("Press Enter to continue...")
        return
    
    clear()
    print(f"Deleted Files in: {deleted_path}")
    print("="*60)
    
    total_size = 0
    total_files = 0
    
    for meta_file in deleted_path.rglob("*.meta.json"):
        try:
            with open(meta_file, 'r') as f:
                meta = json.load(f)
            total_files += 1
            total_size += meta.get('original_size', 0)
            
            print(f"\n{total_files}. {meta.get('original_rel_path', 'Unknown')}")
            print(f"   Deleted: {meta.get('deleted_at', 'Unknown')}")
            print(f"   Direction: {meta.get('direction', 'Unknown')}")
            print(f"   Size: {meta.get('original_size', 0):,} bytes")
            print(f"   Meta: {meta_file}")
        except:
            continue
    
    if total_files == 0:
        print("\nNo deleted files found in this folder.")
    else:
        print(f"\n\nTotal: {total_files:,} files | Total size: {total_size:,} bytes")
        
        print("\nOptions:")
        print("1. Purge files older than 30 days")
        print("2. Back")
        
        choice = input("\nChoice: ").strip()
        if choice == '1':
            # Manual purge
            cutoff_time = time.time() - (DELETION_GRACE_PERIOD_DAYS * 24 * 60 * 60)
            purged = 0
            for meta_file in deleted_path.rglob("*.meta.json"):
                if meta_file.stat().st_mtime < cutoff_time:
                    data_file = meta_file.with_suffix('')
                    if data_file.exists():
                        data_file.unlink()
                    meta_file.unlink()
                    purged += 1
            print(f"\nPurged {purged} files.")
            input("Press Enter to continue...")
    
    input("\nPress Enter to continue...")

def show_deleted_files():
    """Show files in the deleted folder"""
    clear()
    print("="*60)
    print("DELETED FILES MANAGEMENT")
    print("="*60)
    
    print("Deleted files are stored on the destination drive of each sync job.")
    print("To view deleted files, check the sync destination drive for:")
    print("  Drive:\\VaultMirror_Deleted\\[CaseName]\\")
    print("\nExample: If syncing D:\\Data to E:\\Backup, check:")
    print("  E:\\VaultMirror_Deleted\\[YourCaseName]\\")
    
    print("\nOptions:")
    print("1. Enter path to deleted folder manually")
    print("2. Back to main menu")
    
    choice = input("\nChoice: ").strip()
    if choice == '1':
        folder = select_folder("Select Deleted Folder Location")
        if folder:
            view_deleted_folder(Path(folder))

def main_menu():
    if not is_admin():
        print("ERROR: Administrative privileges required.")
        input("\nPress Enter to exit...")
        return
    scheduler = DriveSyncScheduler()
    while True:
        clear()
        print("="*60)
        print("        VAULT MIRROR MANAGER (SAFE SYNC)")
        print("="*60)
        print("1. Create New Sync Task")
        print("2. Manage Existing Tasks")
        print("3. View/Manage Deleted Files")
        print("4. Exit")
        print("\n⚠️  SAFE DELETE ENABLED:")
        print("   • Files are NEVER permanently deleted immediately")
        print("   • Deleted files moved to: [DestinationDrive]:\\VaultMirror_Deleted\\")
        print(f"   • Files purged after {DELETION_GRACE_PERIOD_DAYS} days")
        print("   • Deleted folder is EXCLUDED from sync")
        print("="*60)
        choice = input("\nChoice: ").strip()
        
        if choice == '1':
            clear()
            case = input("Case Name: ").strip()
            if not case: continue
            src, dst = select_folder("Select Source"), select_folder("Select Destination")
            if not src or not dst: continue
            print("\n1. Minute | 2. Hour | 3. Day | 4. Week")
            itv = input("Choice: ").strip()
            bi = input("Bi-directional? (y/n): ").lower() == 'y'
            if scheduler.create_sync_task(case, src, dst, itv, bi):
                print("\n✓ Task Created.")
                
                # Show deleted folder location
                dest_root = Path(dst).drive if Path(dst).drive else Path(src).drive
                deleted_location = f"{dest_root}\\VaultMirror_Deleted\\{case}"
                print(f"✓ Deleted files will be stored at: {deleted_location}")
                print("✓ This folder is automatically excluded from sync")
            input("\nPress Enter...")
            
        elif choice == '2':
            clear()
            tasks = list(scheduler.config['sync_jobs'].keys())
            if not tasks:
                print("No tasks found."); input("Press Enter..."); continue
            for i, t in enumerate(tasks, 1): print(f"{i}. {t}")
            print(f"{len(tasks)+1}. Back")
            idx = input("\nSelect Task: ").strip()
            if idx.isdigit() and 1 <= int(idx) <= len(tasks):
                name = tasks[int(idx)-1]
                details = scheduler.config['sync_jobs'][name]
                clear()
                print(f"--- TASK DETAILS: {name} ---")
                print(f"Source:   {details.get('source_path')}")
                print(f"Dest:     {details.get('dest_path')}")
                print(f"Interval: {details.get('interval_desc', 'Unknown')}")
                print(f"Mode:     {'Bi-Directional (Safe Delete)' if details.get('bidirectional') else 'One-Way (Safe Delete)'}")
                print(f"Deleted files location: {details.get('deleted_location', 'Unknown')}")
                print("-" * 60)
                print("1. Run Now")
                print("2. Delete Task")
                print("3. Back")
                sub = input("\nAction: ").strip()
                if sub == '1':
                    scheduler.run_sync_immediately(name)
                    print("✓ Triggered."); input("Press Enter...")
                elif sub == '2':
                    scheduler.delete_sync_task(name)
                    print("✓ Deleted."); input("Press Enter...")
        elif choice == '3':
            show_deleted_files()
        elif choice == '4':
            break

if __name__ == "__main__":
    if len(sys.argv) > 2 and sys.argv[1] == '--run-task':
        run_standalone_sync(sys.argv[2])
    else:
        main_menu()
