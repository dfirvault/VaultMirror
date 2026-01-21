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

# --- Global Paths ---
# Storing in AppData ensures the app works even if the EXE is moved
BASE_DIR = Path(os.environ.get('APPDATA')) / 'VaultMirror'
SCRIPTS_DIR = BASE_DIR / 'scripts'
STATES_DIR = BASE_DIR / 'sync-states'

# Ensure directories exist
for p in [BASE_DIR, SCRIPTS_DIR, STATES_DIR]:
    p.mkdir(parents=True, exist_ok=True)

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
    """Executes the sync logic using the bundled interpreter"""
    p = Path(script_path)
    if not p.exists():
        print(f"Error: Script path {script_path} not found.")
        return
    
    with open(p, 'r', encoding='utf-8') as f:
        code = f.read()
    
    # Execute with necessary imports available in the context
    exec_globals = {
        'os': os,
        'json': json,
        'shutil': shutil,
        'Path': Path,
        '__name__': '__main__'
    }
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
        
        sync_logic = f'''
import os
import json
import shutil
from pathlib import Path

def get_tree_state(path):
    p = Path(path)
    state = {{}}
    if not p.exists(): return state
    for file in p.rglob("*"):
        if file.is_file():
            try: state[str(file.relative_to(p))] = file.stat().st_mtime
            except: pass
    return state

def sync():
    dir_a, dir_b = Path(r"{source_path}"), Path(r"{dest_path}")
    state_path = Path(r"{state_file}")
    
    last_state = {{}}
    if state_path.exists():
        try:
            with open(state_path, "r") as f: last_state = json.load(f)
        except: pass

    curr_a, curr_b = get_tree_state(dir_a), get_tree_state(dir_b)
    all_paths = set(curr_a.keys()) | set(curr_b.keys()) | set(last_state.keys())
    new_state = {{}}

    for rel in all_paths:
        p_a, p_b = dir_a / rel, dir_b / rel
        in_a, in_b, in_l = rel in curr_a, rel in curr_b, rel in last_state

        if {bidirectional}:
            if in_l and not in_a and in_b:
                if p_b.exists(): os.remove(p_b)
                continue
            if in_l and not in_b and in_a:
                if p_a.exists(): os.remove(p_a)
                continue

        if in_a and (not in_b or curr_a[rel] > curr_b.get(rel, 0)):
            p_b.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(p_a, p_b)
            new_state[rel] = p_a.stat().st_mtime
        elif {bidirectional} and in_b and (not in_a or curr_b[rel] > curr_a.get(rel, 0)):
            p_a.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(p_b, p_a)
            new_state[rel] = p_b.stat().st_mtime
        elif in_a:
            new_state[rel] = curr_a[rel]

    with open(state_path, "w") as f:
        json.dump(new_state, f, indent=2)

if __name__ == "__main__":
    sync()
'''
        with open(script_path, 'w', encoding='utf-8') as f:
            f.write(sync_logic)
        return script_path

    def create_sync_task(self, case_name, source_path, dest_path, interval, bidirectional):
        task_name = f"dfirvault-sync-{case_name}"
        state_file = STATES_DIR / f"state_{task_name}.json"
        sync_script = self._create_sync_script(case_name, source_path, dest_path, bidirectional, state_file)
        
        intervals = {'1': ('MINUTE', '1'), '2': ('HOURLY', '1'), '3': ('DAILY', '1'), '4': ('WEEKLY', '1')}
        sch, mod = intervals.get(interval, ('HOURLY', '1'))
        
        exe_path = sys.executable
        # Building the command to call itself --run-task
        cmd = [
            'schtasks', '/Create', '/TN', task_name,
            '/TR', f'"{exe_path}" --run-task "{sync_script}"',
            '/SC', sch, '/MO', mod, '/F'
        ]
        
        res = subprocess.run(cmd, capture_output=True, text=True, shell=True)
        if res.returncode == 0:
            self.config['sync_jobs'][task_name] = {
                'source_path': str(source_path), 
                'dest_path': str(dest_path), 
                'bidirectional': bidirectional,
                'script_path': str(sync_script)
            }
            self.save_config()
            return True
        return False

    def delete_sync_task(self, task_name):
        """Fixed: defensive deletion logic to prevent KeyErrors"""
        # 1. Unregister from Task Scheduler
        subprocess.run(f'schtasks /Delete /TN "{task_name}" /F', shell=True, capture_output=True)
        
        # 2. Cleanup state file
        state_file = STATES_DIR / f"state_{task_name}.json"
        if state_file.exists():
            try: state_file.unlink()
            except: pass
        
        # 3. Cleanup logic script (with KeyCheck)
        details = self.config['sync_jobs'].get(task_name)
        if details and 'script_path' in details:
            p = Path(details['script_path'])
            if p.exists():
                try: p.unlink()
                except: pass

        # 4. Remove from config
        if task_name in self.config['sync_jobs']:
            del self.config['sync_jobs'][task_name]
            self.save_config()

    def run_sync_immediately(self, task_name):
        subprocess.run(f'schtasks /Run /TN "{task_name}"', shell=True, capture_output=True)

# --- UI ---

def clear():
    os.system('cls' if os.name == 'nt' else 'clear')

def main_menu():
    if not is_admin():
        print("ERROR: Administrative privileges required.")
        input("\nPress Enter to exit...")
        return
    
    scheduler = DriveSyncScheduler()
    
    while True:
        clear()
        print("==============================")
        print("    VAULT MIRROR MANAGER")
        print("==============================")
        print("1. Create New Sync Task")
        print("2. Manage Existing Tasks")
        print("3. Exit")
        print("------------------------------")
        
        choice = input("\nChoice: ").strip()
        
        if choice == '1':
            clear()
            case = input("Case Name: ").strip()
            if not case: continue
            src, dst = select_folder("Select Source"), select_folder("Select Destination")
            if not src or not dst: continue
            print("\n1. Min | 2. Hour | 3. Day | 4. Week")
            itv = input("Choice: ").strip()
            bi = input("Bi-directional? (y/n): ").lower() == 'y'
            if scheduler.create_sync_task(case, src, dst, itv, bi):
                print("\n✓ Task Created.")
            input("\nPress Enter...")
            
        elif choice == '2':
            clear()
            tasks = list(scheduler.config['sync_jobs'].keys())
            if not tasks:
                print("No tasks found."); input("\nPress Enter..."); continue
            for i, t in enumerate(tasks, 1): print(f"{i}. {t}")
            print(f"{len(tasks)+1}. Back")
            idx = input("\nSelect Task: ").strip()
            if idx.isdigit() and 1 <= int(idx) <= len(tasks):
                name = tasks[int(idx)-1]
                clear()
                print(f"TASK: {name}")
                print("1. Run Now | 2. Delete | 3. Back")
                sub = input("\nAction: ").strip()
                if sub == '1':
                    scheduler.run_sync_immediately(name)
                    print("✓ Triggered.")
                elif sub == '2':
                    scheduler.delete_sync_task(name)
                    print("✓ Deleted.")
                input("\nPress Enter to continue...")
        elif choice == '3':
            break

if __name__ == "__main__":
    if len(sys.argv) > 2 and sys.argv[1] == '--run-task':
        run_standalone_sync(sys.argv[2])
    else:
        main_menu()
