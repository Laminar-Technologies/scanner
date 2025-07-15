# asana_auto_main.py (v1.5.1)
import logging
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import time
import threading
import json
from playsound import playsound

from asana_api_client import AsanaClient
from asana_operations import (
    process_heater_board_swap, process_cor_operation, generate_cal_cert,
    process_device_complete, process_custom_operation, _get_resource_path
)

APP_CONFIG = { "PROJECT_NAME": "AMAT AGS", "HEATER_SWAP_TAG_NAME": "Heater Board Replacement", "COR_TAG_NAME": "Return Unrepaired", "DEVICE_COMPLETE_TAG_NAME": "Device Calibrated", "ACCOUNT_MANAGER_ASSIGNEE_NAME": "Mandy McIntosh", "SHARED_SUBTASK_ASSIGNEE_NAME": "Michelle Hughes", "NEEDS_COR_SECTION_NAME": "Needs COR", "READY_FOR_BUYER_SECTION_NAME": "Ready for Buyer", "PURGE_SECTION_NAME": "PURGE", "COR_REASON_TAGS": { "Bad Sensor": "Bad Sensor", "Pressure Oscillates": "Pressure Oscillation", "INTERNAL LEAK": "INTERNAL LEAK" } }

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', filename='asana_automation.log', filemode='a')
ASANA_TOKEN = "2/1209815337657407/1210546925916369:b528afe6d6f2b4ea1cbca2096b9bcb13"

# --- Dialog Classes (Unchanged) ---
class CorReasonDialog(simpledialog.Dialog):
    def __init__(self, parent, title=None): self.result = None; super().__init__(parent, title)
    def body(self, master):
        self.resizable(False, False); ttk.Label(master, text="Select a reason:").grid(row=0, sticky='w', padx=5, pady=5)
        reasons = list(APP_CONFIG["COR_REASON_TAGS"].keys()) + ["Custom"]; self.reason_var = tk.StringVar(value=reasons[0])
        self.reason_menu = ttk.Combobox(master, textvariable=self.reason_var, values=reasons, state="readonly"); self.reason_menu.grid(row=1, padx=5, pady=5); self.reason_menu.bind("<<ComboboxSelected>>", self.toggle_custom_entry)
        ttk.Label(master, text="If custom, enter reason:").grid(row=2, sticky='w', padx=5, pady=5); self.custom_entry = ttk.Entry(master, width=40); self.custom_entry.grid(row=3, padx=5, pady=5)
        self.toggle_custom_entry(); return self.reason_menu
    def toggle_custom_entry(self, event=None):
        if self.reason_var.get() == "Custom": self.custom_entry.config(state="normal"); self.custom_entry.focus_set()
        else: self.custom_entry.config(state="disabled")
    def apply(self):
        reason = self.reason_var.get()
        if reason == "Custom": self.result = {"reason": self.custom_entry.get() or "No custom reason provided."}
        else: self.result = {"reason": reason}

class CustomActionDialog(simpledialog.Dialog):
    def body(self, master):
        self.title("Build Custom Action"); self.actions = []; self.action_var = tk.StringVar()
        action_frame = ttk.Frame(master); action_frame.pack(pady=5, padx=5, fill='x')
        ttk.Label(action_frame, text="Action:").pack(side='left'); self.action_menu = ttk.Combobox(action_frame, textvariable=self.action_var, values=["Add Tag", "Assign To", "Move to Section", "Add Comment"], state="readonly"); self.action_menu.pack(side='left', padx=5)
        value_frame = ttk.Frame(master); value_frame.pack(pady=5, padx=5, fill='x')
        ttk.Label(value_frame, text="Value (Name or GID):").pack(side='left'); self.value_entry = ttk.Entry(value_frame, width=30); self.value_entry.pack(side='left', padx=5)
        self.value_entry.bind("<Return>", self.add_action)
        self.add_button = ttk.Button(master, text="Add Action", command=self.add_action); self.add_button.pack(pady=5)
        self.action_listbox = tk.Listbox(master, height=5, width=50); self.action_listbox.pack(pady=5, padx=5, fill='x', expand=True)
        return self.action_menu
    def add_action(self, event=None):
        action_type = self.action_var.get(); value = self.value_entry.get()
        if not action_type or not value: messagebox.showwarning("Input Error", "Please select an action and provide a value.", parent=self); return
        action_map = {"Add Tag": "add_tag", "Assign To": "assign_to", "Move to Section": "move_to", "Add Comment": "add_comment"}
        self.actions.append({'type': action_map[action_type], 'value': value}); self.action_listbox.insert(tk.END, f"{action_type}: {value}")
        self.value_entry.delete(0, tk.END); self.value_entry.focus_set()
    def apply(self): self.result = self.actions

class AsanaAutomationApp:
    # ... (__init__ setup is unchanged) ...
    def __init__(self, master):
        self.master = master; master.title("Asana Task Automation"); master.geometry("800x600"); master.minsize(600, 500); master.attributes('-topmost', True)
        self.selected_operation = None; self.custom_actions = []; self.last_activity_time = time.time(); self.inactivity_check_interval_ms = 60000; self.inactivity_timeout_ms = 3600000; self.is_closing = False
        master.bind("<Button-1>", self.reset_inactivity_timer); master.bind("<Key>", self.reset_inactivity_timer); master.bind("<Motion>", self.reset_inactivity_timer)
        self.master.after(self.inactivity_check_interval_ms, self.check_inactivity); self.master.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.main_frame = ttk.Frame(master, padding="10"); self.main_frame.pack(fill=tk.BOTH, expand=True); self.left_frame = ttk.Frame(self.main_frame, width=250); self.left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10)); self.left_frame.pack_propagate(False); self.right_frame = ttk.Frame(self.main_frame); self.right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        self.style = ttk.Style(); self.style.theme_use('clam'); self.style.configure('TButton', font=('Arial', 10), padding=10); self.style.map('TButton', foreground=[('pressed', 'white'), ('active', 'white')], background=[('pressed', '#4f46e5'), ('active', '#3b82f6')]); self.style.configure('Selected.TButton', background='#4f46e5', foreground='white'); self.style.map('Selected.TButton', background=[('pressed', '#4338ca'), ('active', '#4f46e5')], foreground=[('pressed', 'white'), ('active', 'white')])
        self.header_label = ttk.Label(self.main_frame, text="Asana Task Automation", font=('Arial', 16, 'bold')); self.header_label.pack(side=tk.TOP, pady=5, anchor='n')
        self.mode_frame = ttk.LabelFrame(self.left_frame, text="Select Operation Mode"); self.mode_frame.pack(pady=10, fill='x')
        button_details = [("1. Heater Board Swapped", 'heater_swap', 'heaterSwapBtn'), ("2. COR Operation", 'cor', 'corOperationBtn'), ("3. Cal Cert Generator", 'cal_cert', 'calCertBtn'), ("4. Device Complete", 'device_complete', 'deviceCompleteBtn'), ("5. Custom Operation", 'custom', 'customBtn')]
        self.buttons = {}
        for text, op_name, btn_id in button_details:
            btn = ttk.Button(self.mode_frame, text=text, command=lambda op=op_name, id=btn_id: self.select_operation(op, id)); btn.pack(pady=5, fill='x'); btn.id = btn_id; self.buttons[btn_id] = btn
        self.wip_frame = ttk.LabelFrame(self.right_frame, text="Enter WIP Number"); self.wip_frame.pack(pady=10, fill='x')
        self.wip_entry = ttk.Entry(self.wip_frame, font=('Arial', 12)); self.wip_entry.pack(pady=5, fill='x'); self.wip_entry.config(state='disabled'); self.wip_entry.bind("<Return>", self.process_wip_from_enter_key)
        self.result_frame = ttk.LabelFrame(self.right_frame, text="Results"); self.result_frame.pack(pady=10, fill='both', expand=True)
        self.result_text = tk.Text(self.result_frame, wrap='word', font=('Arial', 10), state='disabled', bg='lightgray'); self.result_text.pack(fill='both', expand=True)
        self.asana_client = None; self.config_data = {}; self.operational_gids = {}; self.update_result("Initializing...", is_error=False); self.validate_and_initialize()

    def process_wip(self):
        if self.selected_operation == 'custom' and not self.custom_actions:
            self.update_result("No custom actions defined.", is_error=True); return
        
        wip_number = self.wip_entry.get().strip()
        if not self.selected_operation or not wip_number:
            self.update_result("Please select an operation and enter a WIP.", is_error=True); return
        
        self.update_result(f"Processing '{wip_number}'..."); self.wip_entry.config(state='disabled')
        
        operation_to_run = None
        args = ()

        if self.selected_operation == 'cor':
            dialog = CorReasonDialog(self.master, title="Reason for COR");
            if dialog.result is None: self.update_result("COR operation cancelled.", is_error=True); self.wip_entry.config(state='normal'); return
            reason = dialog.result["reason"]; tag_gid_to_add = self.operational_gids["COR_REASON_TAG_GIDS"].get(reason)
            operation_to_run = process_cor_operation
            args = (wip_number, reason, tag_gid_to_add)

        elif self.selected_operation == 'cal_cert':
            tech_id = simpledialog.askstring("Tech ID", "TECH ID?", parent=self.master)
            if tech_id is None: self.update_result("Operation cancelled.", is_error=True); self.wip_entry.config(state='normal'); return
            operation_to_run = generate_cal_cert
            args = (wip_number, tech_id)

        elif self.selected_operation == 'device_complete':
            cert_path = filedialog.askopenfilename(title=f"Select Certificate for WIP {wip_number}", filetypes=[("Excel Files", "*.xlsx"), ("All files", "*.*")])
            if not cert_path: self.update_result("File selection cancelled.", is_error=True); self.wip_entry.config(state='normal'); return
            operation_to_run = process_device_complete
            args = (wip_number, cert_path)

        elif self.selected_operation == 'custom':
            operation_to_run = process_custom_operation
            args = (wip_number, self.custom_actions)

        elif self.selected_operation == 'heater_swap':
            operation_to_run = process_heater_board_swap
            args = (wip_number,)

        if operation_to_run:
            threading.Thread(target=self._run_operation, args=(operation_to_run,) + args, daemon=True).start()

    def _run_operation(self, operation_func, *args):
        op_data = self.operational_gids.copy()
        op_data.update(APP_CONFIG)
        op_data['full_config'] = self.config_data
        
        result = operation_func(self.asana_client, op_data, *args)
        self.master.after(0, self._update_ui_with_result, result)
        
    def _update_ui_with_result(self, result):
        if result and result["success"]: self._play_sound('success.mp3'); self.update_result(result['message'])
        elif result: self._play_sound('error.mp3'); self.update_result(f"FAILURE:\n{result['message']}", is_error=True)
        else: self._play_sound('error.mp3'); self.update_result("An unexpected error occurred.", is_error=True)
        self.wip_entry.config(state='normal'); self.wip_entry.delete(0, tk.END); self.wip_entry.focus_set()
        
    # --- Other methods are unchanged ---
    def find_gids_by_name(self, item_list, name): return [item.get('gid') for item in item_list if item.get('name', '').lower() == name.lower()]
    def resolve_gids_from_dump(self, config_data):
        all_tags = config_data.get('tags', []); gids_ok, errors = True, []
        all_projects, all_users = config_data.get('projects', []), config_data.get('users', [])
        project_gid_list = self.find_gids_by_name(all_projects, APP_CONFIG["PROJECT_NAME"]); project_gid = project_gid_list[0] if project_gid_list else None
        if project_gid:
            self.operational_gids["PROJECT_GID"] = project_gid; project_obj = next((p for p in all_projects if p['gid'] == project_gid), None); project_sections = project_obj.get('sections', []) if project_obj else []
            for sec_key, sec_name_key in [("NEEDS_COR_SECTION_GID", "NEEDS_COR_SECTION_NAME"), ("PURGE_SECTION_GID", "PURGE_SECTION_NAME"), ("READY_FOR_BUYER_SECTION_GID", "READY_FOR_BUYER_SECTION_NAME")]:
                sec_gid_list = self.find_gids_by_name(project_sections, APP_CONFIG[sec_name_key])
                if sec_gid_list: self.operational_gids[sec_key] = sec_gid_list[0]
                else: gids_ok = False; errors.append(f"Section '{APP_CONFIG[sec_name_key]}' not found.")
        else: gids_ok = False; errors.append(f"Project '{APP_CONFIG['PROJECT_NAME']}' not found.")
        for user_key, user_name_key in [("ACCOUNT_MANAGER_ASSIGNEE_GID", "ACCOUNT_MANAGER_ASSIGNEE_NAME"), ("SHARED_SUBTASK_ASSIGNEE_GID", "SHARED_SUBTASK_ASSIGNEE_NAME")]:
            user_gid_list = self.find_gids_by_name(all_users, APP_CONFIG[user_name_key])
            if user_gid_list: self.operational_gids[user_key] = user_gid_list[0]
            else: gids_ok = False; errors.append(f"User '{APP_CONFIG[user_name_key]}' not found.")
        single_tags = ["COR_TAG_NAME", "DEVICE_COMPLETE_TAG_NAME", "PURGE_SECTION_NAME"]
        for tag_name_key in single_tags:
            tag_gids = self.find_gids_by_name(all_tags, APP_CONFIG[tag_name_key])
            if tag_gids: self.operational_gids[f"{tag_name_key}_GID"] = tag_gids[0]
            else: gids_ok = False; errors.append(f"Tag '{APP_CONFIG[tag_name_key]}' not found.")
        heater_gids = self.find_gids_by_name(all_tags, APP_CONFIG["HEATER_SWAP_TAG_NAME"]);
        if heater_gids: self.operational_gids["HEATER_SWAP_TAG_GIDS"] = heater_gids
        else: gids_ok = False; errors.append(f"Tag '{APP_CONFIG['HEATER_SWAP_TAG_NAME']}' not found.")
        self.operational_gids["COR_REASON_TAG_GIDS"] = {}
        for reason, tag_name in APP_CONFIG["COR_REASON_TAGS"].items():
            reason_gids = self.find_gids_by_name(all_tags, tag_name)
            if reason_gids: self.operational_gids["COR_REASON_TAG_GIDS"][reason] = reason_gids[0]
            else: logging.warning(f"COR Reason Tag '{tag_name}' not found during startup. It may not be available for use.")
        return {"success": gids_ok, "message": "\n".join(errors)}
    def validate_and_initialize(self):
        if not ASANA_TOKEN: messagebox.showerror("Configuration Error", "CRITICAL ERROR: ASANA_PAT not set."); self.master.after(100, self.master.destroy); return
        try:
            base_path = _get_resource_path('.')
            with open(os.path.join(base_path, 'config.json'), 'r') as f: self.config_data = json.load(f)
        except Exception as e: messagebox.showerror("Configuration Error", f"Could not load config.json.\nError: {e}"); self.master.after(100, self.master.destroy); return
        self.asana_client = AsanaClient(ASANA_TOKEN, self.config_data.get("workspace_id"))
        gid_result = self.resolve_gids_from_dump(self.config_data)
        if not gid_result["success"]: messagebox.showerror("Configuration Error", f"Could not find GIDs:\n\n{gid_result['message']}"); self.master.after(100, self.master.destroy); return
        logging.info("Application initialized successfully."); self.update_result("Please select an operation mode above.", is_error=False)
    def on_closing(self):
        if not self.is_closing: self.is_closing = True; logging.info("Application closing."); self.master.destroy()
    def reset_inactivity_timer(self, event=None): self.last_activity_time = time.time()
    def check_inactivity(self):
        if self.is_closing: return
        if (time.time() - self.last_activity_time) * 1000 > self.inactivity_timeout_ms:
            if not self.is_closing: self.is_closing = True; logging.info("Closing due to inactivity."); self.master.destroy()
        else: self.master.after(self.inactivity_check_interval_ms, self.check_inactivity)
    def select_operation(self, operation_name, button_id):
        self.selected_operation = operation_name; self.update_button_styles(button_id)
        if operation_name == 'custom':
            dialog = CustomActionDialog(self.master); self.custom_actions = dialog.result if dialog.result is not None else []
            if not self.custom_actions: self.update_result("Custom operation cancelled.", is_error=True); self.wip_entry.config(state='disabled'); return
            else: self.update_result(f"Custom mode active. Enter WIP.", is_error=False)
        else:
            self.custom_actions = []; mode_text = { 'heater_swap': "Heater Board Swapped", 'cor': "COR Operation", 'cal_cert': "Cal Cert Generator", 'device_complete': "Device Complete" }.get(operation_name, "Unknown")
            self.update_result(f"Mode selected: {mode_text}. Enter WIP.", is_error=False)
        self.wip_entry.config(state='normal'); self.wip_entry.delete(0, tk.END); self.wip_entry.focus_set()
    def update_button_styles(self, active_button_id):
        for btn_id, btn in self.buttons.items(): btn.config(style='Selected.TButton' if btn.id == active_button_id else 'TButton')
    def update_result(self, message, is_error=False):
        self.result_text.config(state='normal'); self.result_text.delete('1.0', tk.END); self.result_text.insert(tk.END, message + "\n\n"); self.result_text.see(tk.END); self.result_text.config(state='disabled')
    def process_wip_from_enter_key(self, event=None): self.process_wip()
    def _play_sound(self, sound_file):
        def target():
            try: playsound(_get_resource_path(sound_file))
            except Exception as e: logging.warning(f"Could not play sound {sound_file}: {e}")
        threading.Thread(target=target, daemon=True).start()

if __name__ == "__main__":
    if not ASANA_TOKEN: print("CRITICAL: ASANA_PAT not set."); sys.exit(1)
    root = tk.Tk(); app = AsanaAutomationApp(root); root.mainloop()