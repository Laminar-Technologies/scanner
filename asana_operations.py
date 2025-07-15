# asana_operations.py (v1.5.1)
import logging
import os
import sys
from datetime import datetime
import win32com.client
from win32com.client import gencache
import pythoncom

def _get_resource_path(relative_path):
    try: base_path = sys._MEIPASS
    except Exception: base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

def _find_and_validate_tasks(asana_client, op_data, wip_number):
    # This function is correct and unchanged
    opt_fields = "name,gid,parent,memberships,tags.name,tags.gid"
    task_result = asana_client.find_task_by_wip(wip_number, opt_fields=opt_fields)
    if not task_result.get("success"): return task_result
    task_data = task_result["task_data"]
    subtask_data = None
    if task_data.get('parent'):
        subtask_details_result = asana_client.get_task_details(task_data['gid'], opt_fields=opt_fields);
        if not subtask_details_result.get("success"): return subtask_details_result
        subtask_data = subtask_details_result.get("data", {}).get("data", {}); parent_gid = task_data['parent']['gid']
    else:
        parent_gid = task_data['gid']
        subtasks_result = asana_client.get_subtasks_for_task(parent_gid);
        if not subtasks_result.get("success"): return subtasks_result
        subtasks = subtasks_result.get("data", {}).get("data", [])
        matching_subtask = next((st for st in subtasks if wip_number.lower() in st.get('name', '').lower()), None)
        if not matching_subtask: return {"success": False, "message": f"No subtask for '{wip_number}' found."}
        subtask_details_result = asana_client.get_task_details(matching_subtask['gid'], opt_fields=opt_fields)
        if not subtask_details_result.get("success"): return subtask_details_result
        subtask_data = subtask_details_result.get("data", {}).get("data", {})
    parent_details = asana_client.get_task_details(parent_gid, opt_fields="name,tags.gid")
    parent_data = parent_details.get("data", {}).get("data", {}); parent_name = parent_data.get("name", "Unknown Parent")
    parent_tags = {tag['gid'] for tag in parent_data.get("tags", [])}
    purge_gid = op_data.get("PURGE_SECTION_NAME_GID")
    if purge_gid and purge_gid in parent_tags:
        return {"success": False, "message": f"ERROR: Parent task '{parent_name}' has the PURGE tag."}
    logging.info(f"Found Tasks: Parent='{parent_name}', Subtask='{subtask_data.get('name')}'")
    return {"success": True, "parent_gid": parent_gid, "parent_name": parent_name, "subtask_data": subtask_data}

def _resolve_name_or_gid(value, raw_config):
    # This helper is correct and unchanged
    if value.isdigit(): return value, None
    search_lists = { 'User': raw_config.get('users', []), 'Tag': raw_config.get('tags', []) }
    for item_type, item_list in search_lists.items():
        for item in item_list:
            if item.get('name', '').lower() == value.lower(): return item['gid'], None
    for project in raw_config.get('projects', []):
        if project.get('name', '').lower() == value.lower(): return project['gid'], None
        for section in project.get('sections', []):
            if section.get('name', '').lower() == value.lower(): return section['gid'], None
    return None, f"Could not find GID for name '{value}'."

def generate_cal_cert(asana_client, op_data, wip_number, tech_id):
    # This function is correct and unchanged
    pythoncom.CoInitialize()
    task_validation_result = _find_and_validate_tasks(asana_client, op_data, wip_number)
    if not task_validation_result["success"]: pythoncom.CoUninitialize(); return task_validation_result
    subtask_data = task_validation_result["subtask_data"]; subtask_name = subtask_data.get('name', '')
    title_parts = subtask_name.split(); errors = []
    def safe_get(lst, index, default=''):
        try: return lst[index]
        except IndexError: errors.append(f"Could not find item at position {index + 1}."); return default
    model_number = safe_get(title_parts, 1)
    part_numbers = op_data.get('full_config', {}).get('part_numbers', {})
    part_number = part_numbers.get(model_number, "NOT FOUND")
    if part_number == "NOT FOUND": errors.append(f"Part number for model '{model_number}' not in config.")
    heater_tag_gids = set(op_data.get("HEATER_SWAP_TAG_GIDS", [])); subtask_tag_gids = {tag['gid'] for tag in subtask_data.get('tags', [])}
    service_type = "Level 1 Repair" if heater_tag_gids.intersection(subtask_tag_gids) else "Clean & Calibrate"
    data_to_populate = { 'G5': part_number, 'G6': wip_number, 'G7': datetime.now().strftime("%Y-%m-%d"), 'G8': model_number, 'G9': safe_get(title_parts, 8), 'G10': f" {safe_get(title_parts, 2)} TORR", 'G11': safe_get(title_parts, 4), 'G12': safe_get(title_parts, 5), 'G13': 'PVS100', 'G14': safe_get(title_parts, 6), 'G15': tech_id, 'G16': service_type }
    excel = None
    try:
        template_path = _get_resource_path('template.xltx'); excel = gencache.EnsureDispatch('Excel.Application'); excel.Visible = True; wb = excel.Workbooks.Open(template_path); ws = wb.Worksheets(1)
        for cell, value in data_to_populate.items(): ws.Range(cell).Value = value
        wb.Activate()
        success_message = f"Certificate for WIP {wip_number} is open and unsaved."
        if errors: success_message += f"\n\nWARNINGS:\n" + "\n".join(errors)
        return {"success": True, "message": success_message}
    except Exception as e:
        if excel: excel.Quit()
        logging.error(f"Excel error: {e}", exc_info=True); return {"success": False, "message": f"An error occurred with Excel: {e}"}
    finally:
        ws = None; wb = None; excel = None; pythoncom.CoUninitialize()

def process_heater_board_swap(asana_client, op_data, wip_number):
    # This function is correct and unchanged
    task_validation_result = _find_and_validate_tasks(asana_client, op_data, wip_number)
    if not task_validation_result["success"]: return task_validation_result
    subtask_gid = task_validation_result["subtask_data"].get('gid')
    heater_tag_to_add = op_data.get("HEATER_SWAP_TAG_GIDS", [None])[0]
    if not heater_tag_to_add: return {"success": False, "message": "No 'Heater Board Replacement' tag GID found in config."}
    add_tag_result = asana_client.add_tag_to_task(subtask_gid, heater_tag_to_add)
    if add_tag_result["success"]: return {"success": True, "message": f"Added 'Heater Board Swapped' tag to WIP {wip_number}."}
    else: return {"success": False, "message": f"Failed to add tag.\n{add_tag_result.get('message')}"}

def process_cor_operation(asana_client, op_data, wip_number, reason, tag_gid_to_add=None):
    # This function is correct and unchanged
    task_validation_result = _find_and_validate_tasks(asana_client, op_data, wip_number)
    if not task_validation_result["success"]: return task_validation_result
    parent_gid, subtask_gid = task_validation_result["parent_gid"], task_validation_result["subtask_data"]['gid']
    all_ops_success, messages = True, []
    def log_op(msg, res): nonlocal all_ops_success; status = 'Success' if res["success"] else f"FAILED: {res.get('message', 'Unknown')}"; messages.append(f"• {msg}: {status}");
    comment = f"AUTO: {reason}"; log_op("Adding comment", asana_client.add_comment_to_task(subtask_gid, comment))
    if tag_gid_to_add: log_op(f"Adding tag '{reason}'", asana_client.add_tag_to_task(subtask_gid, tag_gid_to_add))
    log_op("Tagging 'Return Unrepaired'", asana_client.add_tag_to_task(subtask_gid, op_data['COR_TAG_NAME_GID']))
    parent_name = task_validation_result["parent_name"]
    if not parent_name.strip().upper().startswith("*COR*"): log_op("Renaming parent", asana_client.change_task_name(parent_gid, f"*COR* {parent_name}"))
    subtask_name = task_validation_result["subtask_data"]['name']
    if not subtask_name.strip().upper().startswith("*COR*"): log_op("Renaming subtask", asana_client.change_task_name(subtask_gid, f"*COR* {subtask_name}"))
    log_op("Assigning parent", asana_client.assign_task_to_user(parent_gid, op_data['ACCOUNT_MANAGER_ASSIGNEE_GID']))
    log_op("Assigning subtask", asana_client.assign_task_to_user(subtask_gid, op_data['SHARED_SUBTASK_ASSIGNEE_GID']))
    log_op("Moving parent to 'Needs COR'", asana_client.move_task_to_section(parent_gid, op_data['NEEDS_COR_SECTION_GID']))
    summary = f"WIP {wip_number} marked as COR."; final_message = f"{summary}\n\n--- Details ---\n" + "\n".join(messages)
    return {"success": all_ops_success, "message": final_message}

def process_device_complete(asana_client, op_data, wip_number, cert_path):
    # This function is correct and unchanged
    task_validation_result = _find_and_validate_tasks(asana_client, op_data, wip_number)
    if not task_validation_result["success"]: return task_validation_result
    parent_gid, subtask_gid = task_validation_result["parent_gid"], task_validation_result["subtask_data"]['gid']
    all_ops_success, messages = True, []
    def log_op(msg, res): nonlocal all_ops_success; status = 'Success' if res["success"] else f"FAILED: {res.get('message', 'Unknown')}"; messages.append(f"• {msg}: {status}");
    log_op("Adding comment", asana_client.add_comment_to_task(subtask_gid, "AUTO: Device Complete"))
    log_op("Uploading certificate", asana_client.upload_attachment(subtask_gid, cert_path))
    log_op("Assigning parent", asana_client.assign_task_to_user(parent_gid, op_data['ACCOUNT_MANAGER_ASSIGNEE_GID']))
    log_op("Assigning subtask", asana_client.assign_task_to_user(subtask_gid, op_data['SHARED_SUBTASK_ASSIGNEE_GID']))
    log_op("Moving parent", asana_client.move_task_to_section(parent_gid, op_data['READY_FOR_BUYER_SECTION_GID']))
    log_op("Tagging subtask", asana_client.add_tag_to_task(subtask_gid, op_data['DEVICE_COMPLETE_TAG_NAME_GID']))
    summary = f"WIP {wip_number} marked as Device Complete."; final_message = f"{summary}\n\n--- Details ---\n" + "\n".join(messages)
    return {"success": all_ops_success, "message": final_message}

def process_custom_operation(asana_client, op_data, wip_number, actions):
    # CORRECTED: Added pythoncom initialization for thread safety
    pythoncom.CoInitialize()
    try:
        task_validation_result = _find_and_validate_tasks(asana_client, op_data, wip_number)
        if not task_validation_result["success"]: return task_validation_result
        
        subtask_gid = task_validation_result["subtask_data"]['gid']
        raw_config = op_data.get('full_config', {})
        all_ops_success, messages = True, []
        
        def log_op(msg, res):
            nonlocal all_ops_success
            if not res["success"]: all_ops_success = False
            status = 'Success' if res["success"] else f"FAILED: {res.get('message', 'Unknown')}"
            messages.append(f"• {msg}: {status}")

        for action in actions:
            action_type, value = action['type'], action['value']
            gid_to_use, error_msg = value, None
            if action_type != 'add_comment':
                gid_to_use, error_msg = _resolve_name_or_gid(value, raw_config)
            if error_msg:
                log_op(f"Action '{action_type}' with value '{value}'", {"success": False, "message": error_msg}); continue
            if action_type == 'add_tag': log_op(f"Adding tag '{value}'", asana_client.add_tag_to_task(subtask_gid, gid_to_use))
            elif action_type == 'assign_to': log_op(f"Assigning to '{value}'", asana_client.assign_task_to_user(subtask_gid, gid_to_use))
            elif action_type == 'move_to': log_op(f"Moving to section '{value}'", asana_client.move_task_to_section(subtask_gid, gid_to_use))
            elif action_type == 'add_comment': log_op("Adding comment", asana_client.add_comment_to_task(subtask_gid, f"AUTO: {value}"))
        
        summary = f"Custom operation completed for WIP {wip_number}."
        final_message = f"{summary}\n\n--- Details ---\n" + "\n".join(messages)
        return {"success": all_ops_success, "message": final_message}
    finally:
        pythoncom.CoUninitialize()