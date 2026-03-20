import creopyson
import os
import time
import shutil
import glob
import json
import argparse

# ==========================================
# 配置區域
# ==========================================
MODELS_DIR_NAME = "models"
THREED_DIR_NAME = "3D圖檔"
TEMP_NEU_DIR = "temp_neu_final"
OUTPUT_LOG_FILE = "process_log.txt"
OUTPUT_JSON_FILE = "process_results.json"
# 支援的副檔名
INPUT_EXTENSIONS = ('*.stp', '*.STP', '*.prt', '*.PRT', '*.asm', '*.ASM')

# ==========================================
# 核心功能函數
# ==========================================
def parse_neu_v11(file_path):
    """解析 NEU 檔案獲取尺寸"""
    min_pt, max_pt = None, None
    try:
        with open(file_path, 'r', encoding='ascii', errors='ignore') as f:
            for line in f:
                if "outline[0]" in line:
                    content = line.strip().split("outline[0]")[-1].strip().replace('[','').replace(']','')
                    min_pt = [float(x) for x in content.split(",")]
                elif "outline[1]" in line:
                    content = line.strip().split("outline[1]")[-1].strip().replace('[','').replace(']','')
                    max_pt = [float(x) for x in content.split(",")]
                if min_pt is not None and max_pt is not None:
                    return abs(max_pt[0]-min_pt[0]), abs(max_pt[1]-min_pt[1]), abs(max_pt[2]-min_pt[2])
        return None
    except: return None

def cleanup_extra_files(directory):
    """清理 Creo 產生的臨時垃圾檔案"""
    junk_patterns = ["*.log*", "*.xml*", "*.inf*", "*.err*", "*.out*", "trail.txt*", "creo_command_*.txt"]
    for pattern in junk_patterns:
        for f in glob.glob(os.path.join(directory, pattern)):
            try:
                if not any(f.endswith(ext) for ext in [".py", ".stp", ".json"]):
                    os.remove(f)
            except: pass

def normalize_name(name):
    """標準化名稱：移除 :1 序號、轉大寫"""
    if not name: return ""
    return name.split(":")[0].strip().upper()

def process_bom_universal(data, parent_map, all_parts, current_parent=None):
    """強健的 BOM 爬蟲：建立 parent_map 與收集所有零件"""
    if isinstance(data, dict):
        this_file = data.get("file")
        if this_file:
            all_parts.append(this_file)
            if current_parent and normalize_name(this_file) != normalize_name(current_parent):
                parent_map[normalize_name(this_file)] = normalize_name(current_parent)
            new_parent = this_file
        else:
            new_parent = current_parent

        for k, v in data.items():
            if k == "file": continue
            process_bom_universal(v, parent_map, all_parts, new_parent)
    elif isinstance(data, list):
        for item in data:
            process_bom_universal(item, parent_map, all_parts, current_parent)

def build_nested_tree(dim_data, root_name):
    """將扁平的 parent 關係轉換為嵌套字典結構"""
    tree_map = {}
    for comp, info in dim_data.items():
        p = normalize_name(info.get("parent", "None"))
        if p not in tree_map:
            tree_map[p] = []
        tree_map[p].append(comp)

    def _recurse_build(current_node_name):
        node = {}
        children = tree_map.get(normalize_name(current_node_name), [])
        for child in children:
            node[child] = _recurse_build(child)
        return node

    return {root_name: _recurse_build(root_name)}

def generate_tree_visual_lines(tree_dict, prefix=""):
    """生成帶有 ├── └── 的模型樹字串列表"""
    lines = []
    items = list(tree_dict.items())
    for i, (name, children) in enumerate(items):
        is_last = (i == len(items) - 1)
        if prefix == "":
            lines.append(f" {name}")
            lines.extend(generate_tree_visual_lines(children, prefix="    "))
        else:
            connector = "└── " if is_last else "├── "
            lines.append(f"{prefix[:-4]}{connector}{name}")
            new_prefix = prefix + ("    " if is_last else "│   ")
            lines.extend(generate_tree_visual_lines(children, new_prefix))
    return lines

# ==========================================
# 主程式
# ==========================================
def main():
    parser = argparse.ArgumentParser(description="Creo CAD 尺寸與模型樹自動化提取工具")
    parser.add_argument('--debug', action='store_true', help='啟用 Debug 模式，僅處理前 3 筆模型資料夾')
    args = parser.parse_args()

    c = creopyson.Client()
    try:
        c.connect()
        print("[INFO] Connected to Creoson.")
    except: 
        print("[ERROR] Could not connect to Creoson. Please ensure Setup Exe is running.")
        return

    execution_dir = os.path.dirname(os.path.abspath(__file__)).replace("\\", "/")
    models_root_dir = os.path.join(execution_dir, MODELS_DIR_NAME)
    temp_dir = os.path.join(execution_dir, TEMP_NEU_DIR)
    log_path = os.path.join(execution_dir, OUTPUT_LOG_FILE)
    json_path = os.path.join(execution_dir, OUTPUT_JSON_FILE)

    cleanup_extra_files(execution_dir)
    if os.path.exists(temp_dir): shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)

    if not os.path.exists(models_root_dir):
        print(f"[ERROR] '{MODELS_DIR_NAME}' directory not found.")
        c.disconnect()
        return

    final_json_data = {}
    part_folders = sorted([f for f in os.listdir(models_root_dir) if os.path.isdir(os.path.join(models_root_dir, f))])

    if args.debug:
        print(f"[DEBUG] Debug mode enabled. Processing only first 3 folders.")
        part_folders = part_folders[:3]

    with open(log_path, "w", encoding="utf-8") as log:
        log.write(f"=== Creo Automation Process Log ({time.ctime()}) ===\n")
        if args.debug: log.write("[MODE] DEBUG (First 3 folders only)\n")
        log.write("\n")

        for part_no in part_folders:
            current_part_path = os.path.join(models_root_dir, part_no)
            threed_dir_path = os.path.join(current_part_path, THREED_DIR_NAME)

            if not os.path.exists(threed_dir_path):
                print(f"[WARN] Skipped '{part_no}': Folder not found.")
                continue

            print("\n" + "#"*80)
            print(f"[PROJECT] Processing Parent Model Folder: {part_no}")
            print("#"*80)
            log.write(f"\n{'='*20} PROJECT: {part_no} {'='*20}\n")

            final_json_data[part_no] = {}
            target_files = sorted(list(set([os.path.basename(f) for ext in INPUT_EXTENSIONS for f in glob.glob(os.path.join(threed_dir_path, ext))])))

            for main_model in target_files:
                if OUTPUT_LOG_FILE.split('.')[0] in main_model.lower(): continue

                print(f"\n[PROCESS] Current Top File: {main_model}")
                log.write(f"\n>> TOP FILE: {main_model}\n")
                
                file_json_node = {"assembly_hierarchy_paths": {}, "components_dimension_data": {}}
                parent_map = {} 
                all_parts = []

                try:
                    c.file_close_window()
                    c.file_erase_not_displayed()
                    creo_threed_path = threed_dir_path.replace("\\", "/")
                    c.creo_cd(creo_threed_path)

                    # 開啟主模型
                    if main_model.lower().endswith(('.stp', '.step')):
                        active_model = c.interface_import_file(filename=main_model, dirname=creo_threed_path, new_model_type="asm")
                    else:
                        active_model = c.file_open(file_=main_model, dirname=creo_threed_path, display=True, activate=True)

                    # 建立 BOM 地圖 (在開啟子零件前執行)
                    bom_raw = creopyson.bom.get_paths(c, file_=active_model)
                    process_bom_universal(bom_raw, parent_map, all_parts)
                    
                    unique_sub_parts = sorted(set([p for p in all_parts if normalize_name(p) != normalize_name(active_model)]))
                    parts_to_process = [active_model] + unique_sub_parts

                    # 表頭輸出
                    header = f"{'Component Part Name':<45} | {'dX':<10} | {'dY':<10} | {'dZ':<10}"
                    print("-" * 85); print(header); log.write(header + "\n")

                    for part in parts_to_process:
                        try:
                            c.file_open(file_=part, display=True, activate=True)
                            for f in os.listdir(temp_dir): 
                                try: os.remove(os.path.join(temp_dir, f))
                                except: pass

                            c.interface_export_file(file_type="NEUTRAL", filename=part, dirname=temp_dir.replace("\\", "/"))
                            time.sleep(0.5)

                            base_name = part.split(".")[0].upper()
                            target_neu = next((os.path.join(temp_dir, f) for f in os.listdir(temp_dir) if f.upper().startswith(base_name) and ".NEU" in f.upper()), None)

                            if target_neu:
                                dims = parse_neu_v11(target_neu)
                                if dims:
                                    dx, dy, dz = dims
                                    parent_name = parent_map.get(normalize_name(part), "None")
                                    
                                    # 表格行輸出
                                    result_line = f"{part:<45} | {dx:<10.2f} | {dy:<10.2f} | {dz:<10.2f}"
                                    print(result_line)
                                    log.write(result_line + "\n")

                                    file_json_node["components_dimension_data"][part] = {
                                        "name": part, "dX": round(dx, 3), "dY": round(dy, 3), "dZ": round(dz, 3), "parent": parent_name
                                    }
                            c.file_open(file_=active_model, display=True, activate=True)
                        except: pass

                    # --- 生成樹狀結構與視覺化輸出 ---
                    nested_tree = build_nested_tree(file_json_node["components_dimension_data"], active_model)
                    file_json_node["assembly_hierarchy_paths"] = nested_tree
                    
                    tree_visual = "\n".join(generate_tree_visual_lines(nested_tree))
                    print("\n[MODEL TREE]")
                    print(tree_visual)
                    log.write("\n[MODEL TREE]\n" + tree_visual + "\n")

                    final_json_data[part_no][main_model] = file_json_node

                except Exception as e:
                    print(f"[ERROR] {main_model}: {e}")
                
                try:
                    c.file_close_window()
                    c.file_erase_not_displayed()
                except: pass

            cleanup_extra_files(execution_dir)

    # 存檔
    c.disconnect()
    with open(json_path, 'w', encoding='utf-8') as jf:
        json.dump(final_json_data, jf, indent=4, ensure_ascii=False)
    
    if os.path.exists(temp_dir): shutil.rmtree(temp_dir)
    print(f"\n[FINISH] Process complete. Results in '{OUTPUT_JSON_FILE}'.")

if __name__ == "__main__":
    main()