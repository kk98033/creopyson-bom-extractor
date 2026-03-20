import creopyson
import os
import time
import shutil
import glob
import json

# ==========================================
# 配置區域
# ==========================================
MODELS_DIR_NAME = "models"
THREED_DIR_NAME = "3D圖檔"
TEMP_NEU_DIR = "temp_neu_M24248_parent"
OUTPUT_LOG_FILE = "M24248_parent_debug_log.txt"
OUTPUT_JSON_FILE = "M24248_parent_debug_results.json"
TARGET_PART_NO = "M24248" 

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
    """清理 Creo 垃圾檔案"""
    junk_patterns = ["*.log*", "*.xml*", "*.inf*", "*.err*", "*.out*", "trail.txt*", "creo_command_*.txt"]
    for pattern in junk_patterns:
        for f in glob.glob(os.path.join(directory, pattern)):
            try:
                if not any(f.endswith(ext) for ext in [".py", ".stp", ".json", ".asm", ".prt"]):
                    os.remove(f)
            except: pass

def normalize_name(name):
    """標準化名稱：移除 :1 序號、去除空白並轉大寫"""
    if not name: return ""
    return name.split(":")[0].strip().upper()

def process_bom_universal(data, parent_map, all_parts, current_parent=None):
    """強健的 BOM 爬蟲：建立 parent_map"""
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

# ==========================================
# 樹狀結構處理 Function
# ==========================================

def build_nested_tree(dim_data, root_name):
    """將扁平的 parent 關係轉換為嵌套字典結構"""
    tree_map = {}
    # 先建立 父對子 的一對多關係表
    for comp, info in dim_data.items():
        p = normalize_name(info["parent"])
        if p not in tree_map:
            tree_map[p] = []
        tree_map[p].append(comp)

    def _recurse_build(current_node_name):
        node = {}
        # 找出當前節點的所有子節點
        children = tree_map.get(normalize_name(current_node_name), [])
        for child in children:
            node[child] = _recurse_build(child)
        return node

    # 從 Root 開始建立
    return {root_name: _recurse_build(root_name)}

def generate_tree_visual_lines(tree_dict, prefix=""):
    """
    遞迴遍歷『嵌套樹狀字典』，生成帶有 ├── └── 的字串列表
    """
    lines = []
    items = list(tree_dict.items())
    
    for i, (name, children) in enumerate(items):
        is_last = (i == len(items) - 1)
        
        # 如果是頂層 Root
        if prefix == "":
            lines.append(f" {name}")
            lines.extend(generate_tree_visual_lines(children, prefix="    "))
        else:
            # 根據是否為最後一個子節點決定符號
            connector = "└── " if is_last else "├── "
            # prefix 的最後四個字元要換成連接符號
            lines.append(f"{prefix[:-4]}{connector}{name}")
            
            # 遞迴下一層，prefix 增加縮進
            new_prefix = prefix + ("    " if is_last else "│   ")
            lines.extend(generate_tree_visual_lines(children, new_prefix))
            
    return lines

# ==========================================
# 主程式
# ==========================================

def main():
    c = creopyson.Client()
    try:
        c.connect()
        print(f"[INFO] Connected to Creoson. Target: {TARGET_PART_NO}")
    except: 
        print("[ERROR] Could not connect to Creoson.")
        return

    execution_dir = os.path.dirname(os.path.abspath(__file__)).replace("\\", "/")
    models_root_dir = os.path.join(execution_dir, MODELS_DIR_NAME)
    temp_dir = os.path.join(execution_dir, TEMP_NEU_DIR)
    log_path = os.path.join(execution_dir, OUTPUT_LOG_FILE)
    json_path = os.path.join(execution_dir, OUTPUT_JSON_FILE)

    if os.path.exists(temp_dir): shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)

    target_folder_path = os.path.join(models_root_dir, TARGET_PART_NO)
    if not os.path.exists(target_folder_path):
        print(f"[ERROR] Target folder '{TARGET_PART_NO}' not found.")
        c.disconnect()
        return

    final_json_data = {TARGET_PART_NO: {}}

    with open(log_path, "w", encoding="utf-8") as log:
        log.write(f"=== Assembly Tree Automation: {TARGET_PART_NO} ===\n\n")

        threed_dir_path = os.path.join(target_folder_path, THREED_DIR_NAME)
        creo_threed_path = threed_dir_path.replace("\\", "/")
        
        target_files = sorted(list(set([os.path.basename(f) for ext in INPUT_EXTENSIONS for f in glob.glob(os.path.join(threed_dir_path, ext))])))

        for main_model in target_files:
            print(f"\n[PROCESS] Top File: {main_model}")
            
            # 初始化結構
            file_json_node = {
                "assembly_hierarchy_paths": {}, # 這會存一個樹狀字典
                "components_dimension_data": {}
            }
            parent_map = {} 
            all_parts = []

            try:
                c.file_close_window()
                c.file_erase_not_displayed()
                c.creo_cd(creo_threed_path)

                if main_model.lower().endswith(('.stp', '.step')):
                    active_model = c.interface_import_file(filename=main_model, dirname=creo_threed_path, new_model_type="asm")
                else:
                    active_model = c.file_open(file_=main_model, dirname=creo_threed_path, display=True, activate=True)

                # 解析 BOM
                bom_raw = creopyson.bom.get_paths(c, file_=active_model)
                process_bom_universal(bom_raw, parent_map, all_parts)
                
                unique_sub_parts = sorted(set([p for p in all_parts if normalize_name(p) != normalize_name(active_model)]))
                parts_to_process = [active_model] + unique_sub_parts

                print(f"[INFO] Processing dimensions...")

                for part in parts_to_process:
                    try:
                        c.file_open(file_=part, display=True, activate=True)
                        for f in os.listdir(temp_dir): 
                            try: os.remove(os.path.join(temp_dir, f))
                            except: pass
                        
                        c.interface_export_file(file_type="NEUTRAL", filename=part, dirname=temp_dir.replace("\\", "/"))
                        time.sleep(0.4)

                        base_name = part.split(".")[0].upper()
                        target_neu = next((os.path.join(temp_dir, f) for f in os.listdir(temp_dir) if f.upper().startswith(base_name) and ".NEU" in f.upper()), None)

                        if target_neu:
                            dims = parse_neu_v11(target_neu)
                            if dims:
                                dx, dy, dz = dims
                                current_key = normalize_name(part)
                                parent_name = parent_map.get(current_key, "None")
                                
                                file_json_node["components_dimension_data"][part] = {
                                    "name": part, "dX": round(dx, 3), "dY": round(dy, 3), "dZ": round(dz, 3), "parent": parent_name
                                }
                        c.file_open(file_=active_model, display=True, activate=True)
                    except: pass

                # 生成嵌套樹狀結構
                print(f"[TREE] Building nested structure...")
                nested_tree = build_nested_tree(file_json_node["components_dimension_data"], active_model)
                
                # 存入 JSON
                file_json_node["assembly_hierarchy_paths"] = nested_tree
                
                # 生成可視化字串
                tree_lines = generate_tree_visual_lines(nested_tree)
                tree_visual_str = "\n".join(tree_lines)
                
                # 輸出到 CMD
                print("\n" + "V" * 20 + " ASSEMBLY MODEL TREE " + "V" * 20)
                print(tree_visual_str)
                print("A" * 60 + "\n")
                
                # 輸出到 Log
                log.write(f"\n--- VISUAL MODEL TREE ---\n")
                log.write(tree_visual_str + "\n")
                log.write(f"--------------------------\n")

                final_json_data[TARGET_PART_NO][main_model] = file_json_node

            except Exception as e:
                print(f"[ERROR] {e}")

    c.disconnect()
    with open(json_path, 'w', encoding='utf-8') as jf:
        json.dump(final_json_data, jf, indent=4, ensure_ascii=False)
    
    if os.path.exists(temp_dir): shutil.rmtree(temp_dir)
    cleanup_extra_files(execution_dir)
    print(f"\n[FINISH] Results saved. Model tree printed above.")

if __name__ == "__main__":
    main()