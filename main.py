import creopyson
import os
import time
import shutil
import glob

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
                if not f.endswith(".py") and not f.endswith(".stp"):
                    os.remove(f)
            except: pass

def get_all_components(data, res):
    if isinstance(data, dict):
        if "file" in data: res.append(data["file"])
        for v in data.values(): get_all_components(v, res)
    elif isinstance(data, list):
        for i in data: get_all_components(i, res)

def main():
    c = creopyson.Client()
    try:
        c.connect()
        print("[INFO] Connected to Creoson.")
    except: 
        print("[ERROR] Could not connect to Creoson.")
        return

    current_dir = os.path.dirname(os.path.abspath(__file__)).replace("\\", "/")
    temp_dir = os.path.join(current_dir, "temp_neu_final")
    log_path = os.path.join(current_dir, "process_log.txt")

    cleanup_extra_files(current_dir)

    extensions = ('*.stp', '*.STP', '*.prt', '*.PRT', '*.asm', '*.ASM')
    target_files = []
    for ext in extensions:
        target_files.extend([os.path.basename(f) for f in glob.glob(os.path.join(current_dir, ext))])
    target_files = sorted(list(set(target_files)))
    
    with open(log_path, "w", encoding="utf-8") as log:
        log.write(f"=== Creo Automation Process Log ({time.ctime()}) ===\n\n")

        for main_model in target_files:
            # 排除生成的 log 和 temp
            if "process_log" in main_model.lower(): continue

            print("\n" + "="*80)
            print(f"[PROCESS] Current Top Model: {main_model}")
            log.write(f"\n>> TOP MODEL: {main_model}\n")
            
            try:
                c.file_close_window()
                c.file_erase_not_displayed()
                if os.path.exists(temp_dir): shutil.rmtree(temp_dir)
                os.makedirs(temp_dir)

                if main_model.lower().endswith('.stp'):
                    active_model = c.interface_import_file(filename=main_model, dirname=current_dir, new_model_type="asm")
                else:
                    active_model = c.file_open(file_=main_model, dirname=current_dir, display=True, activate=True)

                bom_raw = creopyson.bom.get_paths(c, file_=active_model)
                all_parts = []
                get_all_components(bom_raw, all_parts)
                unique_sub_parts = sorted(set([p for p in all_parts if p.upper() != active_model.upper()]))
                
                # 記錄結構
                relation_str = f"STRUCTURE: {active_model} = " + ", ".join(unique_sub_parts)
                print(f"[BOM] {relation_str}")
                log.write(f"{relation_str}\n" + "-" * 80 + "\n")

                # 表頭
                header = f"{'Component Part Name':<45} | {'dX':<10} | {'dY':<10} | {'dZ':<10}"
                print("-" * 80)
                print(header)
                log.write(header + "\n")

                parts_to_process = unique_sub_parts if unique_sub_parts else [active_model]

                for part in parts_to_process:
                    try:
                        c.file_open(file_=part, display=True, activate=True)
                        for f in os.listdir(temp_dir): os.remove(os.path.join(temp_dir, f))

                        c.interface_export_file(file_type="NEUTRAL", filename=part, dirname=temp_dir.replace("\\", "/"))
                        time.sleep(0.6)

                        base_name = part.split(".")[0].upper()
                        target_neu = None
                        for f in os.listdir(temp_dir):
                            if f.upper().startswith(base_name) and ".NEU" in f.upper():
                                target_neu = os.path.join(temp_dir, f)
                                break

                        if target_neu:
                            dims = parse_neu_v11(target_neu)
                            if dims:
                                dx, dy, dz = dims
                                result_line = f"{part:<45} | {dx:<10.2f} | {dy:<10.2f} | {dz:<10.2f}"
                                print(result_line) # <--- 這裡加回來了
                                log.write(result_line + "\n")
                        
                        c.file_open(file_=active_model, display=True, activate=True)
                    except: pass

                log.write("=" * 80 + "\n")
            except Exception as e:
                print(f"[ERROR] Failed to process {main_model}: {e}")
                log.write(f"FAILED: {main_model} - {e}\n")

    # 最後大掃除
    if os.path.exists(temp_dir): shutil.rmtree(temp_dir)
    cleanup_extra_files(current_dir)
    print("\n" + "="*80)
    print(f"[FINISH] All models done. Check 'process_log.txt' for results.")

if __name__ == "__main__":
    main()