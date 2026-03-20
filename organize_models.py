import os
import shutil
import subprocess
import re
from pathlib import Path
from collections import defaultdict

# ================= 設定區 =================
SEVEN_ZIP_PATH = r"C:\Program Files\7-Zip\7z.exe"
# =========================================

def extract_archive(archive_path, extract_to):
    if not os.path.exists(SEVEN_ZIP_PATH):
        print(f"  [嚴重錯誤] 找不到 7-Zip：{SEVEN_ZIP_PATH}")
        return False
    try:
        cmd = [SEVEN_ZIP_PATH, "x", str(archive_path), f"-o{extract_to}", "-y"]
        subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
        return True
    except Exception as e:
        print(f"  [錯誤] 解壓 {archive_path.name} 失敗：{e}")
        return False

def parse_version(filename):
    """
    解析檔名版本資訊。
    格式: M24217-3D-A01 -> ('M24217', 1, 1)  (A=1, R=2)
    """
    match = re.search(r"(.+)-3D-([RA])(\d+)", filename, re.IGNORECASE)
    if match:
        prefix = match.group(1).upper()
        rev_letter = match.group(2).upper()
        version_num = int(match.group(3))
        rev_weight = 2 if rev_letter == 'R' else 1
        return prefix, rev_weight, version_num
    return filename.upper(), 0, 0

def organize_3d_files(base_path):
    models_dir = Path(base_path) / "models"
    if not models_dir.exists(): return

    supported_exts = {'.zip', '.7z', '.rar'}

    for comp_folder in models_dir.iterdir():
        if not comp_folder.is_dir(): continue
        target_dir = comp_folder / "3D圖檔"
        if not target_dir.exists(): continue

        print(f"正在檢查組件: {comp_folder.name}")

        # 1. 收集所有壓縮檔並分組
        archive_groups = defaultdict(list)
        for f in target_dir.iterdir():
            if f.suffix.lower() in supported_exts:
                prefix, rev_w, ver = parse_version(f.stem)
                archive_groups[prefix].append({
                    'path': f,
                    'weight': (rev_w, ver)
                })

        # 2. 處理每一組零件
        for prefix, files in archive_groups.items():
            # 挑選出最強版本
            latest_file_info = sorted(files, key=lambda x: x['weight'])[-1]
            latest_archive = latest_file_info['path']
            dest_stp = target_dir / f"{prefix}.stp"

            # --- 核心邏輯：刪除所有舊的或格式不符的 stp ---
            # 只要檔名開頭是這個 prefix 且副檔名是 stp 的通通刪掉 (準備放新的)
            for old_stp in target_dir.glob(f"{prefix}*.st[ep]*"):
                try:
                    # 如果目前的最新版 stp 已經存在且檔案大小沒變，可選跳過
                    # 但為了保證最新，這裡採取「先刪除再重新產生」的策略
                    os.remove(old_stp)
                    print(f"  [清理] 已刪除舊版檔案: {old_stp.name}")
                except Exception as e:
                    print(f"  [錯誤] 無法刪除 {old_stp.name}: {e}")

            # 3. 執行解壓
            print(f"  [解壓] 正在提取最新版本: {latest_archive.name}")
            temp_extract_dir = target_dir / f"temp_{prefix}"
            
            if extract_archive(latest_archive, temp_extract_dir):
                stp_files = list(temp_extract_dir.rglob("*.[sS][tT][pP]*"))
                if stp_files:
                    src_stp = stp_files[0]
                    shutil.move(str(src_stp), str(dest_stp))
                    print(f"  [成功] 最終保留版本: {dest_stp.name}")
                else:
                    print(f"  [警告] {latest_archive.name} 內找不到 STP")
            
            if temp_extract_dir.exists():
                shutil.rmtree(temp_extract_dir)

if __name__ == "__main__":
    organize_3d_files(".")
    print("\n[系統] 舊版刪除與最新版整理完成！")