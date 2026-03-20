import os
import shutil
import subprocess
from pathlib import Path

# ================= 設定區 =================
# 7-Zip 標準安裝路徑
SEVEN_ZIP_PATH = r"C:\Program Files\7-Zip\7z.exe"
# =========================================

def extract_archive(archive_path, extract_to):
    """使用 7-Zip 處理所有格式的解壓 (.zip, .7z, .rar)"""
    if not os.path.exists(SEVEN_ZIP_PATH):
        print(f"  [嚴重錯誤] 找不到 7-Zip 執行檔：{SEVEN_ZIP_PATH}")
        return False

    try:
        # 7-Zip 指令說明：
        # x: 解壓縮（包含完整路徑）
        # -o: 指定輸出資料夾 (注意 -o 與路徑之間不加空白)
        # -y: 自動回答 Yes (若有重複檔案直接覆蓋，避免程式卡死)
        cmd = [SEVEN_ZIP_PATH, "x", str(archive_path), f"-o{extract_to}", "-y"]
        
        # 執行指令並隱藏 7-Zip 的文字輸出 (stdout)
        subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
        return True
    except subprocess.CalledProcessError as e:
        print(f"  [錯誤] 解壓 {archive_path.name} 失敗：{e}")
        return False

def organize_3d_files(base_path):
    models_dir = Path(base_path) / "models"
    
    if not models_dir.exists():
        print(f"找不到路徑: {models_dir}")
        return

    # 支援的壓縮檔格式
    supported_exts = {'.zip', '.7z', '.rar'}

    for comp_folder in models_dir.iterdir():
        if not comp_folder.is_dir(): continue
            
        target_dir = comp_folder / "3D圖檔"
        if not target_dir.exists(): continue

        print(f"正在檢查組件: {comp_folder.name}")

        for archive_path in target_dir.iterdir():
            if archive_path.suffix.lower() not in supported_exts:
                continue

            part_name = archive_path.stem
            dest_stp = target_dir / f"{part_name}.stp"

            # 檢查是否已經有解壓好的 stp 檔
            if dest_stp.exists():
                # print(f"  [跳過] {part_name}.stp 已存在")
                continue

            # 開始解壓流程
            temp_extract_dir = target_dir / f"temp_{part_name}"
            
            if extract_archive(archive_path, temp_extract_dir):
                # 搜尋解壓出來的 stp 或 step (不分大小寫)
                stp_files = list(temp_extract_dir.rglob("*.[sS][tT][pP]*"))
                
                if stp_files:
                    src_stp = stp_files[0]
                    shutil.move(str(src_stp), str(dest_stp))
                    print(f"  [完成] 已從 {archive_path.suffix} 提取並重新命名為: {part_name}.stp")
                else:
                    print(f"  [警告] {archive_path.name} 內找不到 STP 檔案")
            
            # 清理臨時資料夾
            if temp_extract_dir.exists():
                shutil.rmtree(temp_extract_dir)

if __name__ == "__main__":
    # 執行位置
    organize_3d_files(".")
    print("\n[系統] 所有檔案檢查與 7-Zip 整理完畢！")