import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
import re

def rename_folders_and_remove_files(root_dir):
    # 定義資料夾重新命名對照表
    rename_map = {
        "01 Welding Summary": "04 Welding Identification Summary",
        "01 Welding Identification Summary": "04 Welding Identification Summary",
        "02 Material Traceability & Mill Cert": "02 Material Traceability",
        "03 Dimension Inspection Record": "07 Dimensional Reports",
        "05 Drawings": "01 Workshop Drawings",
        "05 Workshop Drawings": "01 Workshop Drawings",
        "07 FAT report": "09 FAT reports (Incl. punch list)",
        "Material certificates": "02 Material Traceability",
        "04 NDT Reports": "06 NDT Reports"
    }

    # 必要的資料夾列表
    required_folders = [
        "01 Workshop Drawings",
        "02 Material Traceability",
        "04 Welding Identification Summary",
        "05 Welding Consumable",
        "06 NDT Reports",
        "07 Dimensional Reports",
        "09 FAT reports (Incl. punch list)"
    ]

    # 調整後的目標資料夾名稱模式（正則表達式）
    target_folder_pattern = re.compile(r'^(6S201#|6S202#|6S205#)\d*\s*(V)?$', re.IGNORECASE)

    def get_extension(filename):
        """獲取檔案副檔名（小寫）"""
        return os.path.splitext(filename)[1].lower()

    def is_filename_correct(folder, filename):
        """檢查檔案名稱是否符合預期格式"""
        if folder == "07 Dimensional Reports":
            return filename.startswith("Dimensional Reports_") and "CWP" in filename
        elif folder in ["05 Welding Consumable", "06 NDT Reports", "09 FAT reports (Incl. punch list)"]:
            return True  # 這些資料夾不需要特定前綴
        elif folder == "02 Material Traceability":
            return filename.startswith("02 Material Traceability_") and "CWP" in filename
        elif folder == "04 Welding Identification Summary":
            return filename.startswith("04 Welding Identification Summary_") and "CWP" in filename
        else:
            folder_prefix = folder.split()[0]
            return filename.startswith(folder_prefix)

    def get_original_name(filename):
        """嘗試從錯誤命名的檔案中提取原始資訊"""
        # 獲取副檔名
        ext = get_extension(filename)
        # 移除副檔名以處理主檔名
        basename = os.path.splitext(filename)[0]
        
        # 移除所有可能的錯誤前綴
        clean_name = re.sub(r'^\d{2}\s+\d{2}\s+', '', basename)
        clean_name = re.sub(r'^\d{2}\s+', '', clean_name)
        clean_name = re.sub(r'^Material Identification Summary_', '', clean_name)
        clean_name = re.sub(r'^Material Traceability_', '', clean_name)
        clean_name = re.sub(r'^Dimensional Reports_', '', clean_name)
        clean_name = re.sub(r'^Welding Identification Summary_', '', clean_name)
        
        return clean_name, ext

    def rename_file(folder, filename):
        # 如果檔案名稱已經正確，就不需要更改
        if is_filename_correct(folder, filename):
            return filename

        # 取得資料夾的前綴編號和原始檔名資訊
        folder_prefix = folder.split()[0]
        original_name, ext = get_original_name(filename)
        
        # 根據不同資料夾類型重新命名
        if folder == "07 Dimensional Reports":
            if "CWP" in original_name:
                return f"Dimensional Reports_{original_name}{ext}"
            return filename  # 如果沒有 "CWP"，則保持原名
                
        elif folder in ["05 Welding Consumable", "06 NDT Reports", "09 FAT reports (Incl. punch list)"]:
            return filename  # 不需要重新命名
                
        elif folder == "02 Material Traceability":
            if "CWP" in original_name:
                return f"{folder_prefix} Material Traceability_{original_name}{ext}"
            return f"{folder_prefix} {original_name}{ext}"
                
        elif folder == "04 Welding Identification Summary":
            if "CWP" in original_name:
                return f"{folder_prefix} Welding Identification Summary_{original_name}{ext}"
            return f"{folder_prefix} {original_name}{ext}"
                
        else:
            return f"{folder_prefix} {original_name}{ext}"

    def fix_previous_errors(folder_path):
        """修復之前的命名錯誤"""
        if os.path.exists(folder_path):
            for filename in os.listdir(folder_path):
                if filename.endswith('.pdf.pdf'):  # 修復重複的 .pdf
                    new_name = filename[:-4]  # 移除一個 .pdf
                    old_path = os.path.join(folder_path, filename)
                    new_path = os.path.join(folder_path, new_name)
                    try:
                        os.rename(old_path, new_path)
                        print(f"修正重複的.pdf：{old_path} -> {new_path}")
                    except Exception as e:
                        print(f"修正檔案時發生錯誤：{e}")
                elif filename.endswith('.docx.pdf') or filename.endswith('.xlsx.pdf'):
                    # 修復錯誤添加的 .pdf
                    new_name = filename[:-4]
                    old_path = os.path.join(folder_path, filename)
                    new_path = os.path.join(folder_path, new_name)
                    try:
                        os.rename(old_path, new_path)
                        print(f"移除錯誤添加的.pdf：{old_path} -> {new_path}")
                    except Exception as e:
                        print(f"修正檔案時發生錯誤：{e}")

    def process_directory(dir_path):
        if not os.listdir(dir_path):
            print(f"跳過空資料夾：{dir_path}")
            return

        # 處理資料夾重新命名
        for old_name in list(os.listdir(dir_path)):
            old_path = os.path.join(dir_path, old_name)
            if os.path.isdir(old_path):
                new_name = old_name
                if old_name in rename_map:
                    new_name = rename_map[old_name]
                elif old_name in [folder.split(' ', 1)[1] for folder in required_folders]:
                    prefix = [folder.split(' ', 1)[0] for folder in required_folders if folder.endswith(old_name)][0]
                    new_name = f"{prefix} {old_name}"
                if new_name != old_name:
                    new_path = os.path.join(dir_path, new_name)
                    try:
                        if not os.path.exists(new_path):
                            os.rename(old_path, new_path)
                            print(f"資料夾重新命名：{old_path} -> {new_path}")
                    except Exception as e:
                        print(f"重新命名資料夾時發生錯誤：{e}")
                    old_name = new_name
                    old_path = new_path

        # 處理特殊資料夾移動和刪除
        old_punch_path = os.path.join(dir_path, "08 Punch list")
        new_punch_path = os.path.join(dir_path, "09 FAT reports (Incl. punch list)", "Punch list")
        if os.path.exists(old_punch_path):
            os.makedirs(os.path.dirname(new_punch_path), exist_ok=True)
            shutil.move(old_punch_path, new_punch_path)
            print(f"移動資料夾：{old_punch_path} -> {new_punch_path}")

        folders_to_remove = ["Archive", "06 NCR", "08 Punch list"]
        for folder in folders_to_remove:
            folder_path = os.path.join(dir_path, folder)
            if os.path.exists(folder_path):
                shutil.rmtree(folder_path)
                print(f"刪除資料夾：{folder_path}")

        material_traceability_path = os.path.join(dir_path, "02 Material Traceability")
        if os.path.exists(material_traceability_path):
            welding_consumable_path = os.path.join(material_traceability_path, "Welding Consumable")
            if os.path.exists(welding_consumable_path):
                new_welding_consumable_path = os.path.join(dir_path, "05 Welding Consumable")
                if not os.path.exists(new_welding_consumable_path):
                    shutil.move(welding_consumable_path, new_welding_consumable_path)
                    print(f"移動並重新命名資料夾：{welding_consumable_path} -> {new_welding_consumable_path}")

            material_certificates_folder = os.path.join(material_traceability_path, "Material certificates")
            if os.path.exists(material_certificates_folder):
                shutil.rmtree(material_certificates_folder)
                print(f"刪除資料夾：{material_certificates_folder}")

        # 首先修復之前的命名錯誤
        for folder in required_folders:
            folder_path = os.path.join(dir_path, folder)
            fix_previous_errors(folder_path)

        # 然後處理檔案重新命名
        for folder in required_folders:
            folder_path = os.path.join(dir_path, folder)
            if os.path.exists(folder_path):
                for filename in os.listdir(folder_path):
                    old_file_path = os.path.join(folder_path, filename)
                    if os.path.isfile(old_file_path):
                        # 檢查檔案名稱是否需要修正
                        if not is_filename_correct(folder, filename):
                            new_filename = rename_file(folder, filename)
                            new_file_path = os.path.join(folder_path, new_filename)
                            try:
                                if os.path.exists(new_file_path):
                                    print(f"檔案已存在，跳過重新命名：{old_file_path}")
                                else:
                                    os.rename(old_file_path, new_file_path)
                                    print(f"修正檔案命名：{old_file_path} -> {new_file_path}")
                            except Exception as e:
                                print(f"重新命名檔案時發生錯誤：{e}")

        # 刪除特定檔案
        for root, dirs, files in os.walk(dir_path):
            for file in files:
                file_lower = file.lower()
                if ((file_lower.endswith('.xlsx') and 'welding' in file_lower) or 
                    (file_lower.endswith('.docx') and 'material' in file_lower)):
                    file_path = os.path.join(root, file)
                    try:
                        os.remove(file_path)
                        print(f"刪除檔案：{file_path}")
                    except Exception as e:
                        print(f"刪除檔案時發生錯誤：{e}")

    def find_and_process_target_folders(current_dir):
        for entry in os.listdir(current_dir):
            dir_path = os.path.join(current_dir, entry)
            if os.path.isdir(dir_path):
                if entry.startswith('XB') or target_folder_pattern.match(entry):
                    process_directory(dir_path)
                else:
                    # 繼續在子目錄中查找
                    find_and_process_target_folders(dir_path)

    # 從根目錄開始搜索
    find_and_process_target_folders(root_dir)

def select_directory():
    root = tk.Tk()
    root.withdraw()

    directory = filedialog.askdirectory(title="選擇要處理的根目錄")
    
    if directory:
        confirm = messagebox.askyesno("確認", 
            f"您選擇的目錄是：\n{directory}\n\n"
            "此操作將：\n"
            "1. 修正之前命名錯誤的檔案\n"
            "2. 移除錯誤添加的.pdf副檔名\n"
            "3. 重新組織資料夾結構\n"
            "4. 刪除指定的檔案\n\n"
            "確定要繼續嗎？")
        
        if confirm:
            rename_folders_and_remove_files(directory)
            messagebox.showinfo("完成", "資料夾重新命名、重組和檔案刪除已完成。")
        else:
            messagebox.showinfo("取消", "操作已取消。")
    else:
        messagebox.showinfo("取消", "沒有選擇目錄，程式結束。")

if __name__ == "__main__":
    select_directory()
