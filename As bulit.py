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
        "07 FAT reports": "09 FAT reports (Incl. punch list)",
        "Material certificates": "02 Material Traceability",
        "Welding Comsumable": "05 Welding Consumable",
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

    # 指定需要重命名檔案的資料夾
    folders_to_rename_files = [
        "02 Material Traceability",
        "04 Welding Identification Summary"
    ]

    # 正則表達式模式
    pattern = re.compile(r'(XB1#|XB[1-4][ABC]#|6S21[1-7]#|6S20[12356]#|XB3B\.002#|XB4B\.002#)\d+', re.IGNORECASE)

    def get_extension(filename):
        """取得檔案副檔名"""
        return os.path.splitext(filename)[1]

    def fix_incorrect_04_files(root_dir):
        """修復被錯誤重命名的 04 Welding Identification Summary 檔案"""
        for root, dirs, files in os.walk(root_dir):
            # 只處理 04 Welding Identification Summary 資料夾
            if os.path.basename(root) == "04 Welding Identification Summary":
                for filename in files:
                    if "Welding Identification Summary_Welding Identification Summary" in filename:
                        # 移除重複的部分
                        new_name = filename.replace(
                            "04 Welding Identification Summary_Welding Identification Summary_",
                            "04 Welding Identification Summary_"
                        )
                        old_path = os.path.join(root, filename)
                        new_path = os.path.join(root, new_name)
                        try:
                            os.rename(old_path, new_path)
                            print(f"修復檔案名稱：\n從：{filename}\n到：{new_name}")
                        except Exception as e:
                            print(f"重命名檔案時發生錯誤：{e}")

    def get_original_name(filename, folder):
        """嘗試從錯誤命名的檔案中提取原始資訊，移除所有重複的前綴及不必要的字串"""
        ext = get_extension(filename)
        basename = filename[:-len(ext)] if ext else filename

        if folder == "02 Material Traceability":
            # 移除 "Material Identification " 或 "Material Traceability " 這類字串
            basename = re.sub(r'^(?:Material Identification|Material Traceability)\s+', '', basename, flags=re.IGNORECASE)
            # 移除所有 "02 Material Traceability_" 前綴
            basename = re.sub(r'^(?:02 Material Traceability_)+', '', basename, flags=re.IGNORECASE)
            # 移除所有 "Material Traceability_" 前綴
            basename = re.sub(r'^(?:Material Traceability_)+', '', basename, flags=re.IGNORECASE)
            # 移除前導底線
            basename = re.sub(r'^_+', '', basename)
            # 添加正確的前綴
            return f"02 Material Traceability_{basename}"
        
        return basename

    def rename_file(folder, original_filename):
        """重新命名檔案（針對指定的資料夾）"""
        if folder in folders_to_rename_files:
            original_name = get_original_name(original_filename, folder)
            ext = get_extension(original_filename)
            new_filename = f"{original_name}{ext}"
            return new_filename
        return original_filename

    def is_filename_correct(folder, filename):
        """檢查檔案名稱是否符合預期格式"""
        expected_prefix = f"{folder}_"
        return filename.lower().startswith(expected_prefix.lower())

    def restore_original_name(folder, filename):
        """嘗試恢復被錯誤重命名的檔案名稱"""
        ext = get_extension(filename)
        basename = filename[:-len(ext)] if ext else filename

        # 檢查是否有資料夾名稱作為前綴，並移除
        pattern_prefix = re.compile(r'^\d{2}\s+[A-Za-z\s]+_', re.IGNORECASE)
        new_basename = pattern_prefix.sub('', basename)

        if new_basename != basename:
            # 名稱被修改過，進行還原
            original_name = new_basename + ext
            return original_name
        else:
            return filename  # 名稱未被修改，保持不變

    def fix_previous_errors(folder_path, folder_name):
        """修復之前的命名錯誤"""
        if os.path.exists(folder_path):
            for filename in os.listdir(folder_path):
                if os.path.isfile(os.path.join(folder_path, filename)):
                    if folder_name in folders_to_rename_files:
                        new_filename = rename_file(folder_name, filename)
                    else:
                        new_filename = restore_original_name(folder_name, filename)

                    if new_filename != filename:
                        old_file_path = os.path.join(folder_path, filename)
                        new_file_path = os.path.join(folder_path, new_filename)
                        try:
                            if not os.path.exists(new_file_path):
                                os.rename(old_file_path, new_file_path)
                                print(f"重命名檔案：{old_file_path} -> {new_file_path}")
                            else:
                                print(f"檔案已存在，無法重新命名：{new_file_path}")
                        except Exception as e:
                            print(f"重新命名檔案時發生錯誤：{e}")

    def process_directory(dir_path):
        """處理符合條件的資料夾"""
        print(f"處理資料夾：{dir_path}")
        if not os.path.exists(dir_path):
            print(f"資料夾不存在：{dir_path}")
            return

        # 處理資料夾重新命名
        for old_name in os.listdir(dir_path):
            old_path = os.path.join(dir_path, old_name)
            if os.path.isdir(old_path):
                new_name = old_name
                lower_old_name = old_name.lower().replace(' ', '')
                # 檢查重命名映射
                for key in rename_map:
                    key_normalized = key.lower().replace(' ', '')
                    if lower_old_name == key_normalized:
                        new_name = rename_map[key]
                        break
                if new_name != old_name:
                    new_path = os.path.join(dir_path, new_name)
                    try:
                        if not os.path.exists(new_path):
                            os.rename(old_path, new_path)
                            print(f"資料夾重新命名：{old_path} -> {new_path}")
                        else:
                            print(f"目標資料夾已存在，合併內容：{old_path} -> {new_path}")
                            for item in os.listdir(old_path):
                                shutil.move(os.path.join(old_path, item), new_path)
                            os.rmdir(old_path)
                    except Exception as e:
                        print(f"重新命名資料夾時發生錯誤：{e}")
                    old_name = new_name
                    old_path = new_path

        # 處理特殊資料夾移動和刪除
        old_punch_path = os.path.join(dir_path, "08 Punch list")
        new_punch_path = os.path.join(dir_path, "09 FAT reports (Incl. punch list)", "Punch list")
        if os.path.exists(old_punch_path):
            os.makedirs(os.path.dirname(new_punch_path), exist_ok=True)
            try:
                shutil.move(old_punch_path, new_punch_path)
                print(f"移動資料夾：{old_punch_path} -> {new_punch_path}")
            except Exception as e:
                print(f"移動資料夾時發生錯誤：{e}")

        folders_to_remove = ["Archive", "06 NCR", "08 Punch list"]
        for folder in folders_to_remove:
            folder_path = os.path.join(dir_path, folder)
            if os.path.exists(folder_path):
                try:
                    shutil.rmtree(folder_path)
                    print(f"刪除資料夾：{folder_path}")
                except Exception as e:
                    print(f"刪除資料夾時發生錯誤：{e}")

        material_traceability_path = os.path.join(dir_path, "02 Material Traceability")
        if os.path.exists(material_traceability_path):
            welding_consumable_path = os.path.join(material_traceability_path, "Welding Consumable")
            if os.path.exists(welding_consumable_path):
                new_welding_consumable_path = os.path.join(dir_path, "05 Welding Consumable")
                if not os.path.exists(new_welding_consumable_path):
                    try:
                        shutil.move(welding_consumable_path, new_welding_consumable_path)
                        print(f"移動並重新命名資料夾：{welding_consumable_path} -> {new_welding_consumable_path}")
                    except Exception as e:
                        print(f"移動資料夾時發生錯誤：{e}")
                else:
                    # 如果目標資料夾已存在，則合併內容
                    for item in os.listdir(welding_consumable_path):
                        try:
                            shutil.move(os.path.join(welding_consumable_path, item), new_welding_consumable_path)
                        except Exception as e:
                            print(f"移動檔案時發生錯誤：{e}")
                    try:
                        os.rmdir(welding_consumable_path)
                        print(f"刪除資料夾：{welding_consumable_path}")
                    except Exception as e:
                        print(f"刪除資料夾時發生錯誤：{e}")

            material_certificates_folder = os.path.join(material_traceability_path, "Material certificates")
            if os.path.exists(material_certificates_folder):
                for item in os.listdir(material_certificates_folder):
                    try:
                        shutil.move(os.path.join(material_certificates_folder, item), material_traceability_path)
                        print(f"移動檔案：{os.path.join(material_certificates_folder, item)} -> {material_traceability_path}")
                    except Exception as e:
                        print(f"移動檔案時發生錯誤：{e}")
                try:
                    os.rmdir(material_certificates_folder)
                    print(f"刪除資料夾：{material_certificates_folder}")
                except Exception as e:
                    print(f"刪除資料夾時發生錯誤：{e}")

        # 修復之前的命名錯誤並處理檔案重命名
        for folder in required_folders:
            folder_path = os.path.join(dir_path, folder)
            if os.path.exists(folder_path):
                print(f"正在處理資料夾：{folder_path}")
                fix_previous_errors(folder_path, folder)

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

        # 刪除空資料夾
        for folder_name in os.listdir(dir_path):
            folder_path = os.path.join(dir_path, folder_name)
            if os.path.isdir(folder_path):
                try:
                    if not os.listdir(folder_path):
                        os.rmdir(folder_path)
                        print(f"刪除空資料夾：{folder_path}")
                except Exception as e:
                    print(f"檢查或刪除資料夾時發生錯誤：{e}")

    def find_and_process_target_folders(current_dir):
        """遞迴查找並處理目標資料夾"""
        if not os.path.exists(current_dir):
            print(f"目錄不存在：{current_dir}")
            return
        for entry in os.listdir(current_dir):
            dir_path = os.path.join(current_dir, entry)
            if os.path.isdir(dir_path):
                folder_name = os.path.basename(dir_path)
                if pattern.search(folder_name):
                    print(f"找到目標資料夾：{dir_path}")
                    process_directory(dir_path)
                # 繼續在子目錄中查找
                find_and_process_target_folders(dir_path)

    # 從根目錄開始搜索
    print(f"開始處理根目錄：{root_dir}")
    find_and_process_target_folders(root_dir)

    # 修復 04 Welding Identification Summary 資料夾中的檔案名稱
    print("開始修復 04 Welding Identification Summary 資料夾中的檔案名稱。")
    fix_incorrect_04_files(root_dir)
    print("所有處理已完成。")

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