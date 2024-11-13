import os
import shutil
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox
import re
from difflib import get_close_matches

def rename_file_if_needed(file_path):
    """檢查檔案名稱中是否包含 CWP06G-XB4C 並取代"""
    file_name = os.path.basename(file_path)
    if "CWP06G-XB4C" in file_name:
        new_file_name = file_name.replace("CWP06G-XB4C", "CWP06C-XB4C")
        new_file_path = os.path.join(os.path.dirname(file_path), new_file_name)
        os.rename(file_path, new_file_path)
        return new_file_path
    return file_path

def check_and_rename_files_in_folder(folder_path):
    renamed_files = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            file_path = os.path.join(root, file)
            new_file_path = rename_file_if_needed(file_path)
            if new_file_path != file_path:
                renamed_files.append((file_path, new_file_path))
    return renamed_files

def get_ndt_codes_from_pdf(file_path):
    codes_with_filenames = dict()
    pattern = re.compile(r'CWPQRJKNDT(\d+)')

    doc = fitz.open(file_path)
    text = ""
    for page in doc:
        text += page.get_text()

    matches = pattern.findall(text)
    if matches:
        codes_with_filenames.update({code: f'CWP-Q-R-JK-NDT-{code}.pdf' for code in matches})

    return codes_with_filenames

def get_welding_codes_from_pdf(file_path):
    codes = set()
    pattern = re.compile(r'\b\d{6,10}\b')  # 匹配六位到十位數字

    doc = fitz.open(file_path)
    text = ""
    for page in doc:
        text += page.get_text()

    matches = pattern.findall(text)
    if matches:
        codes.update(matches)

    return codes

def process_pdf_files_in_folder(folder, is_as_built):
    ndt_codes_with_filenames_total = {}
    welding_codes_total = set()
    target_folder_name = "04 Welding Identification Summary" if is_as_built else "01 Welding Identification Summary"

    for root, dirs, files in os.walk(folder):
        if target_folder_name not in dirs:
            close_matches = get_close_matches(target_folder_name, dirs, n=1, cutoff=0.6)
            if close_matches:
                similar_folder = close_matches[0]
                if is_as_built:
                    target_folder_name = similar_folder
                else:
                    response = messagebox.askyesnocancel("資料夾名稱不符",
                                                         f"未找到 '{target_folder_name}' 資料夾，但找到相似的資料夾 '{similar_folder}'。\n"
                                                         f"是否要將 '{similar_folder}' 重新命名為 '{target_folder_name}'？\n"
                                                         f"選擇「是」將重命名資料夾，選擇「否」將使用現有資料夾而不重命名，選擇「取消」將終止程序。")
                    if response is True:  # 用戶選擇「是」
                        os.rename(os.path.join(root, similar_folder), os.path.join(root, target_folder_name))
                        dirs[dirs.index(similar_folder)] = target_folder_name
                    elif response is False:  # 用戶選擇「否」
                        target_folder_name = similar_folder
                    else:  # 用戶選擇「取消」或關閉對話框
                        messagebox.showwarning("警告", f"未找到 '{target_folder_name}' 資料夾，程序將終止執行。")
                        raise SystemExit
            else:
                if is_as_built:
                    os.mkdir(os.path.join(root, target_folder_name))
                    dirs.append(target_folder_name)
                else:
                    response = messagebox.askyesno("資料夾名稱不符",
                                                   f"未找到 '{target_folder_name}' 資料夾。\n是否要建立該資料夾？")
                    if response:
                        os.mkdir(os.path.join(root, target_folder_name))
                        dirs.append(target_folder_name)
                    else:
                        messagebox.showwarning("警告", f"未找到 '{target_folder_name}' 資料夾，程序將終止執行。")
                        raise SystemExit

        summary_folder = os.path.join(root, target_folder_name)
        for summary_file in os.listdir(summary_folder):
            if summary_file.endswith('.pdf') and not summary_file.startswith('~$'):
                summary_path = os.path.join(summary_folder, summary_file)
                summary_path = rename_file_if_needed(summary_path)

                ndt_codes_with_filenames = get_ndt_codes_from_pdf(summary_path)
                welding_codes = get_welding_codes_from_pdf(summary_path)
                ndt_codes_with_filenames_total.update(ndt_codes_with_filenames)
                welding_codes_total.update(welding_codes)
        break  # 只處理第一層子資料夾

    return ndt_codes_with_filenames_total, welding_codes_total


def search_and_copy_ndt_pdfs(source_folder, target_folder, codes_with_filenames, is_as_built):
    pattern = re.compile(r'CWP-Q-R-JK-NDT-(\d+)\(.*?\)\s?\(已完成\)?.pdf')
    copied_files = 0
    not_found_filenames = set(codes_with_filenames.values())

    target_folder_path = os.path.join(target_folder, '06 NDT Reports' if is_as_built else '04 NDT Reports')
    os.makedirs(target_folder_path, exist_ok=True)

    for root, dirs, files in os.walk(source_folder):
        for file in files:
            if file.endswith('.pdf') and not file.startswith('~$'):
                source_file_path = os.path.join(root, file)
                source_file_path = rename_file_if_needed(source_file_path)

                match = pattern.match(file)
                if match:
                    ndt_code = match.group(1)
                    if ndt_code in codes_with_filenames:
                        not_found_filenames.discard(codes_with_filenames[ndt_code])
                        shutil.copy2(source_file_path, os.path.join(target_folder_path, file))
                        copied_files += 1

    return copied_files, not_found_filenames

def search_and_copy_welding_pdfs(source_folder, target_folder, codes, is_as_built):
    copied_files = 0
    not_found_codes = set(codes)
    if is_as_built:
        target_folder_path = os.path.join(target_folder, '05 Welding Consumable')
    else:
        target_folder_path = os.path.join(target_folder, '02 Material Traceability & Mill Cert', 'Welding Consumable')
    os.makedirs(target_folder_path, exist_ok=True)

    for code in codes:
        found = False
        pattern = re.compile(re.escape(code))
        for root, dirs, files in os.walk(source_folder):
            for file in files:
                if file.endswith('.pdf') and not file.startswith('~$'):
                    source_file_path = os.path.join(root, file)
                    source_file_path = rename_file_if_needed(source_file_path)

                    if pattern.search(file):
                        not_found_codes.discard(code)
                        target_file_path = os.path.join(target_folder_path, file)
                        shutil.copy2(source_file_path, target_file_path)
                        copied_files += 1
                        found = True
                        break
            if found:
                break
    return copied_files, not_found_codes
def delete_all_welding_pdfs(target_folder, is_as_built):
    deleted_files = 0
    if is_as_built:
        target_folder_path = os.path.join(target_folder, '05 Welding Consumable')
    else:
        target_folder_path = os.path.join(target_folder, '02 Material Traceability & Mill Cert', 'Welding Consumable')

    if os.path.exists(target_folder_path):
        for file in os.listdir(target_folder_path):
            if file.endswith('.pdf') and not file.startswith('~$'):
                file_path = os.path.join(target_folder_path, file)
                os.remove(file_path)
                deleted_files += 1

    return deleted_files

def delete_all_ndt_pdfs(target_folder, is_as_built):
    deleted_files = 0
    target_folder_path = os.path.join(target_folder, '06 NDT Reports' if is_as_built else '04 NDT Reports')

    if os.path.exists(target_folder_path):
        for file in os.listdir(target_folder_path):
            if file.endswith('.pdf') and not file.startswith('~$'):
                file_path = os.path.join(target_folder_path, file)
                os.remove(file_path)
                deleted_files += 1

    return deleted_files

def extract_base_folder_name(folder_name):
    """
    提取資料夾名稱，統一處理 XB 系列和 6S 系列的命名格式，包含錯誤處理
    例如：
    - XB4C.001#001
    - 6S201.001#021
    """
    if not folder_name:  # 加入空值檢查
        return ""
        
    # 統一格式模式
    patterns = [
        r'(XB[1-4][ABC](?:\.\d{3})?#\d+)',  # XB 系列，流水號可選
        r'(6S(?:20[1-6]|21[1-7])(?:\.\d{3})?#?\d*)'  # 6S 系列，流水號可選
    ]
    
    combined_pattern = '|'.join(patterns)
    match = re.search(combined_pattern, str(folder_name))
    if match:
        return match.group(1)
    return str(folder_name)  # 確保返回字串

def is_target_folder(folder_name):
    """
    檢查是否為目標資料夾
    """
    if not folder_name:  # 加入空值檢查
        return False
        
    patterns = [
        r'XB[1-4][ABC](?:\.\d{3})?#\d+',
        r'6S(?:20[1-6]|21[1-7])(?:\.\d{3})?#?\d*'
    ]
    
    combined_pattern = '|'.join(patterns)
    return bool(re.match(combined_pattern, str(folder_name)))

def clean_unmatched_files(pdf_folder, is_as_built):
    """
    改進的檔案匹配邏輯，加入錯誤處理
    """
    deleted_files = []
    
    try:
        folder_name = os.path.basename(pdf_folder)
        base_folder_name = extract_base_folder_name(folder_name)
        
        if not base_folder_name:  # 檢查是否為空字串
            print(f"警告：無法從資料夾名稱中提取基本名稱: {folder_name}")
            return deleted_files
        
        # 建立檔案名稱匹配模式
        base_patterns = [
            re.escape(base_folder_name),  # 完全匹配
            r'CWP\d+[A-Z]-' + re.escape(base_folder_name),  # 允許 CWP 前綴
            # 統一的匹配模式，處理基本名稱和可選的流水號
            r'.*?' + re.escape(base_folder_name.split('.')[0]) + r'(?:\.\d{3})?(?:#\d+)?' + r'.*'
        ]
        
        combined_pattern = '|'.join(f'({pattern})' for pattern in base_patterns)
        
        if is_as_built:
            target_folders = ["04 Welding Identification Summary", "03 Material Traceability & Mill Cert"]
        else:
            target_folders = ["01 Welding Identification Summary", "02 Material Traceability & Mill Cert"]

        print(f"資料夾基本名稱: {base_folder_name}")
        print(f"使用的匹配模式: {combined_pattern}")

        for subfolder in target_folders:
            subfolder_path = os.path.join(pdf_folder, subfolder)
            if os.path.exists(subfolder_path):
                for root, dirs, files in os.walk(subfolder_path):
                    for file in files:
                        if file.endswith('.pdf') and not file.startswith('~$'):
                            file_path = os.path.join(root, file)
                            file_name_without_extension = os.path.splitext(file)[0]
                            
                            print(f"檢查檔案: {file_name_without_extension}")
                            
                            try:
                                # 檢查檔案名稱是否符合模式
                                if not re.search(combined_pattern, file_name_without_extension, re.IGNORECASE):
                                    if messagebox.askyesno("確認刪除", 
                                                         f"是否要刪除檔案：{file}\n" +
                                                         f"資料夾名稱：{base_folder_name}\n" +
                                                         "此檔案似乎不符合命名規則。"):
                                        os.remove(file_path)
                                        deleted_files.append(file_path)
                                        print(f"已刪除檔案: {file_path}")
                                    else:
                                        print(f"使用者選擇保留檔案: {file_path}")
                            except Exception as e:
                                print(f"處理檔案時發生錯誤 {file_path}: {str(e)}")
                    break
                    
    except Exception as e:
        print(f"處理資料夾時發生錯誤 {pdf_folder}: {str(e)}")
        messagebox.showerror("錯誤", f"處理資料夾時發生錯誤：\n{str(e)}")
    
    return deleted_files

def process_folders(pdf_folder, ndt_source_pdf_folder, welding_source_pdf_folder):
    """
    處理多個資料夾，加入錯誤處理
    """
    total_ndt_copied = 0
    total_welding_copied = 0
    not_found_ndt_filenames_total = set()
    not_found_welding_codes_total = set()
    reasons_total = []
    deleted_files_total = []

    try:
        is_as_built = "As-Built" in pdf_folder

        renamed_files = check_and_rename_files_in_folder(pdf_folder)
        if renamed_files:
            message = "以下檔案已被重新命名：\n"
            for old_path, new_path in renamed_files:
                message += f"舊名稱：{os.path.basename(old_path)} -> 新名稱：{os.path.basename(new_path)}\n"
            messagebox.showinfo("檔案重命名", message)

        if is_target_folder(os.path.basename(pdf_folder)):
            results = process_single_folder(pdf_folder, ndt_source_pdf_folder, welding_source_pdf_folder, is_as_built)
            ndt_copied, welding_copied, not_found_ndt_filenames, not_found_welding_codes, reasons, deleted_files = results

            total_ndt_copied += ndt_copied
            total_welding_copied += welding_copied
            not_found_ndt_filenames_total.update(not_found_ndt_filenames)
            not_found_welding_codes_total.update(not_found_welding_codes)
            reasons_total.extend(reasons)
            deleted_files_total.extend(deleted_files)
        else:
            for root, dirs, files in os.walk(pdf_folder):
                for dir in dirs:
                    if is_target_folder(dir):
                        subfolder_path = os.path.join(root, dir)
                        try:
                            results = process_single_folder(subfolder_path, ndt_source_pdf_folder, welding_source_pdf_folder, is_as_built)
                            ndt_copied, welding_copied, not_found_ndt_filenames, not_found_welding_codes, reasons, deleted_files = results

                            total_ndt_copied += ndt_copied
                            total_welding_copied += welding_copied
                            not_found_ndt_filenames_total.update(not_found_ndt_filenames)
                            not_found_welding_codes_total.update(not_found_welding_codes)
                            reasons_total.extend(reasons)
                            deleted_files_total.extend(deleted_files)
                        except Exception as e:
                            print(f"處理資料夾時發生錯誤 {subfolder_path}: {str(e)}")
                break  # 只處理第一層子資料夾

    except Exception as e:
        print(f"處理資料夾時發生錯誤: {str(e)}")
        messagebox.showerror("錯誤", f"處理資料夾時發生錯誤：\n{str(e)}")

    return (total_ndt_copied, total_welding_copied, not_found_ndt_filenames_total, 
            not_found_welding_codes_total, reasons_total, deleted_files_total)
def main():
    root = tk.Tk()
    root.withdraw()

    pdf_initialdir = "C:/Users/CWP-PC-E-COM302/Box/T460 風電 品管 簡瑞成/FAT package"
    if not os.path.exists(pdf_initialdir):
        pdf_initialdir = os.path.expanduser("~")

    pdf_folder = filedialog.askdirectory(
        initialdir=pdf_initialdir,
        title="請選擇包含銲道追溯檔案的資料夾:")

    ndt_source_pdf_initialdir = "U:/N-品管部/@品管部共用資料區/專案/CWP06 台電二期專案/報驗單/001_JK報驗單/完成 (pdf)"
    if not os.path.exists(ndt_source_pdf_initialdir):
        ndt_source_pdf_initialdir = os.path.expanduser("~")

    ndt_source_pdf_folder = filedialog.askdirectory(
        initialdir=ndt_source_pdf_initialdir,
        title="請選擇包含報驗單 PDF 檔案的資料夾:")

    welding_source_pdf_initialdir = "U:/N-品管部/@品管部共用資料區/品管人員資料夾/T460 風電 品管 簡瑞成/焊材材證"
    if not os.path.exists(welding_source_pdf_initialdir):
        welding_source_pdf_initialdir = os.path.expanduser("~")

    welding_source_pdf_folder = filedialog.askdirectory(
        initialdir=welding_source_pdf_initialdir,
        title="請選擇包含焊材材證 PDF 檔案的資料夾:")

    is_as_built = "As-Built" in pdf_folder
    mode = "竣工模式" if is_as_built else "一般模式"

    total_ndt_copied, total_welding_copied, not_found_ndt_filenames_total, not_found_welding_codes_total, reasons_total, deleted_files_total = process_folders(pdf_folder, ndt_source_pdf_folder, welding_source_pdf_folder)

    message = f"執行模式: {mode}\n\n"
    message += f"報驗單: 共複製了 {total_ndt_copied} 份。\n"
    message += f"焊材材證: 共複製了 {total_welding_copied} 份。\n\n"

    if not_found_ndt_filenames_total:
        missing_ndt_filenames_str = ", ".join(not_found_ndt_filenames_total)
        message += f"以下報驗單編號的檔案在報驗單資料夾中未找到，有可能尚未上傳：{missing_ndt_filenames_str}\n\n"

    if not_found_welding_codes_total:
        missing_welding_codes_str = ", ".join(not_found_welding_codes_total)
        message += f"以下焊材材證編號的檔案在焊材材證資料夾中未找到，有可能輸入有誤：{missing_welding_codes_str}"

    # 如果兩者皆為 0，則顯示具體原因
    if total_ndt_copied == 0 and total_welding_copied == 0 and reasons_total:
        message += "\n\n沒有檔案被複製，具體原因如下：\n"
        message += "\n".join(reasons_total)

    if deleted_files_total:
        message += "\n\n以下檔案因為檔名不符合已被刪除：\n"
        for file_path in deleted_files_total:
            message += f"{file_path}\n"

    messagebox.showinfo("完成", message)

if __name__ == "__main__":
    main()
