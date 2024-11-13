import os
import shutil
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox
import re
from difflib import get_close_matches
import json
import time
from concurrent.futures import ThreadPoolExecutor, ProcessPoolExecutor, as_completed
from functools import partial
import threading

CACHE_FILE = 'folder_cache.json'
CACHE_EXPIRY = 24 * 3600  # 快取有效期為 24 小時

def load_cache():
    """載入快取資料夾結構"""
    if os.path.exists(CACHE_FILE):
        cache_mtime = os.path.getmtime(CACHE_FILE)
        if time.time() - cache_mtime < CACHE_EXPIRY:
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        else:
            print("快取已過期，重新建立快取。")
    return {}

def save_cache(cache):
    """儲存快取資料夾結構"""
    with open(CACHE_FILE, 'w', encoding='utf-8') as f:
        json.dump(cache, f, ensure_ascii=False, indent=4)

def build_folder_cache(folder_path):
    """建立資料夾結構快取"""
    cache = {}
    for root, dirs, files in os.walk(folder_path):
        relative_root = os.path.relpath(root, folder_path)
        cache[relative_root] = {}
        cache[relative_root]['dirs'] = dirs
        cache[relative_root]['files'] = {}
        for file in files:
            file_path = os.path.join(root, file)
            try:
                mtime = os.path.getmtime(file_path)
                cache[relative_root]['files'][file] = mtime
            except Exception as e:
                print(f"無法取得檔案 {file_path} 的修改時間: {e}")
    return cache

def update_folder_cache(folder_path, cache):
    """更新快取資料夾結構"""
    updated = False
    for root, dirs, files in os.walk(folder_path):
        relative_root = os.path.relpath(root, folder_path)
        if relative_root not in cache:
            cache[relative_root] = {'dirs': dirs, 'files': {}}
            updated = True
            print(f"新增資料夾至快取: {relative_root}")
        else:
            # 檢查新增或刪除的資料夾
            existing_dirs = set(cache[relative_root]['dirs'])
            current_dirs = set(dirs)
            added_dirs = current_dirs - existing_dirs
            removed_dirs = existing_dirs - current_dirs
            if added_dirs or removed_dirs:
                cache[relative_root]['dirs'] = dirs
                updated = True
                print(f"更新資料夾列表: {relative_root}")

        # 檢查檔案變更
        existing_files = set(cache[relative_root]['files'].keys())
        current_files = set(files)
        added_files = current_files - existing_files
        removed_files = existing_files - current_files
        common_files = existing_files & current_files

        if added_files:
            for file in added_files:
                file_path = os.path.join(root, file)
                try:
                    mtime = os.path.getmtime(file_path)
                    cache[relative_root]['files'][file] = mtime
                    updated = True
                    print(f"新增檔案至快取: {os.path.join(relative_root, file)}")
                except Exception as e:
                    print(f"無法取得檔案 {file_path} 的修改時間: {e}")

        if removed_files:
            for file in removed_files:
                del cache[relative_root]['files'][file]
                updated = True
                print(f"從快取中移除檔案: {os.path.join(relative_root, file)}")

        for file in common_files:
            file_path = os.path.join(root, file)
            try:
                mtime = os.path.getmtime(file_path)
                if cache[relative_root]['files'][file] != mtime:
                    cache[relative_root]['files'][file] = mtime
                    updated = True
                    print(f"更新檔案修改時間至快取: {os.path.join(relative_root, file)}")
            except Exception as e:
                print(f"無法取得檔案 {file_path} 的修改時間: {e}")

    return updated

def get_cached_files(folder_path, cache):
    """從快取中獲取所有檔案路徑"""
    files = []
    for relative_root, data in cache.items():
        for file in data['files']:
            files.append(os.path.join(folder_path, relative_root, file))
    return files

def rename_file_if_needed(file_path):
    """檢查檔案名稱中是否包含 CWP06G-XB4C 並取代"""
    file_name = os.path.basename(file_path)
    if "CWP06G-XB4C" in file_name:
        new_file_name = file_name.replace("CWP06G-XB4C", "CWP06C-XB4C")
        new_file_path = os.path.join(os.path.dirname(file_path), new_file_name)
        try:
            os.rename(file_path, new_file_path)
            print(f"檔案已重新命名: {file_path} -> {new_file_path}")
            return new_file_path, True
        except Exception as e:
            print(f"無法重新命名檔案 {file_path}: {e}")
    return file_path, False

def check_and_rename_file(file_path):
    """檢查並重新命名單一檔案"""
    new_file_path, renamed = rename_file_if_needed(file_path)
    return (file_path, new_file_path) if renamed else None

def check_and_rename_files_in_folder_cached(folder_path, cache, max_workers):
    """基於快取檢查並重新命名檔案"""
    renamed_files = []
    files = get_cached_files(folder_path, cache)
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(check_and_rename_file, file_path) for file_path in files]
        for future in as_completed(futures):
            result = future.result()
            if result:
                renamed_files.append(result)
    return renamed_files

def get_ndt_codes_from_pdf(file_path):
    """從 PDF 中提取 NDT 編號"""
    codes_with_filenames = {}
    pattern = re.compile(r'CWPQRJKNDT(\d+)', re.IGNORECASE)

    try:
        doc = fitz.open(file_path)
        text = ""
        for page in doc:
            text += page.get_text()
        doc.close()

        matches = pattern.findall(text)
        if matches:
            codes_with_filenames.update({code: f'CWP-Q-R-JK-NDT-{code}.pdf' for code in matches})
            print(f"從 {file_path} 提取到 NDT 編號: {matches}")
        else:
            print(f"從 {file_path} 未提取到任何 NDT 編號。")

    except Exception as e:
        print(f"無法讀取 PDF 檔案 {file_path}: {e}")

    return codes_with_filenames

def get_welding_codes_from_pdf(file_path):
    """從 PDF 中提取焊材材證編號"""
    codes = set()
    pattern = re.compile(r'\b\d{6,10}\b')  # 匹配六位到十位數字

    try:
        doc = fitz.open(file_path)
        text = ""
        for page in doc:
            text += page.get_text()
        doc.close()

        matches = pattern.findall(text)
        if matches:
            codes.update(matches)
            print(f"從 {file_path} 提取到焊材材證編號: {matches}")
        else:
            print(f"從 {file_path} 未提取到任何焊材材證編號。")

    except Exception as e:
        print(f"無法讀取 PDF 檔案 {file_path}: {e}")

    return codes

def is_target_folder(folder_name):
    """判斷是否為目標資料夾，並支援如 '6S201#021' 的格式"""
    pattern = re.compile(r'XB1#\d+|XB[1-4][ABC]#\d+|6S21[1-7]#\d+|6S20[1256]#\d+', re.IGNORECASE)
    return pattern.match(folder_name) is not None

def extract_base_folder_name(folder_name):
    """從資料夾名稱中提取基本名稱，支援如 '6S201#021' 的格式"""
    pattern = re.compile(r'(XB1#\d+|XB[1-4][ABC]#\d+|6S21[1-7]#\d+|6S20[1256]#\d+)', re.IGNORECASE)
    match = pattern.search(folder_name)
    if match:
        return match.group(1)
    else:
        return folder_name

def process_pdf_file(summary_path):
    """處理單一 PDF 檔案，提取 NDT 和焊材材證編號"""
    ndt_codes_with_filenames = get_ndt_codes_from_pdf(summary_path)
    welding_codes = get_welding_codes_from_pdf(summary_path)
    return ndt_codes_with_filenames, welding_codes

def process_pdf_files_in_folder(folder, is_as_built, cache, max_workers):
    """處理指定資料夾中的 PDF 檔案，提取 NDT 和焊材材證編號"""
    ndt_codes_with_filenames_total = {}
    welding_codes_total = set()
    target_folder_name = "04 Welding Identification Summary" if is_as_built else "01 Welding Identification Summary"

    # 更新快取
    if update_folder_cache(folder, cache):
        save_cache(cache)
        print("快取已更新。")

    # 取得目標資料夾
    summary_folders = []
    for root, dirs, files in os.walk(folder):
        if target_folder_name in dirs:
            summary_folders.append(os.path.join(root, target_folder_name))
        else:
            # 檢查相似資料夾
            close_matches = get_close_matches(target_folder_name, dirs, n=1, cutoff=0.6)
            if close_matches:
                similar_folder = close_matches[0]
                if is_as_built:
                    target_folder_name = similar_folder
                else:
                    response = messagebox.askyesnocancel(
                        "資料夾名稱不符",
                        f"未找到 '{target_folder_name}' 資料夾，但找到相似的資料夾 '{similar_folder}'。\n"
                        f"是否要將 '{similar_folder}' 重新命名為 '{target_folder_name}'？\n"
                        f"選擇「是」將重命名資料夾，選擇「否」將使用現有資料夾而不重命名，選擇「取消」將終止程序。"
                    )
                    if response is True:  # 用戶選擇「是」
                        os.rename(os.path.join(root, similar_folder), os.path.join(root, target_folder_name))
                        dirs[dirs.index(similar_folder)] = target_folder_name
                        print(f"資料夾已重新命名: {similar_folder} -> {target_folder_name}")
                    elif response is False:  # 用戶選擇「否」
                        target_folder_name = similar_folder
                        print(f"使用現有相似資料夾: {similar_folder}")
                    else:  # 用戶選擇「取消」或關閉對話框
                        messagebox.showwarning("警告", f"未找到 '{target_folder_name}' 資料夾，程序將終止執行。")
                        raise SystemExit
                summary_folders.append(os.path.join(root, target_folder_name))
        break  # 只處理第一層子資料夾

    # 如果找不到目標資料夾，根據模式建立
    if not summary_folders:
        if is_as_built:
            new_folder = os.path.join(folder, target_folder_name)
            os.mkdir(new_folder)
            summary_folders.append(new_folder)
            print(f"已建立資料夾: {target_folder_name}")
        else:
            response = messagebox.askyesno(
                "資料夾名稱不符",
                f"未找到 '{target_folder_name}' 資料夾。\n是否要建立該資料夾？"
            )
            if response:
                new_folder = os.path.join(folder, target_folder_name)
                os.mkdir(new_folder)
                summary_folders.append(new_folder)
                print(f"已建立資料夾: {target_folder_name}")
            else:
                messagebox.showwarning("警告", f"未找到 '{target_folder_name}' 資料夾，程序將終止執行。")
                raise SystemExit

    # 處理所有 summary 資料夾中的 PDF 檔案
    summary_paths = []
    for summary_folder in summary_folders:
        for file in os.listdir(summary_folder):
            if file.endswith('.pdf') and not file.startswith('~$'):
                summary_path = os.path.join(summary_folder, file)
                summary_path, renamed = rename_file_if_needed(summary_path)
                if renamed:
                    # 更新快取中檔案名稱
                    relative_root = os.path.relpath(summary_folder, folder)
                    cache[relative_root]['files'][file] = os.path.getmtime(summary_path)
                    cache[relative_root]['files'][os.path.basename(summary_path)] = os.path.getmtime(summary_path)
                    del cache[relative_root]['files'][file]
                    save_cache(cache)
                summary_paths.append(summary_path)

    # 使用多進程處理 PDF 解析
    with ProcessPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(process_pdf_file, summary_path) for summary_path in summary_paths]
        for future in as_completed(futures):
            ndt_codes, welding_codes = future.result()
            ndt_codes_with_filenames_total.update(ndt_codes)
            welding_codes_total.update(welding_codes)

    return ndt_codes_with_filenames_total, welding_codes_total

def find_ndt_file(file_path, pattern, codes_with_filenames):
    """檢查檔案是否符合 NDT 編號並返回相關資訊"""
    file_name = os.path.basename(file_path)
    match = pattern.search(file_name)
    if match:
        ndt_code = match.group(1)
        if ndt_code in codes_with_filenames:
            is_cancelled = "作廢" in file_name
            return (ndt_code, file_path, is_cancelled)
    return None

def copy_ndt_file(source_file_path, target_folder_path):
    """複製 NDT 檔案到目標資料夾"""
    source_file_path, renamed = rename_file_if_needed(source_file_path)
    target_file_name = os.path.basename(source_file_path)
    target_file_path = os.path.join(target_folder_path, target_file_name)
    try:
        shutil.copy2(source_file_path, target_file_path)
        print(f"已複製 NDT 檔案: {source_file_path} -> {target_file_path}")
        return True
    except Exception as e:
        print(f"無法複製檔案 {source_file_path} 到 {target_file_path}: {e}")
        return False

def search_and_copy_ndt_pdfs(source_folder, target_folder, codes_with_filenames, is_as_built, cache, max_workers):
    """搜尋並複製 NDT PDF 檔案，避免複製「作廢」版本"""
    pattern = re.compile(r'CWP-Q-R-JK-NDT-(\d+).*?\.pdf', re.IGNORECASE)
    copied_files = 0
    not_found_codes = set(codes_with_filenames.keys())

    # 建立一個字典來存放每個 NDT 編號對應的檔案
    ndt_files = {}

    # 使用快取來搜尋檔案
    files = get_cached_files(source_folder, cache)
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(find_ndt_file, file, pattern, codes_with_filenames) for file in files]
        for future in as_completed(futures):
            result = future.result()
            if result:
                ndt_code, file_path, is_cancelled = result
                if ndt_code not in ndt_files:
                    ndt_files[ndt_code] = {'valid': None, 'cancelled': None}
                if is_cancelled:
                    ndt_files[ndt_code]['cancelled'] = file_path
                else:
                    ndt_files[ndt_code]['valid'] = file_path

    target_folder_path = os.path.join(target_folder, '06 NDT Reports' if is_as_built else '04 NDT Reports')
    os.makedirs(target_folder_path, exist_ok=True)

    # 使用多線程複製檔案
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = []
        for ndt_code, files_dict in ndt_files.items():
            if files_dict['valid']:
                futures.append(executor.submit(copy_ndt_file, files_dict['valid'], target_folder_path))
                not_found_codes.discard(ndt_code)
            elif files_dict['cancelled']:
                # 根據需求，可以選擇是否複製作廢版本
                # 這裡選擇不複製作廢版本
                print(f"找到作廢的 NDT 檔案，但不複製: {files_dict['cancelled']}")
                not_found_codes.discard(ndt_code)

        for future in as_completed(futures):
            if future.result():
                copied_files += 1

    # 構建未找到的檔案名稱列表
    not_found_filenames = {codes_with_filenames[code] for code in not_found_codes}

    return copied_files, not_found_filenames

def find_welding_file(file_path, codes):
    """檢查檔案是否符合焊材材證編號並返回相關資訊"""
    file_name = os.path.basename(file_path)
    for code in codes:
        if re.search(re.escape(code), file_name, re.IGNORECASE):
            return (code, file_path)
    return None

def copy_welding_file(file_path, target_folder_path):
    """複製焊材材證檔案到目標資料夾"""
    try:
        shutil.copy2(file_path, target_folder_path)
        print(f"已複製焊材材證檔案: {file_path} -> {target_folder_path}")
        return True
    except Exception as e:
        print(f"無法複製檔案 {file_path} 到 {target_folder_path}: {e}")
        return False

def search_and_copy_welding_pdfs(source_folder, target_folder, codes, is_as_built, cache, max_workers):
    """搜尋並複製焊材材證 PDF 檔案"""
    copied_files = 0
    not_found_codes = set(codes)
    if is_as_built:
        target_folder_path = os.path.join(target_folder, '05 Welding Consumable')
    else:
        target_folder_path = os.path.join(target_folder, '02 Material Traceability & Mill Cert', 'Welding Consumable')
    os.makedirs(target_folder_path, exist_ok=True)

    # 使用快取來搜尋檔案
    files = get_cached_files(source_folder, cache)
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(find_welding_file, file, codes) for file in files]
        welding_files = []
        for future in as_completed(futures):
            result = future.result()
            if result:
                code, file_path = result
                welding_files.append(file_path)
                not_found_codes.discard(code)

    # 使用多線程複製檔案
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        copy_futures = [executor.submit(copy_welding_file, file_path, target_folder_path) for file_path in welding_files]
        for future in as_completed(copy_futures):
            if future.result():
                copied_files += 1

    return copied_files, not_found_codes

def delete_file(file_path):
    """刪除單一檔案"""
    try:
        os.remove(file_path)
        print(f"已刪除檔案: {file_path}")
        return True
    except Exception as e:
        print(f"無法刪除檔案 {file_path}: {e}")
        return False

def delete_all_welding_pdfs(target_folder, is_as_built, max_workers):
    """刪除目標資料夾中所有的焊材材證 PDF 檔案"""
    deleted_files = 0
    if is_as_built:
        target_folder_path = os.path.join(target_folder, '05 Welding Consumable')
    else:
        target_folder_path = os.path.join(target_folder, '02 Material Traceability & Mill Cert', 'Welding Consumable')

    if os.path.exists(target_folder_path):
        files = [os.path.join(target_folder_path, file) for file in os.listdir(target_folder_path) if file.endswith('.pdf') and not file.startswith('~$')]
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(delete_file, file_path) for file_path in files]
            for future in as_completed(futures):
                if future.result():
                    deleted_files += 1

    return deleted_files

def delete_all_ndt_pdfs(target_folder, is_as_built, max_workers):
    """刪除目標資料夾中所有的 NDT PDF 檔案"""
    deleted_files = 0
    target_folder_path = os.path.join(target_folder, '06 NDT Reports' if is_as_built else '04 NDT Reports')

    if os.path.exists(target_folder_path):
        files = [os.path.join(target_folder_path, file) for file in os.listdir(target_folder_path) if file.endswith('.pdf') and not file.startswith('~$')]
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(delete_file, file_path) for file_path in files]
            for future in as_completed(futures):
                if future.result():
                    deleted_files += 1

    return deleted_files

def check_and_delete_file(file_path, regex):
    """檢查檔案是否符合命名規則，若不符合則詢問是否刪除"""
    file_name_without_extension = os.path.splitext(os.path.basename(file_path))[0]

    try:
        if not regex.search(file_name_without_extension):
            response = messagebox.askyesno(
                "確認刪除", 
                f"是否要刪除檔案：{os.path.basename(file_path)}\n"
                f"資料夾名稱：{extract_base_folder_name(os.path.basename(os.path.dirname(file_path)))}\n"
                "此檔案似乎不符合命名規則。"
            )
            if response:
                os.remove(file_path)
                print(f"已刪除檔案: {file_path}")
                return file_path
            else:
                print(f"使用者選擇保留檔案: {file_path}")
    except Exception as e:
        print(f"處理檔案時發生錯誤 {file_path}: {str(e)}")
    return None

def clean_unmatched_files(pdf_folder, is_as_built, cache, max_workers):
    """
    清理不符合命名規則的檔案，考慮檔名中的 .001 格式
    """
    deleted_files = []

    try:
        folder_name = os.path.basename(pdf_folder)
        base_folder_name = extract_base_folder_name(folder_name)

        if not base_folder_name:
            print(f"警告：無法從資料夾名稱中提取基本名稱: {folder_name}")
            return deleted_files

        # 建立更精確的檔案名稱匹配模式
        base_name_without_number = re.sub(r'#\d+$', '', base_folder_name)
        base_patterns = [
            re.escape(base_folder_name),  # 完全匹配
            r'CWP\d+[A-Z]-' + re.escape(base_folder_name),  # 允許 CWP 前綴
            # 更精確的匹配模式，考慮檔名中的 .001
            (r'.*?' + 
             re.escape(base_name_without_number.split('.')[0]) +  # 取基本名稱（不含流水號）
             r'(?:\.\d{3})?' +  # 可選的第一個流水號（資料夾名稱中的）
             r'#\d{3}' + 
             r'(?:\.\d{3})?' +  # 可選的第二個流水號（檔名中的）
             r'.*')
        ]

        combined_pattern = '|'.join(f'({pattern})' for pattern in base_patterns)
        combined_regex = re.compile(combined_pattern, re.IGNORECASE)

        if is_as_built:
            target_folders = ["04 Welding Identification Summary", "03 Material Traceability & Mill Cert"]
        else:
            target_folders = ["01 Welding Identification Summary", "02 Material Traceability & Mill Cert"]

        print(f"資料夾基本名稱: {base_folder_name}")
        print(f"使用的匹配模式: {combined_pattern}")

        # 使用多線程檢查並刪除不符合的檔案
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = []
            for subfolder in target_folders:
                subfolder_path = os.path.join(pdf_folder, subfolder)
                if os.path.exists(subfolder_path):
                    files = [os.path.join(subfolder_path, file) for file in os.listdir(subfolder_path) if file.endswith('.pdf') and not file.startswith('~$')]
                    futures = [executor.submit(check_and_delete_file, file_path, combined_regex) for file_path in files]
                    break  # 只處理第一層子資料夾

            for future in as_completed(futures):
                deleted_file = future.result()
                if deleted_file:
                    deleted_files.append(deleted_file)

    except Exception as e:
        print(f"處理資料夾時發生錯誤 {pdf_folder}: {str(e)}")
        messagebox.showerror("錯誤", f"處理資料夾時發生錯誤：\n{str(e)}")

    return deleted_files

def process_single_folder(pdf_folder, ndt_source_pdf_folder, welding_source_pdf_folder, is_as_built, cache, max_workers):
    """處理單一資料夾中的所有操作"""
    # 先處理 NDT 報驗單和焊材材證編號
    ndt_codes_with_filenames_total, welding_codes_total = process_pdf_files_in_folder(pdf_folder, is_as_built, cache, max_workers)

    reasons = []

    # 檢查是否找到 NDT 編號
    if not ndt_codes_with_filenames_total:
        reasons.append("沒有在 PDF 中找到任何符合條件的 NDT 編號。")

    # 檢查是否找到焊材材證編號
    if not welding_codes_total:
        reasons.append("沒有在 PDF 中找到任何符合條件的焊材材證編號。")

    # 刪除所有現有的 NDT 報驗單檔案
    delete_all_ndt_pdfs(pdf_folder, is_as_built, max_workers)

    # 在複製焊材材證之前，先刪除所有現有的焊材材證檔案
    delete_all_welding_pdfs(pdf_folder, is_as_built, max_workers)

    # 開始複製 NDT 報驗單
    ndt_copied, not_found_ndt_filenames = search_and_copy_ndt_pdfs(
        ndt_source_pdf_folder, pdf_folder, ndt_codes_with_filenames_total, is_as_built, cache, max_workers
    )

    if ndt_copied == 0 and ndt_codes_with_filenames_total:
        reasons.append("找到了 NDT 編號，但沒有找到對應的報驗單 PDF 檔案。")

    # 開始複製焊材材證
    welding_copied, not_found_welding_codes = search_and_copy_welding_pdfs(
        welding_source_pdf_folder, pdf_folder, welding_codes_total, is_as_built, cache, max_workers
    )

    if welding_copied == 0 and welding_codes_total:
        reasons.append("找到了焊材材證編號，但沒有找到對應的焊材材證 PDF 檔案。")

    # 清理不符合的檔案
    deleted_files = clean_unmatched_files(pdf_folder, is_as_built, cache, max_workers)

    return ndt_copied, welding_copied, not_found_ndt_filenames, not_found_welding_codes, reasons, deleted_files

def process_folders(pdf_folder, ndt_source_pdf_folder, welding_source_pdf_folder):
    """處理所有目標資料夾"""
    total_ndt_copied = 0
    total_welding_copied = 0
    not_found_ndt_filenames_total = set()
    not_found_welding_codes_total = set()
    reasons_total = []
    deleted_files_total = []

    is_as_built = "As-Built" in pdf_folder

    # 載入快取
    cache = load_cache()

    # 設置最大工作線程數
    cpu_count = os.cpu_count() or 12  # 預設為 12，符合 i5-12500 的超線程
    max_workers = cpu_count * 2

    # 先檢查並重新命名檔案
    renamed_files = check_and_rename_files_in_folder_cached(pdf_folder, cache, max_workers)
    if renamed_files:
        message = "以下檔案已被重新命名：\n"
        for old_path, new_path in renamed_files:
            message += f"舊名稱：{os.path.basename(old_path)} -> 新名稱：{os.path.basename(new_path)}\n"
        messagebox.showinfo("檔案重命名", message)
        # 更新快取
        save_cache(cache)

    if is_target_folder(os.path.basename(pdf_folder)):
        # 如果根目錄就是目標資料夾
        ndt_copied, welding_copied, not_found_ndt_filenames, not_found_welding_codes, reasons, deleted_files = process_single_folder(
            pdf_folder, ndt_source_pdf_folder, welding_source_pdf_folder, is_as_built, cache, max_workers
        )

        total_ndt_copied += ndt_copied
        total_welding_copied += welding_copied
        not_found_ndt_filenames_total.update(not_found_ndt_filenames)
        not_found_welding_codes_total.update(not_found_welding_codes)
        reasons_total.extend(reasons)
        deleted_files_total.extend(deleted_files)
    else:
        # 如果根目錄不是目標資料夾，遍歷子資料夾尋找目標資料夾
        for root, dirs, files in os.walk(pdf_folder):
            for dir in dirs:
                if is_target_folder(dir):
                    subfolder_path = os.path.join(root, dir)
                    ndt_copied, welding_copied, not_found_ndt_filenames, not_found_welding_codes, reasons, deleted_files = process_single_folder(
                        subfolder_path, ndt_source_pdf_folder, welding_source_pdf_folder, is_as_built, cache, max_workers
                    )

                    total_ndt_copied += ndt_copied
                    total_welding_copied += welding_copied
                    not_found_ndt_filenames_total.update(not_found_ndt_filenames)
                    not_found_welding_codes_total.update(not_found_welding_codes)
                    reasons_total.extend(reasons)
                    deleted_files_total.extend(deleted_files)
            break  # 只處理第一層子資料夾

    # 儲存快取
    save_cache(cache)

    return total_ndt_copied, total_welding_copied, not_found_ndt_filenames_total, not_found_welding_codes_total, reasons_total, deleted_files_total

def main():
    """主函式"""
    root = tk.Tk()
    root.withdraw()

    # 選擇包含銲道追溯檔案的資料夾
    pdf_initialdir = "C:/Users/CWP-PC-E-COM302/Box/T460 風電 品管 簡瑞成/FAT package"
    if not os.path.exists(pdf_initialdir):
        pdf_initialdir = os.path.expanduser("~")

    pdf_folder = filedialog.askdirectory(
        initialdir=pdf_initialdir,
        title="請選擇包含銲道追溯檔案的資料夾:"
    )

    if not pdf_folder:
        messagebox.showwarning("警告", "未選擇任何資料夾，程序將終止。")
        return

    # 選擇包含報驗單 PDF 檔案的資料夾
    ndt_source_pdf_initialdir = "U:/N-品管部/@品管部共用資料區/專案/CWP06 台電二期專案/報驗單/001_JK報驗單/完成 (pdf)"
    if not os.path.exists(ndt_source_pdf_initialdir):
        ndt_source_pdf_initialdir = os.path.expanduser("~")

    ndt_source_pdf_folder = filedialog.askdirectory(
        initialdir=ndt_source_pdf_initialdir,
        title="請選擇包含報驗單 PDF 檔案的資料夾:"
    )

    if not ndt_source_pdf_folder:
        messagebox.showwarning("警告", "未選擇任何報驗單資料夾，程序將終止。")
        return

    # 選擇包含焊材材證 PDF 檔案的資料夾
    welding_source_pdf_initialdir = "U:/N-品管部/@品管部共用資料區/品管人員資料夾/T460 風電 品管 簡瑞成/焊材材證"
    if not os.path.exists(welding_source_pdf_initialdir):
        welding_source_pdf_initialdir = os.path.expanduser("~")

    welding_source_pdf_folder = filedialog.askdirectory(
        initialdir=welding_source_pdf_initialdir,
        title="請選擇包含焊材材證 PDF 檔案的資料夾:"
    )

    if not welding_source_pdf_folder:
        messagebox.showwarning("警告", "未選擇任何焊材材證資料夾，程序將終止。")
        return

    is_as_built = "As-Built" in pdf_folder
    mode = "竣工模式" if is_as_built else "一般模式"

    # 處理資料夾
    total_ndt_copied, total_welding_copied, not_found_ndt_filenames_total, not_found_welding_codes_total, reasons_total, deleted_files_total = process_folders(
        pdf_folder, ndt_source_pdf_folder, welding_source_pdf_folder
    )

    # 構建訊息內容
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

    # 顯示完成訊息
    messagebox.showinfo("完成", message)

if __name__ == "__main__":
    main()
