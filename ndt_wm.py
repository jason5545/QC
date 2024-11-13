import os
import shutil
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox
import re
from difflib import get_close_matches
import json
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading

class FileCache:
    def __init__(self, cache_file='file_cache.json'):
        self.cache_file = cache_file
        self.cache = {}
        self.load_cache()
    
    def load_cache(self):
        if os.path.exists(self.cache_file):
            try:
                with open(self.cache_file, 'r', encoding='utf-8') as f:
                    self.cache = json.load(f)
            except Exception as e:
                print(f"無法載入快取檔案 {self.cache_file}: {e}")
                self.cache = {}
    
    def save_cache(self):
        try:
            with open(self.cache_file, 'w', encoding='utf-8') as f:
                json.dump(self.cache, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"無法儲存快取檔案 {self.cache_file}: {e}")
    
    def get_file_data(self, file_path):
        if file_path in self.cache:
            cached = self.cache[file_path]
            try:
                current_mtime = os.path.getmtime(file_path)
                if cached['mtime'] == current_mtime:
                    return cached['data']
            except Exception as e:
                print(f"無法獲取檔案修改時間 {file_path}: {e}")
        return None
    
    def update_file_data(self, file_path, data):
        try:
            mtime = os.path.getmtime(file_path)
            self.cache[file_path] = {'mtime': mtime, 'data': data}
        except Exception as e:
            print(f"無法更新快取檔案 {file_path}: {e}")

def rename_file_if_needed(file_path, cache):
    """檢查檔案名稱中是否包含 CWP06G-XB4C 並取代"""
    # 檢查快取
    cached_data = cache.get_file_data(file_path)
    if cached_data and 'renamed_path' in cached_data:
        new_file_path = cached_data['renamed_path']
        if os.path.exists(new_file_path):
            return new_file_path
    file_name = os.path.basename(file_path)
    if "CWP06G-XB4C" in file_name:
        new_file_name = file_name.replace("CWP06G-XB4C", "CWP06C-XB4C")
        new_file_path = os.path.join(os.path.dirname(file_path), new_file_name)
        try:
            os.rename(file_path, new_file_path)
            print(f"檔案已重新命名: {file_path} -> {new_file_path}")
            cache.update_file_data(file_path, {'renamed_path': new_file_path})
            return new_file_path
        except Exception as e:
            print(f"無法重新命名檔案 {file_path}: {e}")
    return file_path

def check_and_rename_files_in_folder(folder_path, cache):
    """遍歷資料夾中的所有檔案，並進行必要的重新命名"""
    renamed_files = []
    with ThreadPoolExecutor(max_workers=12) as executor:
        future_to_file = {executor.submit(rename_file_if_needed, os.path.join(root, file), cache): os.path.join(root, file)
                          for root, dirs, files in os.walk(folder_path) for file in files}
        for future in as_completed(future_to_file):
            file_path = future_to_file[future]
            try:
                new_file_path = future.result()
                if new_file_path != file_path:
                    renamed_files.append((file_path, new_file_path))
            except Exception as e:
                print(f"錯誤處理檔案 {file_path}: {e}")
    return renamed_files

def get_ndt_codes_from_pdf(file_path, cache):
    """從 PDF 中提取 NDT 編號"""
    cached_data = cache.get_file_data(file_path)
    if cached_data and 'ndt_codes' in cached_data:
        return {code: f'CWP-Q-R-JK-NDT-{code}.pdf' for code in cached_data['ndt_codes']}
    
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
            codes_with_filenames = {code: f'CWP-Q-R-JK-NDT-{code}.pdf' for code in matches}
            print(f"從 {file_path} 提取到 NDT 編號: {matches}")
        else:
            print(f"從 {file_path} 未提取到任何 NDT 編號。")

        # 更新快取
        cache.update_file_data(file_path, {'ndt_codes': matches})

    except Exception as e:
        print(f"無法讀取 PDF 檔案 {file_path}: {e}")

    return codes_with_filenames

def get_welding_codes_from_pdf(file_path, cache):
    """從 PDF 中提取焊材材證編號"""
    cached_data = cache.get_file_data(file_path)
    if cached_data and 'welding_codes' in cached_data:
        return set(cached_data['welding_codes'])
    
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
            codes = set(matches)
            print(f"從 {file_path} 提取到焊材材證編號: {matches}")
        else:
            print(f"從 {file_path} 未提取到任何焊材材證編號。")

        # 更新快取
        cache.update_file_data(file_path, {'welding_codes': list(codes)})

    except Exception as e:
        print(f"無法讀取 PDF 檔案 {file_path}: {e}")

    return codes

def is_target_folder(folder_name):
    """判斷是否為目標資料夾，並支援如 '6S201#021' 的格式"""
    # 修改正則表達式以捕捉更多數字位數
    pattern = re.compile(r'XB1#\d+|XB[1-4][ABC]#\d+|6S21[1-7]#\d+|6S20[1256]#\d+', re.IGNORECASE)
    return pattern.match(folder_name) is not None

def extract_base_folder_name(folder_name):
    """從資料夾名稱中提取基本名稱，支援如 '6S201#021' 的格式"""
    # 修改正則表達式以捕捉更多數字位數
    pattern = re.compile(r'(XB1#\d+|XB[1-4][ABC]#\d+|6S21[1-7]#\d+|6S20[1256]#\d+)', re.IGNORECASE)
    match = pattern.search(folder_name)
    if match:
        return match.group(1)
    else:
        return folder_name

def process_pdf_files_in_folder(folder, is_as_built, cache):
    """處理指定資料夾中的 PDF 檔案，提取 NDT 和焊材材證編號"""
    ndt_codes_with_filenames_total = {}
    welding_codes_total = set()
    target_folder_name = "04 Welding Identification Summary" if is_as_built else "01 Welding Identification Summary"

    # 找到目標資料夾，可能需要重新命名
    for root, dirs, files in os.walk(folder):
        if target_folder_name not in dirs:
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
            else:
                if is_as_built:
                    os.mkdir(os.path.join(root, target_folder_name))
                    dirs.append(target_folder_name)
                    print(f"已建立資料夾: {target_folder_name}")
                else:
                    response = messagebox.askyesno(
                        "資料夾名稱不符",
                        f"未找到 '{target_folder_name}' 資料夾。\n是否要建立該資料夾？"
                    )
                    if response:
                        os.mkdir(os.path.join(root, target_folder_name))
                        dirs.append(target_folder_name)
                        print(f"已建立資料夾: {target_folder_name}")
                    else:
                        messagebox.showwarning("警告", f"未找到 '{target_folder_name}' 資料夾，程序將終止執行。")
                        raise SystemExit

        summary_folder = os.path.join(root, target_folder_name)
        print(f"處理資料夾: {summary_folder}")

        # 獲取所有 PDF 檔案
        pdf_files = [os.path.join(summary_folder, f) for f in os.listdir(summary_folder)
                     if f.endswith('.pdf') and not f.startswith('~$')]

        with ThreadPoolExecutor(max_workers=12) as executor:
            # 重新命名檔案
            rename_futures = {executor.submit(rename_file_if_needed, file_path, cache): file_path for file_path in pdf_files}
            renamed_files = []
            for future in as_completed(rename_futures):
                file_path = rename_futures[future]
                try:
                    new_file_path = future.result()
                    if new_file_path != file_path:
                        renamed_files.append((file_path, new_file_path))
                except Exception as e:
                    print(f"錯誤重新命名檔案 {file_path}: {e}")

            # 更新 pdf_files 列表
            pdf_files = [new_path if new_path != old_path else old_path for old_path, new_path in renamed_files]

            # 提取 NDT 編號
            ndt_futures = {executor.submit(get_ndt_codes_from_pdf, file_path, cache): file_path for file_path in pdf_files}
            for future in as_completed(ndt_futures):
                file_path = ndt_futures[future]
                try:
                    ndt_codes = future.result()
                    ndt_codes_with_filenames_total.update(ndt_codes)
                except Exception as e:
                    print(f"錯誤提取 NDT 編號從檔案 {file_path}: {e}")

            # 提取焊材材證編號
            welding_futures = {executor.submit(get_welding_codes_from_pdf, file_path, cache): file_path for file_path in pdf_files}
            for future in as_completed(welding_futures):
                file_path = welding_futures[future]
                try:
                    welding_codes = future.result()
                    welding_codes_total.update(welding_codes)
                except Exception as e:
                    print(f"錯誤提取焊材材證編號從檔案 {file_path}: {e}")

        break  # 只處理第一層子資料夾

    return ndt_codes_with_filenames_total, welding_codes_total

def copy_file(source, target, code, not_found_codes, lock):
    try:
        shutil.copy2(source, target)
        print(f"已複製 NDT 檔案: {source} -> {target}")
        with lock:
            not_found_codes.discard(code)
        return True
    except Exception as e:
        print(f"無法複製檔案 {source} 到 {target}: {e}")
        return False

def search_and_copy_ndt_pdfs(source_folder, target_folder, codes_with_filenames, is_as_built, cache):
    """搜尋並複製 NDT PDF 檔案，避免複製「作廢」版本"""
    pattern = re.compile(r'CWP-Q-R-JK-NDT-(\d+).*?\.pdf', re.IGNORECASE)
    copied_files = 0
    not_found_codes = set(codes_with_filenames.keys())

    # 建立一個字典來存放每個 NDT 編號對應的檔案
    ndt_files = {}
    ndt_lock = threading.Lock()

    # 索引檔案
    def index_file(root, file):
        if file.endswith('.pdf') and not file.startswith('~$'):
            match = pattern.search(file)
            if match:
                ndt_code = match.group(1)
                file_path = os.path.join(root, file)
                # 檢查快取並重新命名
                file_path = rename_file_if_needed(file_path, cache)
                # 判斷是否為「作廢」版本
                is_cancelled = "作廢" in file
                with ndt_lock:
                    if ndt_code not in ndt_files:
                        ndt_files[ndt_code] = {'valid': None, 'cancelled': None}
                    if is_cancelled:
                        ndt_files[ndt_code]['cancelled'] = file_path
                    else:
                        ndt_files[ndt_code]['valid'] = file_path

    with ThreadPoolExecutor(max_workers=12) as executor:
        futures = []
        for root, dirs, files in os.walk(source_folder):
            for file in files:
                futures.append(executor.submit(index_file, root, file))
            break  # 只處理第一層子資料夾
        for future in as_completed(futures):
            try:
                future.result()
            except Exception as e:
                print(f"錯誤索引檔案: {e}")

    target_folder_path = os.path.join(target_folder, '06 NDT Reports' if is_as_built else '04 NDT Reports')
    os.makedirs(target_folder_path, exist_ok=True)

    # 複製檔案
    with ThreadPoolExecutor(max_workers=12) as executor:
        futures = []
        for ndt_code, files_dict in ndt_files.items():
            if ndt_code in not_found_codes:
                if files_dict['valid']:
                    source_file_path = files_dict['valid']
                    target_file_name = os.path.basename(source_file_path)
                    target_file_path = os.path.join(target_folder_path, target_file_name)
                    futures.append(executor.submit(copy_file, source_file_path, target_file_path, ndt_code, not_found_codes, ndt_lock))
                elif files_dict['cancelled']:
                    # 根據需求，這裡選擇不複製作廢版本
                    print(f"找到作廢的 NDT 檔案，但不複製: {files_dict['cancelled']}")
                    with ndt_lock:
                        not_found_codes.discard(ndt_code)
        for future in as_completed(futures):
            try:
                if future.result():
                    copied_files += 1
            except Exception as e:
                print(f"錯誤複製檔案: {e}")

    # 構建未找到的檔案名稱列表
    not_found_filenames = {codes_with_filenames[code] for code in not_found_codes}

    return copied_files, not_found_filenames

def search_and_copy_welding_pdfs(source_folder, target_folder, codes, is_as_built, cache):
    """搜尋並複製焊材材證 PDF 檔案"""
    copied_files = 0
    not_found_codes = set(codes)
    if is_as_built:
        target_folder_path = os.path.join(target_folder, '05 Welding Consumable')
    else:
        target_folder_path = os.path.join(target_folder, '02 Material Traceability & Mill Cert', 'Welding Consumable')
    os.makedirs(target_folder_path, exist_ok=True)

    # 預先編譯所有需要的正則表達式
    patterns = {code: re.compile(re.escape(code), re.IGNORECASE) for code in codes}

    welding_lock = threading.Lock()

    # 定義複製任務
    def copy_welding_file(file, root):
        if file.endswith('.pdf') and not file.startswith('~$'):
            source_file_path = os.path.join(root, file)
            source_file_path = rename_file_if_needed(source_file_path, cache)
            for code, pattern in patterns.items():
                if pattern.search(file):
                    with welding_lock:
                        if code in not_found_codes:
                            not_found_codes.discard(code)
                            target_file_path = os.path.join(target_folder_path, file)
                            try:
                                shutil.copy2(source_file_path, target_file_path)
                                print(f"已複製焊材材證檔案: {source_file_path} -> {target_file_path}")
                                return 1
                            except Exception as e:
                                print(f"無法複製檔案 {source_file_path} 到 {target_file_path}: {e}")
                    break  # 一個檔案只會對應到一個 code
        return 0

    with ThreadPoolExecutor(max_workers=12) as executor:
        futures = []
        for root, dirs, files in os.walk(source_folder):
            for file in files:
                futures.append(executor.submit(copy_welding_file, file, root))
        for future in as_completed(futures):
            try:
                copied_files += future.result()
            except Exception as e:
                print(f"錯誤複製焊材材證檔案: {e}")

    return copied_files, not_found_codes

def delete_all_welding_pdfs(target_folder, is_as_built):
    """刪除目標資料夾中所有的焊材材證 PDF 檔案"""
    deleted_files = 0
    if is_as_built:
        target_folder_path = os.path.join(target_folder, '05 Welding Consumable')
    else:
        target_folder_path = os.path.join(target_folder, '02 Material Traceability & Mill Cert', 'Welding Consumable')

    if os.path.exists(target_folder_path):
        for file in os.listdir(target_folder_path):
            if file.endswith('.pdf') and not file.startswith('~$'):
                file_path = os.path.join(target_folder_path, file)
                try:
                    os.remove(file_path)
                    deleted_files += 1
                    print(f"已刪除焊材材證檔案: {file_path}")
                except Exception as e:
                    print(f"無法刪除檔案 {file_path}: {e}")

    return deleted_files

def delete_all_ndt_pdfs(target_folder, is_as_built):
    """刪除目標資料夾中所有的 NDT PDF 檔案"""
    deleted_files = 0
    target_folder_path = os.path.join(target_folder, '06 NDT Reports' if is_as_built else '04 NDT Reports')

    if os.path.exists(target_folder_path):
        for file in os.listdir(target_folder_path):
            if file.endswith('.pdf') and not file.startswith('~$'):
                file_path = os.path.join(target_folder_path, file)
                try:
                    os.remove(file_path)
                    deleted_files += 1
                    print(f"已刪除 NDT 檔案: {file_path}")
                except Exception as e:
                    print(f"無法刪除檔案 {file_path}: {e}")

    return deleted_files

def clean_unmatched_files(pdf_folder, is_as_built):
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
                                if not combined_regex.search(file_name_without_extension):
                                    response = messagebox.askyesno(
                                        "確認刪除", 
                                        f"是否要刪除檔案：{file}\n"
                                        f"資料夾名稱：{base_folder_name}\n"
                                        "此檔案似乎不符合命名規則。"
                                    )
                                    if response:
                                        os.remove(file_path)
                                        deleted_files.append(file_path)
                                        print(f"已刪除檔案: {file_path}")
                                    else:
                                        print(f"使用者選擇保留檔案: {file_path}")
                            except Exception as e:
                                print(f"處理檔案時發生錯誤 {file_path}: {str(e)}")
                    break  # 只處理第一層子資料夾
    except Exception as e:
        print(f"處理資料夾時發生錯誤 {pdf_folder}: {str(e)}")
        messagebox.showerror("錯誤", f"處理資料夾時發生錯誤：\n{str(e)}")

    return deleted_files

def process_single_folder(pdf_folder, ndt_source_pdf_folder, welding_source_pdf_folder, is_as_built, cache):
    """處理單一資料夾中的所有操作"""
    # 先處理 NDT 報驗單和焊材材證編號
    ndt_codes_with_filenames_total, welding_codes_total = process_pdf_files_in_folder(pdf_folder, is_as_built, cache)

    reasons = []

    # 檢查是否找到 NDT 編號
    if not ndt_codes_with_filenames_total:
        reasons.append("沒有在 PDF 中找到任何符合條件的 NDT 編號。")

    # 檢查是否找到焊材材證編號
    if not welding_codes_total:
        reasons.append("沒有在 PDF 中找到任何符合條件的焊材材證編號。")

    # 刪除所有現有的 NDT 報驗單檔案
    delete_all_ndt_pdfs(pdf_folder, is_as_built)

    # 在複製焊材材證之前，先刪除所有現有的焊材材證檔案
    delete_all_welding_pdfs(pdf_folder, is_as_built)

    # 開始複製 NDT 報驗單
    ndt_copied, not_found_ndt_filenames = search_and_copy_ndt_pdfs(
        ndt_source_pdf_folder, pdf_folder, ndt_codes_with_filenames_total, is_as_built, cache
    )

    if ndt_copied == 0 and ndt_codes_with_filenames_total:
        reasons.append("找到了 NDT 編號，但沒有找到對應的報驗單 PDF 檔案。")

    # 開始複製焊材材證
    welding_copied, not_found_welding_codes = search_and_copy_welding_pdfs(
        welding_source_pdf_folder, pdf_folder, welding_codes_total, is_as_built, cache
    )

    if welding_copied == 0 and welding_codes_total:
        reasons.append("找到了焊材材證編號，但沒有找到對應的焊材材證 PDF 檔案。")

    # 清理不符合的檔案
    deleted_files = clean_unmatched_files(pdf_folder, is_as_built)

    return ndt_copied, welding_copied, not_found_ndt_filenames, not_found_welding_codes, reasons, deleted_files

def process_folders(pdf_folder, ndt_source_pdf_folder, welding_source_pdf_folder, cache):
    """處理所有目標資料夾"""
    total_ndt_copied = 0
    total_welding_copied = 0
    not_found_ndt_filenames_total = set()
    not_found_welding_codes_total = set()
    reasons_total = []
    deleted_files_total = []

    is_as_built = "As-Built" in pdf_folder

    # 先檢查並重新命名檔案
    renamed_files = check_and_rename_files_in_folder(pdf_folder, cache)
    if renamed_files:
        message = "以下檔案已被重新命名：\n"
        for old_path, new_path in renamed_files:
            message += f"舊名稱：{os.path.basename(old_path)} -> 新名稱：{os.path.basename(new_path)}\n"
        messagebox.showinfo("檔案重命名", message)

    if is_target_folder(os.path.basename(pdf_folder)):
        # 如果根目錄就是目標資料夾
        ndt_copied, welding_copied, not_found_ndt_filenames, not_found_welding_codes, reasons, deleted_files = process_single_folder(
            pdf_folder, ndt_source_pdf_folder, welding_source_pdf_folder, is_as_built, cache
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
                        subfolder_path, ndt_source_pdf_folder, welding_source_pdf_folder, is_as_built, cache
                    )

                    total_ndt_copied += ndt_copied
                    total_welding_copied += welding_copied
                    not_found_ndt_filenames_total.update(not_found_ndt_filenames)
                    not_found_welding_codes_total.update(not_found_welding_codes)
                    reasons_total.extend(reasons)
                    deleted_files_total.extend(deleted_files)
            break  # 只處理第一層子資料夾

    return total_ndt_copied, total_welding_copied, not_found_ndt_filenames_total, not_found_welding_codes_total, reasons_total, deleted_files_total

def main():
    """主函式"""
    root = tk.Tk()
    root.withdraw()

    # 初始化快取
    cache = FileCache()

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
        pdf_folder, ndt_source_pdf_folder, welding_source_pdf_folder, cache
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

    # 儲存快取
    cache.save_cache()

if __name__ == "__main__":
    main()
