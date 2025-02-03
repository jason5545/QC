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
import logging  # 導入 logging 模組

# 設定日誌記錄
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 定義常數
CACHE_FILE = 'file_cache.json'
SUMMARY_FOLDER_NAME_AS_BUILT = "04 Welding Identification Summary"
SUMMARY_FOLDER_NAME_GENERAL = "01 Welding Identification Summary"
NDT_REPORTS_FOLDER_AS_BUILT = '06 NDT Reports'
NDT_REPORTS_FOLDER_GENERAL = '04 NDT Reports'
WELDING_CONSUMABLE_FOLDER_AS_BUILT = '05 Welding Consumable'
WELDING_CONSUMABLE_FOLDER_GENERAL = os.path.join('02 Material Traceability & Mill Cert', 'Welding Consumable')
# 修改部分：竣工模式下的物料追溯應為 "02 Material Traceability"
MATERIAL_TRACEABILITY_FOLDER_AS_BUILT = "02 Material Traceability"
MATERIAL_TRACEABILITY_FOLDER_GENERAL = "02 Material Traceability & Mill Cert"

# 編譯常用的正則表達式，避免重複編譯
RE_OLD_FILENAME_PATTERN = re.compile(r"CWP06G-XB4C")  # 更具體的命名
RE_NDT_CODE = re.compile(r'CWPQRJKNDT(\d+)', re.IGNORECASE)
RE_WELDING_CODE = re.compile(r'\b\d{6,10}\b')
RE_TARGET_FOLDER = re.compile(r'XB1#\d+|XB[1-4][ABC]#\d+|6S21[1-7]#\d+|6S20[12356]#\d+', re.IGNORECASE)
RE_BASE_FOLDER_NAME = re.compile(r'(XB1#\d+|XB[1-4][ABC]#\d+|6S21[1-7]#\d+|6S20[12356]#\d+)', re.IGNORECASE)

class FileCache:
    """檔案快取類別，用於儲存檔案的修改時間和資料。"""
    def __init__(self, cache_file=CACHE_FILE):
        self.cache_file = cache_file
        self.cache = {}
        self.lock = threading.Lock()
        self._load_cache()

    def _load_cache(self):
        """從檔案載入快取資料。"""
        if os.path.exists(self.cache_file):
            try:
                with open(self.cache_file, 'r', encoding='utf-8') as f:
                    self.cache = json.load(f)
                logging.info(f"成功載入快取檔案: {self.cache_file}")
            except Exception as e:
                logging.error(f"無法載入快取檔案 {self.cache_file}: {e}")
                self.cache = {}

    def save_cache(self):
        """將快取資料儲存到檔案。"""
        with self.lock:
            try:
                with open(self.cache_file, 'w', encoding='utf-8') as f:
                    json.dump(self.cache, f, ensure_ascii=False, indent=4)
                logging.info(f"成功儲存快取檔案: {self.cache_file}")
            except Exception as e:
                logging.error(f"無法儲存快取檔案 {self.cache_file}: {e}")

    def get_file_data(self, file_path):
        """
        取得檔案的快取資料，若檔案未修改則返回快取，否則返回 None。
        Args:
            file_path (str): 檔案路徑。
        Returns:
            dict or None: 快取資料，如果檔案已修改則返回 None。
        """
        with self.lock:
            cached = self.cache.get(file_path)
            if cached:
                try:
                    current_mtime = os.path.getmtime(file_path)
                    if cached['mtime'] == current_mtime:
                        return cached['data']
                except FileNotFoundError:
                    logging.warning(f"快取中檔案不存在: {file_path}")
                    return None
                except Exception as e:
                    logging.error(f"無法取得檔案修改時間 {file_path}: {e}")
            return None

    def update_file_data(self, file_path, data):
        """
        更新檔案的快取資料。
        Args:
            file_path (str): 檔案路徑。
            data (dict): 要儲存的資料。
        """
        with self.lock:
            try:
                mtime = os.path.getmtime(file_path)
                self.cache[file_path] = {'mtime': mtime, 'data': data}
            except Exception as e:
                logging.error(f"無法更新快取檔案 {file_path}: {e}")

# 檔案操作相關函數
def rename_file_if_needed(file_path, cache):
    """檢查檔案名稱中是否包含 CWP06G-XB4C 並取代，使用快取。"""
    cached_data = cache.get_file_data(file_path)
    if cached_data and 'renamed_path' in cached_data and os.path.exists(cached_data['renamed_path']):
        return cached_data['renamed_path']

    file_name = os.path.basename(file_path)
    if RE_OLD_FILENAME_PATTERN.search(file_name):
        new_file_name = RE_OLD_FILENAME_PATTERN.sub("CWP06C-XB4C", file_name)
        new_file_path = os.path.join(os.path.dirname(file_path), new_file_name)
        try:
            os.rename(file_path, new_file_path)
            logging.info(f"檔案已重新命名: {file_path} -> {new_file_path}")
            cache.update_file_data(file_path, {'renamed_path': new_file_path})
            return new_file_path
        except Exception as e:
            logging.error(f"無法重新命名檔案 {file_path}: {e}")
    return file_path

def check_and_rename_files_in_folder(folder_path, cache):
    """遍歷資料夾中的所有檔案，並使用多執行緒進行必要的重新命名。"""
    renamed_files = []
    with ThreadPoolExecutor(max_workers=os.cpu_count()) as executor:
        future_to_file = {executor.submit(rename_file_if_needed, os.path.join(root, file), cache): os.path.join(root, file)
                          for root, dirs, files in os.walk(folder_path) for file in files}
        for future in as_completed(future_to_file):
            file_path = future_to_file[future]
            try:
                new_file_path = future.result()
                if new_file_path != file_path:
                    renamed_files.append((file_path, new_file_path))
            except Exception as e:
                logging.error(f"錯誤處理檔案 {file_path}: {e}")
    return renamed_files

# PDF 內容提取相關函數
def get_ndt_codes_from_pdf(file_path, cache):
    """從 PDF 中提取 NDT 編號，使用快取。"""
    cached_data = cache.get_file_data(file_path)
    if cached_data and 'ndt_codes' in cached_data:
        return {code: f'CWP-Q-R-JK-NDT-{code}.pdf' for code in cached_data['ndt_codes']}

    codes_with_filenames = {}
    try:
        doc = fitz.open(file_path)
        text = "".join(page.get_text() for page in doc)
        doc.close()

        matches = RE_NDT_CODE.findall(text)
        if matches:
            codes_with_filenames = {code: f'CWP-Q-R-JK-NDT-{code}.pdf' for code in matches}
            logging.info(f"從 {file_path} 提取到 NDT 編號: {matches}")
        else:
            logging.info(f"從 {file_path} 未提取到任何 NDT 編號。")
        cache.update_file_data(file_path, {'ndt_codes': matches})
    except Exception as e:
        logging.error(f"無法讀取 PDF 檔案 {file_path}: {e}")
    return codes_with_filenames

def get_welding_codes_from_pdf(file_path, cache):
    """從 PDF 中提取焊材材證編號，使用快取。"""
    cached_data = cache.get_file_data(file_path)
    if cached_data and 'welding_codes' in cached_data:
        return set(cached_data['welding_codes'])

    codes = set()
    try:
        doc = fitz.open(file_path)
        text = "".join(page.get_text() for page in doc)
        doc.close()

        matches = RE_WELDING_CODE.findall(text)
        if matches:
            codes = set(matches)
            logging.info(f"從 {file_path} 提取到焊材材證編號: {matches}")
        else:
            logging.info(f"從 {file_path} 未提取到任何焊材材證編號。")
        cache.update_file_data(file_path, {'welding_codes': list(codes)})
    except Exception as e:
        logging.error(f"無法讀取 PDF 檔案 {file_path}: {e}")
    return codes

# 資料夾判斷與處理函數
def is_target_folder(folder_name):
    """判斷是否為目標資料夾。"""
    return RE_TARGET_FOLDER.match(folder_name) is not None

def extract_base_folder_name(folder_name):
    """從資料夾名稱中提取基本名稱。"""
    match = RE_BASE_FOLDER_NAME.search(folder_name)
    return match.group(1) if match else folder_name

def process_pdf_files_in_folder(folder, is_as_built, cache):
    """處理指定資料夾中的 PDF 檔案，提取 NDT 和焊材材證編號。"""
    ndt_codes_with_filenames_total = {}
    welding_codes_total = set()
    target_folder_name = SUMMARY_FOLDER_NAME_AS_BUILT if is_as_built else SUMMARY_FOLDER_NAME_GENERAL

    for root, dirs, files in os.walk(folder):
        if target_folder_name not in dirs:
            close_matches = get_close_matches(target_folder_name, dirs, n=1, cutoff=0.6)
            if close_matches:
                similar_folder = close_matches[0]
                if not is_as_built:
                    response = messagebox.askyesnocancel(
                        "資料夾名稱不符",
                        f"未找到 '{target_folder_name}' 資料夾，但找到相似的資料夾 '{similar_folder}'。\n"
                        f"是否要將 '{similar_folder}' 重新命名為 '{target_folder_name}'？\n"
                        f"選擇「是」將重命名資料夾，選擇「否」將使用現有資料夾而不重命名，選擇「取消」將終止程序。"
                    )
                    if response is True:
                        os.rename(os.path.join(root, similar_folder), os.path.join(root, target_folder_name))
                        dirs[dirs.index(similar_folder)] = target_folder_name
                        logging.info(f"資料夾已重新命名: {similar_folder} -> {target_folder_name}")
                    elif response is False:
                        target_folder_name = similar_folder
                        logging.info(f"使用現有相似資料夾: {similar_folder}")
                    else:
                        messagebox.showwarning("警告", f"未找到 '{target_folder_name}' 資料夾，程序將終止執行。")
                        raise SystemExit
                else:
                    target_folder_name = similar_folder  # AS BUILT 模式不主動重新命名
            else:
                if not is_as_built:
                    response = messagebox.askyesno(
                        "資料夾名稱不符",
                        f"未找到 '{target_folder_name}' 資料夾。\n是否要建立該資料夾？"
                    )
                    if response:
                        os.makedirs(os.path.join(root, target_folder_name), exist_ok=True)
                        dirs.append(target_folder_name)
                        logging.info(f"已建立資料夾: {target_folder_name}")
                    else:
                        messagebox.showwarning("警告", f"未找到 '{target_folder_name}' 資料夾，程序將終止執行。")
                        raise SystemExit
                else:
                    os.makedirs(os.path.join(root, target_folder_name), exist_ok=True)
                    dirs.append(target_folder_name)
                    logging.info(f"已建立資料夾: {target_folder_name}")

        summary_folder = os.path.join(root, target_folder_name)
        logging.info(f"處理資料夾: {summary_folder}")

        pdf_files = [os.path.join(summary_folder, f) for f in os.listdir(summary_folder)
                     if f.endswith('.pdf') and not f.startswith('~$')]

        with ThreadPoolExecutor(max_workers=os.cpu_count()) as executor:
            ndt_futures = {executor.submit(get_ndt_codes_from_pdf, file_path, cache): file_path for file_path in pdf_files}
            welding_futures = {executor.submit(get_welding_codes_from_pdf, file_path, cache): file_path for file_path in pdf_files}

            for future in as_completed(ndt_futures):
                file_path = ndt_futures[future]
                try:
                    ndt_codes = future.result()
                    ndt_codes_with_filenames_total.update(ndt_codes)
                except Exception as e:
                    logging.error(f"錯誤提取 NDT 編號從檔案 {file_path}: {e}")

            for future in as_completed(welding_futures):
                file_path = welding_futures[future]
                try:
                    welding_codes = future.result()
                    welding_codes_total.update(welding_codes)
                except Exception as e:
                    logging.error(f"錯誤提取焊材材證編號從檔案 {file_path}: {e}")
        break
    return ndt_codes_with_filenames_total, welding_codes_total

# 檔案搜尋與複製函數
def search_and_copy_ndt_pdfs(source_folder, target_folder, codes_with_filenames, is_as_built, cache):
    """搜尋並複製 NDT PDF 檔案，避免複製「作廢」版本。"""
    copied_files = 0
    not_found_codes = set(codes_with_filenames.keys())
    target_folder_path = os.path.join(target_folder, NDT_REPORTS_FOLDER_AS_BUILT if is_as_built else NDT_REPORTS_FOLDER_GENERAL)
    os.makedirs(target_folder_path, exist_ok=True)

    ndt_files = {}
    for root, _, files in os.walk(source_folder):
        for file in files:
            if file.endswith('.pdf') and not file.startswith('~$'):
                match = re.search(r'CWP-Q-R-JK-NDT-(\d+)', file, re.IGNORECASE)
                if match:
                    ndt_code = match.group(1)
                    is_cancelled = "作廢" in file
                    if ndt_code not in ndt_files:
                        ndt_files[ndt_code] = {'valid': None, 'cancelled': None}
                    ndt_files[ndt_code]['cancelled' if is_cancelled else 'valid'] = os.path.join(root, file)

    for ndt_code in list(not_found_codes):
        if ndt_code in ndt_files:
            file_to_copy = ndt_files[ndt_code]['valid']
            if file_to_copy:
                target_file_path = os.path.join(target_folder_path, os.path.basename(file_to_copy))
                try:
                    shutil.copy2(file_to_copy, target_file_path)
                    copied_files += 1
                    not_found_codes.remove(ndt_code)
                    logging.info(f"已複製 NDT 檔案: {file_to_copy} -> {target_file_path}")
                except Exception as e:
                    logging.error(f"無法複製檔案 {file_to_copy} 到 {target_file_path}: {e}")
            elif ndt_files[ndt_code]['cancelled']:
                logging.info(f"找到作廢的 NDT 檔案，但不複製: {ndt_files[ndt_code]['cancelled']}")
                not_found_codes.remove(ndt_code)

    not_found_filenames = {codes_with_filenames[code] for code in not_found_codes}
    return copied_files, not_found_filenames

def search_and_copy_welding_pdfs(source_folder, target_folder, codes, is_as_built, cache):
    """搜尋並複製焊材材證 PDF 檔案。"""
    copied_files = 0
    not_found_codes = set(codes)
    target_subfolder = WELDING_CONSUMABLE_FOLDER_AS_BUILT if is_as_built else WELDING_CONSUMABLE_FOLDER_GENERAL
    target_folder_path = os.path.join(target_folder, target_subfolder)
    os.makedirs(target_folder_path, exist_ok=True)

    for root, _, files in os.walk(source_folder):
        for file in files:
            if file.endswith('.pdf') and not file.startswith('~$'):
                for code in list(not_found_codes):
                    if re.search(re.escape(code), file, re.IGNORECASE):
                        source_file_path = os.path.join(root, file)
                        target_file_path = os.path.join(target_folder_path, file)
                        try:
                            shutil.copy2(source_file_path, target_file_path)
                            copied_files += 1
                            not_found_codes.remove(code)
                            logging.info(f"已複製焊材材證檔案: {source_file_path} -> {target_file_path}")
                        except Exception as e:
                            logging.error(f"無法複製檔案 {source_file_path} 到 {target_file_path}: {e}")
                        break
    return copied_files, not_found_codes

# 檔案刪除函數
def delete_all_welding_pdfs(target_folder, is_as_built):
    """刪除目標資料夾中所有的焊材材證 PDF 檔案。"""
    deleted_files = 0
    target_subfolder = WELDING_CONSUMABLE_FOLDER_AS_BUILT if is_as_built else WELDING_CONSUMABLE_FOLDER_GENERAL
    target_folder_path = os.path.join(target_folder, target_subfolder)

    if os.path.exists(target_folder_path):
        for file in os.listdir(target_folder_path):
            if file.endswith('.pdf') and not file.startswith('~$'):
                file_path = os.path.join(target_folder_path, file)
                try:
                    os.remove(file_path)
                    deleted_files += 1
                    logging.info(f"已刪除焊材材證檔案: {file_path}")
                except Exception as e:
                    logging.error(f"無法刪除檔案 {file_path}: {e}")
    return deleted_files

def delete_all_ndt_pdfs(target_folder, is_as_built):
    """刪除目標資料夾中所有的 NDT PDF 檔案。"""
    deleted_files = 0
    target_folder_path = os.path.join(target_folder, NDT_REPORTS_FOLDER_AS_BUILT if is_as_built else NDT_REPORTS_FOLDER_GENERAL)

    if os.path.exists(target_folder_path):
        for file in os.listdir(target_folder_path):
            if file.endswith('.pdf') and not file.startswith('~$'):
                file_path = os.path.join(target_folder_path, file)
                try:
                    os.remove(file_path)
                    deleted_files += 1
                    logging.info(f"已刪除 NDT 檔案: {file_path}")
                except Exception as e:
                    logging.error(f"無法刪除檔案 {file_path}: {e}")
    return deleted_files

def clean_unmatched_files(pdf_folder, is_as_built):
    """清理不符合命名規則的檔案。"""
    deleted_files = []
    try:
        folder_name = os.path.basename(pdf_folder)
        base_folder_name = extract_base_folder_name(folder_name)
        if not base_folder_name:
            logging.warning(f"警告：無法從資料夾名稱中提取基本名稱: {folder_name}")
            return deleted_files

        folder_number_match = re.search(r'#(\d+)$', base_folder_name)
        folder_number = folder_number_match.group(1) if folder_number_match else r'\d+'

        base_name_without_number = re.sub(r'#\d+$', '', base_folder_name)

        patterns = [
            re.compile(re.escape(base_folder_name), re.IGNORECASE),
            re.compile(r'CWP\d+[A-Z]-' + re.escape(base_folder_name), re.IGNORECASE),
            re.compile(
                r'.*?' + re.escape(base_name_without_number.split('.')[0]) +
                r'(?:\.\d{3})?' + f'#{folder_number}' + r'(?:\.\d{3})?.*',
                re.IGNORECASE
            )
        ]

        target_folders = [SUMMARY_FOLDER_NAME_AS_BUILT, MATERIAL_TRACEABILITY_FOLDER_AS_BUILT] if is_as_built else [SUMMARY_FOLDER_NAME_GENERAL, MATERIAL_TRACEABILITY_FOLDER_GENERAL]

        logging.info(f"資料夾基本名稱: {base_folder_name}")
        logging.info(f"使用的匹配模式: {[p.pattern for p in patterns]}")

        for subfolder in target_folders:
            subfolder_path = os.path.join(pdf_folder, subfolder)
            if os.path.exists(subfolder_path):
                for root, _, files in os.walk(subfolder_path):
                    for file in files:
                        if file.endswith('.pdf') and not file.startswith('~$'):
                            file_path = os.path.join(root, file)
                            file_name_without_extension = os.path.splitext(file)[0]
                            logging.info(f"檢查檔案: {file_name_without_extension}")
                            if not any(p.search(file_name_without_extension) for p in patterns):
                                try:
                                    os.remove(file_path)
                                    deleted_files.append(file_path)
                                    logging.info(f"已刪除檔案: {file_path}")
                                except Exception as e:
                                    logging.error(f"處理檔案時發生錯誤 {file_path}: {str(e)}")
                    break
    except Exception as e:
        logging.error(f"處理資料夾時發生錯誤 {pdf_folder}: {str(e)}")
        messagebox.showerror("錯誤", f"處理資料夾時發生錯誤：\n{str(e)}")
    return deleted_files

# 新增檢查函數：檢查目標資料夾底下是否存在 PDF 檔案
def check_required_pdf_files(pdf_folder, is_as_built):
    """
    檢查每個目標資料夾下的指定子資料夾是否存在 PDF 檔案。
    一般模式下，檢查「01 Welding Identification Summary」與「02 Material Traceability & Mill Cert」；
    竣工模式下，檢查「04 Welding Identification Summary」與「02 Material Traceability」。
    對於物料追溯的資料夾僅檢查該層，不包含子資料夾。
    Returns:
        tuple: (缺少銲道追溯PDF的資料夾列表, 缺少物料追溯清單 PDF的資料夾列表)
    """
    missing_welding_identification = []
    missing_material_traceability = []
    welding_folder_name = SUMMARY_FOLDER_NAME_AS_BUILT if is_as_built else SUMMARY_FOLDER_NAME_GENERAL
    # 依據模式選擇物料追溯資料夾名稱
    traceability_folder_name = MATERIAL_TRACEABILITY_FOLDER_AS_BUILT if is_as_built else MATERIAL_TRACEABILITY_FOLDER_GENERAL

    def check_in_target_folder(target_folder):
        welding_path = os.path.join(target_folder, welding_folder_name)
        traceability_path = os.path.join(target_folder, traceability_folder_name)

        # 檢查 Welding Identification (銲道追溯) 資料夾（遞迴搜尋）
        welding_has_pdf = False
        if os.path.exists(welding_path) and os.path.isdir(welding_path):
            for root, dirs, files in os.walk(welding_path):
                for file in files:
                    if file.endswith('.pdf') and not file.startswith('~$'):
                        welding_has_pdf = True
                        break
                if welding_has_pdf:
                    break
        if not welding_has_pdf:
            missing_welding_identification.append(welding_path)

        # 檢查 Material Traceability (物料追溯清單) 資料夾（僅檢查該層，不遞迴）
        traceability_has_pdf = False
        if os.path.exists(traceability_path) and os.path.isdir(traceability_path):
            for file in os.listdir(traceability_path):
                if file.endswith('.pdf') and not file.startswith('~$'):
                    traceability_has_pdf = True
                    break
        if not traceability_has_pdf:
            missing_material_traceability.append(traceability_path)

    if is_target_folder(os.path.basename(pdf_folder)):
        check_in_target_folder(pdf_folder)
    else:
        for name in os.listdir(pdf_folder):
            sub_path = os.path.join(pdf_folder, name)
            if os.path.isdir(sub_path) and is_target_folder(name):
                check_in_target_folder(sub_path)
    return missing_welding_identification, missing_material_traceability

# 單一資料夾處理函數
def process_single_folder(pdf_folder, ndt_source_pdf_folder, welding_source_pdf_folder, is_as_built, cache):
    """處理單一資料夾中的所有操作。"""
    ndt_codes_with_filenames_total, welding_codes_total = process_pdf_files_in_folder(pdf_folder, is_as_built, cache)
    reasons = []

    if not ndt_codes_with_filenames_total:
        reasons.append("沒有在 PDF 中找到任何符合條件的 NDT 編號。")
    if not welding_codes_total:
        reasons.append("沒有在 PDF 中找到任何符合條件的焊材材證編號。")

    delete_all_ndt_pdfs(pdf_folder, is_as_built)
    delete_all_welding_pdfs(pdf_folder, is_as_built)

    ndt_copied, not_found_ndt_filenames = search_and_copy_ndt_pdfs(
        ndt_source_pdf_folder, pdf_folder, ndt_codes_with_filenames_total, is_as_built, cache
    )
    if ndt_copied == 0 and ndt_codes_with_filenames_total:
        reasons.append("找到了 NDT 編號，但沒有找到對應的報驗單 PDF 檔案。")

    welding_copied, not_found_welding_codes = search_and_copy_welding_pdfs(
        welding_source_pdf_folder, pdf_folder, welding_codes_total, is_as_built, cache
    )
    if welding_copied == 0 and welding_codes_total:
        reasons.append("找到了焊材材證編號，但沒有找到對應的焊材材證 PDF 檔案。")

    deleted_files = clean_unmatched_files(pdf_folder, is_as_built)

    return ndt_copied, welding_copied, not_found_ndt_filenames, not_found_welding_codes, reasons, deleted_files

# 多個資料夾處理函數
def process_folders(pdf_folder, ndt_source_pdf_folder, welding_source_pdf_folder, cache):
    """處理所有目標資料夾。"""
    total_ndt_copied = 0
    total_welding_copied = 0
    not_found_ndt_filenames_total = set()
    not_found_welding_codes_total = set()
    reasons_total = []
    deleted_files_total = []

    is_as_built = "FOXWELL" in pdf_folder
    mode = "竣工模式" if is_as_built else "一般模式"
    logging.info(f"開始處理資料夾: {pdf_folder}，模式: {mode}")

    renamed_files = check_and_rename_files_in_folder(pdf_folder, cache)
    if renamed_files:
        message = "以下檔案已被重新命名：\n"
        message += "\n".join(f"舊名稱：{os.path.basename(old)} -> 新名稱：{os.path.basename(new)}" for old, new in renamed_files)
        messagebox.showinfo("檔案重命名", message)

    if is_target_folder(os.path.basename(pdf_folder)):
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
        for root, dirs, _ in os.walk(pdf_folder):
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
            break  # 僅處理第一層子資料夾

    logging.info(f"完成處理資料夾: {pdf_folder}")
    return total_ndt_copied, total_welding_copied, not_found_ndt_filenames_total, not_found_welding_codes_total, reasons_total, deleted_files_total

# 主函數
def main():
    """主函式。"""
    root = tk.Tk()
    root.withdraw()

    cache = FileCache()

    pdf_initialdir = "C:/Users/CWP-PC-E-COM302/Box/T460 風電 品管 簡瑞成/FAT package"
    pdf_folder = filedialog.askdirectory(
        initialdir=pdf_initialdir,
        title="請選擇包含銲道追溯檔案的資料夾:"
    )

    if not pdf_folder:
        messagebox.showwarning("警告", "未選擇任何資料夾，程序將終止。")
        return

    ndt_source_pdf_initialdir = "U:/N-品管部/@品管部共用資料區/專案/CWP06 台電二期專案/報驗單/001_JK報驗單/完成 (pdf)"
    ndt_source_pdf_folder = filedialog.askdirectory(
        initialdir=ndt_source_pdf_initialdir,
        title="請選擇包含報驗單 PDF 檔案的資料夾:"
    )

    if not ndt_source_pdf_folder:
        messagebox.showwarning("警告", "未選擇任何報驗單資料夾，程序將終止。")
        return

    welding_source_pdf_initialdir = "U:/N-品管部/@品管部共用資料區/品管人員資料夾/T460 風電 品管 簡瑞成/焊材材證"
    welding_source_pdf_folder = filedialog.askdirectory(
        initialdir=welding_source_pdf_initialdir,
        title="請選擇包含焊材材證 PDF 檔案的資料夾:"
    )

    if not welding_source_pdf_folder:
        messagebox.showwarning("警告", "未選擇任何焊材材證資料夾，程序將終止。")
        return

    is_as_built = "FOXWELL" in pdf_folder
    mode = "竣工模式" if is_as_built else "一般模式"

    total_ndt_copied, total_welding_copied, not_found_ndt_filenames_total, not_found_welding_codes_total, reasons_total, deleted_files_total = process_folders(
        pdf_folder, ndt_source_pdf_folder, welding_source_pdf_folder, cache
    )

    # 檢查每個目標資料夾下的 Welding Identification 與 Material Traceability 資料夾是否存在 PDF 檔案
    missing_welding_identification, missing_material_traceability = check_required_pdf_files(pdf_folder, is_as_built)
    warning_message = ""
    if missing_welding_identification:
        warning_message += "缺少銲道追溯清單 PDF:\n" + "\n".join(missing_welding_identification) + "\n\n"
    if missing_material_traceability:
        warning_message += "缺少物料追溯清單 PDF:\n" + "\n".join(missing_material_traceability)
    if warning_message:
        messagebox.showwarning("警告", warning_message)

    message = f"執行模式: {mode}\n\n"
    message += f"報驗單: 共複製了 {total_ndt_copied} 份。\n"
    message += f"焊材材證: 共複製了 {total_welding_copied} 份。\n\n"

    if not_found_ndt_filenames_total:
        missing_ndt_filenames_str = ", ".join(not_found_ndt_filenames_total)
        message += f"以下報驗單編號的檔案在報驗單資料夾中未找到，有可能尚未上傳：{missing_ndt_filenames_str}\n\n"

    if not_found_welding_codes_total:
        missing_welding_codes_str = ", ".join(not_found_welding_codes_total)
        message += f"以下焊材材證編號的檔案在焊材材證資料夾中未找到，有可能輸入有誤：{missing_welding_codes_str}"

    if total_ndt_copied == 0 and total_welding_copied == 0 and reasons_total:
        message += "\n\n沒有檔案被複製，具體原因如下：\n"
        message += "\n".join(reasons_total)

    if deleted_files_total:
        message += "\n\n以下檔案因為檔名不符合已被刪除：\n"
        for file_path in deleted_files_total:
            message += f"{file_path}\n"

    messagebox.showinfo("完成", message)

    cache.save_cache()
    logging.info("程式執行完成。")

if __name__ == "__main__":
    main()