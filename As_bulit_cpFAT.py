import os
import json
import shutil
import logging
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime

class FolderCopyApp:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("資料夾複製工具")
        self.window.geometry("1000x700")  # 調整視窗大小以適應所有元件

        # 預設的資料夾選項
        self.folder_options = [
            "XB1", "6S201", "XB2B", "6S203", "XB3B", 
            "6S206", "XB3B.002", "XB4B", "XB4B.002"
        ]

        # 儲存映射資料的字典
        self.folder_mapping = {}

        # 儲存資料夾連動關係
        self.folder_links = {}

        # 載入預設或已存在的映射
        self.config_file = "folder_mapping.json"
        self.links_file = "folder_links.json"
        self.load_mapping()
        self.load_links()

        # 設置日誌
        self.setup_logging()

        self.create_widgets()

    def setup_logging(self):
        """設置日誌記錄"""
        self.log_filename = f'folder_copy_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(self.log_filename, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )

    def create_widgets(self):
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 使用 PanedWindow 來更好地管理元件佈局
        paned_window = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned_window.pack(fill=tk.BOTH, expand=True)

        # 左側 - 資料夾列表
        left_frame = ttk.Frame(paned_window)
        paned_window.add(left_frame, weight=1)

        # 資料夾列表
        list_frame = ttk.LabelFrame(left_frame, text="資料夾列表")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Treeview for folder list
        self.folder_tree = ttk.Treeview(list_frame, columns=("folder", "serials"), show="headings")
        self.folder_tree.heading("folder", text="資料夾名稱")
        self.folder_tree.heading("serials", text="流水號")
        self.folder_tree.column("folder", width=150)
        self.folder_tree.column("serials", width=250)
        self.folder_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Scrollbar for Treeview
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.folder_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.folder_tree.configure(yscrollcommand=scrollbar.set)

        # 右側 - 編輯區域
        right_frame = ttk.Frame(paned_window)
        paned_window.add(right_frame, weight=1)

        # 編輯區域
        edit_frame = ttk.LabelFrame(right_frame, text="編輯資料")
        edit_frame.pack(fill=tk.X, expand=False, padx=5, pady=5)

        # 資料夾名稱下拉選單
        folder_label = ttk.Label(edit_frame, text="資料夾名稱:")
        folder_label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.folder_combobox = ttk.Combobox(edit_frame, values=self.folder_options)
        self.folder_combobox.grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)

        # 流水號輸入
        serials_label = ttk.Label(edit_frame, text="流水號 (用逗號分隔):")
        serials_label.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.serials_entry = ttk.Entry(edit_frame)
        self.serials_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.EW)

        # 資料夾連動輸入
        links_label = ttk.Label(edit_frame, text="連動資料夾 (用逗號分隔):")
        links_label.grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.links_entry = ttk.Entry(edit_frame)
        self.links_entry.grid(row=2, column=1, padx=5, pady=5, sticky=tk.EW)

        # 設置列的權重
        edit_frame.columnconfigure(1, weight=1)

        # 按鈕區域
        button_frame = ttk.Frame(edit_frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=10)

        ttk.Button(button_frame, text="新增/更新", command=self.add_or_update_mapping).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="刪除", command=self.delete_mapping).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="清空", command=self.clear_entries).pack(side=tk.LEFT, padx=5)

        # 複製設定區域
        copy_frame = ttk.LabelFrame(right_frame, text="複製設定")
        copy_frame.pack(fill=tk.X, expand=False, padx=5, pady=5)

        # 來源資料夾選擇
        source_frame = ttk.Frame(copy_frame)
        source_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(source_frame, text="來源資料夾:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.source_path = tk.StringVar()
        ttk.Entry(source_frame, textvariable=self.source_path).grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)
        ttk.Button(source_frame, text="瀏覽", command=lambda: self.browse_folder("source")).grid(row=0, column=2, padx=5, pady=5)
        source_frame.columnconfigure(1, weight=1)

        # 目標資料夾選擇
        target_frame = ttk.Frame(copy_frame)
        target_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(target_frame, text="目標資料夾:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.target_path = tk.StringVar()
        ttk.Entry(target_frame, textvariable=self.target_path).grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)
        ttk.Button(target_frame, text="瀏覽", command=lambda: self.browse_folder("target")).grid(row=0, column=2, padx=5, pady=5)
        target_frame.columnconfigure(1, weight=1)

        # 底部按鈕
        bottom_frame = ttk.Frame(self.window)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=5)

        ttk.Button(bottom_frame, text="儲存設定", command=self.save_mapping).pack(side=tk.LEFT, padx=5)
        ttk.Button(bottom_frame, text="開始複製", command=self.start_copy).pack(side=tk.RIGHT, padx=5)

        # 綁定選擇事件
        self.folder_tree.bind("<<TreeviewSelect>>", self.on_select)

        # 更新顯示
        self.update_folder_list()

    def browse_folder(self, folder_type):
        """瀏覽資料夾"""
        folder_path = filedialog.askdirectory(title=f"選擇{folder_type}資料夾")
        if folder_path:
            if folder_type == "source":
                self.source_path.set(folder_path)
            else:
                self.target_path.set(folder_path)

    def load_mapping(self):
        """載入映射設定"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    self.folder_mapping = json.load(f)
            else:
                # 預設映射
                self.folder_mapping = {
                    "XB1": ["071", "079", "091", "095"],
                    "6S201": ["071", "079", "091", "095"],
                    "XB2B": ["016", "021", "034", "043"],
                    "6S203": ["016", "021", "034", "043"],
                    "XB3B": ["011", "014", "015", "016"],
                    "6S206": ["011", "014", "015", "016"],
                    "XB3B.002": ["011", "014", "015", "016"],
                    "XB4B": ["001", "015", "017", "019"],
                    "XB4B.002": ["001", "015", "017", "019"]
                }
        except Exception as e:
            messagebox.showerror("錯誤", f"載入設定時發生錯誤: {e}")

    def load_links(self):
        """載入資料夾連動關係"""
        try:
            if os.path.exists(self.links_file):
                with open(self.links_file, 'r', encoding='utf-8') as f:
                    self.folder_links = json.load(f)
            else:
                # 預設連動關係
                self.folder_links = {
                    "XB1": ["6S201"],
                    "6S201": ["XB1"],
                    "XB2B": ["6S203"],
                    "6S203": ["XB2B"],
                    "XB3B": ["6S206", "XB3B.002"],
                    "6S206": ["XB3B", "XB3B.002"],
                    "XB3B.002": ["XB3B", "6S206"],
                    "XB4B": ["XB4B.002"],
                    "XB4B.002": ["XB4B"]
                }
        except Exception as e:
            messagebox.showerror("錯誤", f"載入連動關係時發生錯誤: {e}")

    def save_mapping(self):
        """儲存映射設定"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.folder_mapping, f, ensure_ascii=False, indent=4)
            with open(self.links_file, 'w', encoding='utf-8') as f:
                json.dump(self.folder_links, f, ensure_ascii=False, indent=4)
            messagebox.showinfo("成功", "設定已儲存")
        except Exception as e:
            messagebox.showerror("錯誤", f"儲存設定時發生錯誤: {e}")

    def update_folder_list(self):
        """更新資料夾列表顯示"""
        for item in self.folder_tree.get_children():
            self.folder_tree.delete(item)

        for folder, serials in self.folder_mapping.items():
            self.folder_tree.insert("", tk.END, values=(folder, ", ".join(serials)))

    def get_all_linked_folders(self, folder_name):
        """取得與指定資料夾連動的所有資料夾（包含自身）"""
        linked = set()
        stack = [folder_name]
        while stack:
            current = stack.pop()
            if current not in linked:
                linked.add(current)
                neighbors = self.folder_links.get(current, [])
                stack.extend(neighbors)
        return linked

    def add_or_update_mapping(self):
        """新增或更新映射"""
        folder_name = self.folder_combobox.get().strip()
        serials = [s.strip() for s in self.serials_entry.get().split(",") if s.strip()]
        links = [s.strip() for s in self.links_entry.get().split(",") if s.strip()]

        if not folder_name:
            messagebox.showwarning("警告", "請輸入資料夾名稱")
            return

        if serials:
            # 更新資料夾及其連動資料夾的映射
            linked_folders = self.get_all_linked_folders(folder_name)
            for fname in linked_folders:
                self.folder_mapping[fname] = serials
        else:
            # 如果流水號清空，則刪除該資料夾的映射
            if folder_name in self.folder_mapping:
                del self.folder_mapping[folder_name]

        # 更新資料夾連動關係
        if links:
            self.folder_links[folder_name] = links
            # 確保雙向連動
            for link in links:
                if link in self.folder_links:
                    if folder_name not in self.folder_links[link]:
                        self.folder_links[link].append(folder_name)
                else:
                    self.folder_links[link] = [folder_name]
        else:
            if folder_name in self.folder_links:
                del self.folder_links[folder_name]
            # 刪除其他資料夾中指向該資料夾的連動
            for key in self.folder_links:
                if folder_name in self.folder_links[key]:
                    self.folder_links[key].remove(folder_name)

        self.update_folder_list()
        self.clear_entries()

    def delete_mapping(self):
        """刪除選定的映射"""
        selected_item = self.folder_tree.selection()
        if not selected_item:
            messagebox.showwarning("警告", "請選擇要刪除的項目")
            return

        folder_name = self.folder_tree.item(selected_item[0])["values"][0]
        linked_folders = self.get_all_linked_folders(folder_name)
        for fname in linked_folders:
            if fname in self.folder_mapping:
                del self.folder_mapping[fname]
            if fname in self.folder_links:
                del self.folder_links[fname]

        # 刪除其他資料夾中指向這些資料夾的連動
        for key in list(self.folder_links.keys()):
            for fname in linked_folders:
                if fname in self.folder_links.get(key, []):
                    self.folder_links[key].remove(fname)
            if not self.folder_links.get(key):
                del self.folder_links[key]

        self.update_folder_list()
        self.clear_entries()

    def clear_entries(self):
        """清空輸入欄位"""
        self.folder_combobox.set("")
        self.serials_entry.delete(0, tk.END)
        self.links_entry.delete(0, tk.END)

    def on_select(self, event):
        """處理選擇事件"""
        selected_item = self.folder_tree.selection()
        if selected_item:
            folder_name = self.folder_tree.item(selected_item[0])["values"][0]
            serials = self.folder_mapping.get(folder_name, [])
            links = self.folder_links.get(folder_name, [])

            self.folder_combobox.set(folder_name)
            self.serials_entry.delete(0, tk.END)
            self.serials_entry.insert(0, ", ".join(serials))
            self.links_entry.delete(0, tk.END)
            self.links_entry.insert(0, ", ".join(links))

    def get_folder_size(self, folder_path):
        """計算資料夾大小"""
        total_size = 0
        try:
            for dirpath, dirnames, filenames in os.walk(folder_path):
                for f in filenames:
                    fp = os.path.join(dirpath, f)
                    total_size += os.path.getsize(fp)
        except Exception as e:
            logging.error(f"計算資料夾大小時發生錯誤: {e}")
        return total_size

    def check_disk_space(self, target_folder, required_space):
        """檢查目標路徑是否有足夠空間"""
        try:
            free_space = shutil.disk_usage(target_folder).free
            return free_space > required_space * 1.1  # 預留10%空間
        except Exception as e:
            logging.error(f"檢查磁碟空間時發生錯誤: {e}")
            return False

    def copy_with_progress(self, source, destination, progress_callback):
        """複製檔案並更新進度"""
        total_size = os.path.getsize(source)
        copied_size = 0
        buffer_size = 1024 * 1024  # 1MB

        with open(source, 'rb') as src_file:
            with open(destination, 'wb') as dest_file:
                while True:
                    buffer = src_file.read(buffer_size)
                    if not buffer:
                        break
                    dest_file.write(buffer)
                    copied_size += len(buffer)
                    progress_callback(len(buffer), total_size)

    def copy_folder_with_progress(self, source_folder, dest_folder, file_progress_callback, overall_progress_callback):
        """複製資料夾並更新進度"""
        if not os.path.exists(dest_folder):
            os.makedirs(dest_folder)
        total_files = sum(len(files) for _, _, files in os.walk(source_folder))
        copied_files = 0

        for root, dirs, files in os.walk(source_folder):
            relative_path = os.path.relpath(root, source_folder)
            dest_root = os.path.join(dest_folder, relative_path)
            if not os.path.exists(dest_root):
                os.makedirs(dest_root)
            for file in files:
                source_file = os.path.join(root, file)
                dest_file = os.path.join(dest_root, file)
                file_progress_callback(source_file)
                try:
                    self.copy_with_progress(source_file, dest_file, lambda copied, total: None)
                except Exception as e:
                    logging.error(f"複製檔案失敗: {source_file} -> {dest_file}, 錯誤: {e}")
                copied_files += 1
                overall_progress_callback(copied_files, total_files)

    def search_and_copy_folders(self, source_folder, target_folder):
        """搜尋並複製符合條件的資料夾"""
        success_count = 0
        failed_count = 0
        skipped_count = 0

        # 確保目標資料夾存在
        ljb_folder = os.path.join(target_folder, "LJB#")
        ujb_folder = os.path.join(target_folder, "UJB")
        os.makedirs(ljb_folder, exist_ok=True)
        os.makedirs(ujb_folder, exist_ok=True)

        # 收集符合條件的資料夾
        matched_folders = []
        for root, dirs, _ in os.walk(source_folder):
            for dir_name in dirs:
                for folder_name, serials in self.folder_mapping.items():
                    if any(f"{folder_name}#{serial}" in dir_name for serial in serials):
                        matched_folders.append((root, dir_name))
                        break

        if not matched_folders:
            logging.warning("未找到符合條件的資料夾")
            messagebox.showwarning("警告", "未找到符合條件的資料夾")
            return success_count, failed_count, skipped_count

        # 計算所需空間
        total_size = sum(self.get_folder_size(os.path.join(root, dir_name)) 
                        for root, dir_name in matched_folders)

        # 檢查磁碟空間
        if not self.check_disk_space(target_folder, total_size):
            logging.error("目標磁碟空間不足")
            messagebox.showerror("錯誤", "目標磁碟空間不足")
            return success_count, failed_count, skipped_count

        # 複製資料夾
        progress_window = tk.Toplevel(self.window)
        progress_window.title("複製進度")
        progress_window.geometry("500x300")

        progress_label = ttk.Label(progress_window, text="正在複製資料夾...")
        progress_label.pack(pady=5)

        current_file_label = ttk.Label(progress_window, text="")
        current_file_label.pack(pady=5)

        overall_progress_bar = ttk.Progressbar(progress_window, length=400, mode='determinate', maximum=len(matched_folders))
        overall_progress_bar.pack(pady=5)

        overall_progress_label = ttk.Label(progress_window, text="整體進度")
        overall_progress_label.pack(pady=5)

        file_progress_bar = ttk.Progressbar(progress_window, length=400, mode='determinate')
        file_progress_bar.pack(pady=5)

        file_progress_label = ttk.Label(progress_window, text="檔案進度")
        file_progress_label.pack(pady=5)

        progress_window.update()

        # 開始複製
        for idx, (root, dir_name) in enumerate(matched_folders):
            source_dir = os.path.join(root, dir_name)

            # 按照規則決定目標資料夾
            if "XB4B" in dir_name:
                dest_dir = os.path.join(ljb_folder, dir_name)
            else:
                dest_dir = os.path.join(ujb_folder, dir_name)

            try:
                if os.path.exists(dest_dir):
                    # 目標資料夾已存在，跳過或覆蓋
                    # 這裡假設跳過
                    logging.info(f"資料夾已存在，跳過: {dest_dir}")
                    skipped_count += 1
                else:
                    def file_progress_callback(current_file):
                        current_file_label.config(text=f"正在複製: {current_file}")
                        progress_window.update()

                    def overall_progress_callback(copied_files, total_files):
                        file_progress_bar['maximum'] = total_files
                        file_progress_bar['value'] = copied_files
                        file_progress_label.config(text=f"檔案進度: {copied_files}/{total_files}")
                        progress_window.update()

                    self.copy_folder_with_progress(
                        source_dir, dest_dir, 
                        file_progress_callback=file_progress_callback, 
                        overall_progress_callback=overall_progress_callback
                    )
                    logging.info(f"成功複製: {source_dir} -> {dest_dir}")
                    success_count += 1
            except Exception as e:
                logging.error(f"複製失敗: {source_dir} -> {dest_dir}, 錯誤: {e}")
                failed_count += 1

            # 更新整體進度條
            overall_progress_bar['value'] = idx + 1
            progress_label.config(text=f"正在複製資料夾 {idx + 1}/{len(matched_folders)}")
            progress_window.update()

        progress_window.destroy()

        return success_count, failed_count, skipped_count

    def start_copy(self):
        """開始複製"""
        source_folder = self.source_path.get()
        target_folder = self.target_path.get()

        if not source_folder or not target_folder:
            messagebox.showwarning("警告", "請選擇來源和目標資料夾")
            return

        # 檢查來源和目標資料夾是否存在
        if not os.path.exists(source_folder):
            messagebox.showerror("錯誤", "來源資料夾不存在")
            return
        if not os.path.exists(target_folder):
            os.makedirs(target_folder)

        # 開始複製
        success_count, failed_count, skipped_count = self.search_and_copy_folders(source_folder, target_folder)

        # 顯示結果
        messagebox.showinfo("完成", f"複製完成\n成功: {success_count}\n失敗: {failed_count}\n跳過: {skipped_count}")

if __name__ == "__main__":
    app = FolderCopyApp()
    app.window.mainloop()
