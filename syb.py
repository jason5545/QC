import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageDraw, ImageFont
from PyPDF2 import PdfWriter, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import io
import os
import json
import shutil
from datetime import datetime
import threading
import time

class SignatureTool:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("PDF簽名工具")
        self.input_pdf_path = None
        self.positions = self.load_positions()
        self.setup_gui()
        self.pending_files = []  # 儲存待處理的檔案資訊
        self.start_file_check_thread()

    def load_positions(self):
        try:
            if os.path.exists('signature_positions.json'):
                with open('signature_positions.json', 'r') as f:
                    return json.load(f)
        except Exception:
            pass
        return {'Welding': None, 'Material': None}

    def setup_gui(self):
        self.window.geometry("400x300")
        
        # 日期輸入框
        date_frame = tk.Frame(self.window, pady=10)
        date_frame.pack(fill=tk.X, padx=20)
        tk.Label(date_frame, text="輸入日期:").pack(side=tk.LEFT)
        self.date_entry = tk.Entry(date_frame)
        self.date_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(10, 0))

        # 選擇PDF按鈕
        pdf_frame = tk.Frame(self.window, pady=10)
        pdf_frame.pack(fill=tk.X, padx=20)
        self.select_button = tk.Button(
            pdf_frame, 
            text="選擇PDF文件", 
            command=self.select_pdf
        )
        self.select_button.pack(fill=tk.X)
        
        self.file_label = tk.Label(pdf_frame, text="未選擇文件", wraplength=350)
        self.file_label.pack(fill=tk.X, pady=(5, 0))

        # 處理按鈕
        process_frame = tk.Frame(self.window, pady=10)
        process_frame.pack(fill=tk.X, padx=20)
        self.process_button = tk.Button(
            process_frame, 
            text="處理PDF", 
            
            command=self.process_single_pdf,
            state=tk.DISABLED
        )
        self.process_button.pack(fill=tk.X)

        # 狀態顯示
        status_frame = tk.Frame(self.window, pady=10)
        status_frame.pack(fill=tk.X, padx=20)
        self.status_label = tk.Label(status_frame, text="", wraplength=350)
        self.status_label.pack(fill=tk.X)

    def start_file_check_thread(self):
        """啟動檔案檢查線程"""
        self.file_check_thread = threading.Thread(target=self.check_pending_files, daemon=True)
        self.file_check_thread.start()

    def check_pending_files(self):
        """定期檢查並處理待處理的檔案"""
        while True:
            current_time = time.time()
            files_to_process = []
            files_to_remove = []

            # 檢查每個待處理的檔案
            for file_info in self.pending_files:
                if current_time - file_info['timestamp'] >= 1:  # 5秒後處理
                    files_to_process.append(file_info)
                    files_to_remove.append(file_info)

            # 處理檔案
            for file_info in files_to_process:
                try:
                    # 複製檔案
                    shutil.copy2(file_info['temp_path'], file_info['original_path'])
                    # 刪除臨時檔案
                    os.remove(file_info['temp_path'])
                    # 更新狀態
                    self.window.after(0, lambda: self.status_label.config(
                        text=f"檔案已成功更新: {os.path.basename(file_info['original_path'])}"
                    ))
                except Exception as e:
                    self.window.after(0, lambda error=str(e): self.status_label.config(
                        text=f"更新檔案時發生錯誤: {error}"
                    ))

            # 從待處理列表中移除已處理的檔案
            for file_info in files_to_remove:
                self.pending_files.remove(file_info)

            # 休眠一段時間再檢查
            time.sleep(5)

    def get_document_type(self, filename):
        if 'Welding' in filename:
            return 'Welding'
        elif 'Material' in filename:
            return 'Material'
        return None

    def create_signature_image(self):
        try:
            date_text = self.date_entry.get()
            if not date_text:
                self.status_label.config(text="請輸入日期")
                return None

            # 開啟原始簽名圖片並轉換為RGBA模式
            img = Image.open("紹宇.jpg").convert('RGBA')
            
            # 將灰色背景轉換為透明
            data = img.getdata()
            new_data = []
            # 設定閾值來決定哪些像素要變透明
            threshold = 100  # 可以調整這個值來改變透明化的程度
            
            for item in data:
                # 檢查像素是否接近灰色或白色
                if item[0] > threshold and item[1] > threshold and item[2] > threshold:
                    new_data.append((255, 255, 255, 0))  # 完全透明
                else:
                    new_data.append(item)  # 保持原來的顏色
                    
            img.putdata(new_data)
            
            # 設定文字和字體
            font = ImageFont.truetype("JasonHandwriting2-Regular.ttf", 100)

            # 獲取文字尺寸
            draw = ImageDraw.Draw(img)
            try:
                text_bbox = draw.textbbox((0, 0), date_text, font=font)
                text_width = text_bbox[2] - text_bbox[0]
                text_height = text_bbox[3] - text_bbox[1]
            except AttributeError:
                text_width, text_height = draw.textsize(date_text, font=font)

            # 創建新的透明圖片
            padding = 20
            new_image_width = img.width + text_width + padding
            new_image_height = max(img.height, text_height)
            new_img = Image.new('RGBA', (new_image_width, new_image_height), (255, 255, 255, 0))  # 完全透明背景
            
            # 貼上處理過的簽名圖片
            new_img.paste(img, (0, 0), img)

            # 添加日期文字到透明背景上
            draw_new_img = ImageDraw.Draw(new_img)
            text_y = (new_image_height - text_height) // 2
            text_position = (img.width + padding, text_y)
            draw_new_img.text(text_position, date_text, font=font, fill=(0, 0, 0, 255))  # 黑色文字，完全不透明

            # 縮小整個圖片
            target_height = 50
            scale_factor = target_height / new_img.height
            target_width = int(new_image_width * scale_factor)
            
            resized_img = new_img.resize((target_width, target_height), Image.Resampling.LANCZOS)

            temp_path = f"temp_signature_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            resized_img.save(temp_path, format='PNG')  # 使用PNG格式以保持透明度
            return temp_path

        except Exception as e:
            self.status_label.config(text=f"創建簽名圖片時發生錯誤: {e}")
            return None

    def add_signature_to_pdf(self, input_path, signature_path, x, y):
        try:
            # 創建臨時檔案路徑，使用時間戳來確保唯一性
            temp_output_path = f"{input_path}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.temp"
            
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=A4)
            
            # 使用傳入的位置，並設定固定的大小
            width = 40
            height = 15
            
            can.drawImage(signature_path, x-width/2, y-height/2, width, height, preserveAspectRatio=True, mask='auto')
            can.save()
            packet.seek(0)
            
            new_pdf = PdfReader(packet)
            existing_pdf = PdfReader(open(input_path, "rb"))
            output = PdfWriter()
            
            page = existing_pdf.pages[0]
            page.merge_page(new_pdf.pages[0])
            output.add_page(page)
            
            # 寫入臨時檔案
            with open(temp_output_path, "wb") as outputStream:
                output.write(outputStream)
            
            # 將檔案資訊加入待處理列表
            self.pending_files.append({
                'temp_path': temp_output_path,
                'original_path': input_path,
                'timestamp': time.time()
            })
            
            self.status_label.config(text="檔案處理完成！將在1分鐘後自動更新原始檔案")
            return True

        except Exception as e:
            self.status_label.config(text=f"處理PDF時發生錯誤: {e}")
            # 如果發生錯誤，嘗試清理臨時檔案
            if os.path.exists(temp_output_path):
                try:
                    os.remove(temp_output_path)
                except:
                    pass
            return False

    def select_pdf(self):
        self.input_pdf_path = filedialog.askopenfilename(
            title="選擇PDF文件",
            filetypes=[("PDF files", "*.pdf")]
        )
        if self.input_pdf_path:
import tkinter as tk
from tkinter import filedialog, ttk
from PIL import Image, ImageDraw, ImageFont
from PyPDF2 import PdfWriter, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import io
import os
import json
import shutil
from datetime import datetime
import threading
import time
import re

class SignatureTool:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("PDF簽名工具")
        self.root_folder = None
        self.positions = self.load_positions()
        self.setup_gui()
        self.processing = False

    def load_positions(self):
        try:
            if os.path.exists('signature_positions.json'):
                with open('signature_positions.json', 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            print(f"載入簽名位置時發生錯誤: {e}")
        return {'Welding': None, 'Material': None}

    def setup_gui(self):
        self.window.geometry("600x500")
        
        # 日期輸入框
        date_frame = tk.Frame(self.window, pady=10)
        date_frame.pack(fill=tk.X, padx=20)
        tk.Label(date_frame, text="輸入日期:").pack(side=tk.LEFT)
        self.date_entry = tk.Entry(date_frame)
        self.date_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(10, 0))

        # 選擇資料夾按鈕
        folder_frame = tk.Frame(self.window, pady=10)
        folder_frame.pack(fill=tk.X, padx=20)
        self.select_button = tk.Button(
            folder_frame, 
            text="選擇根目錄資料夾", 
            command=self.select_folder
        )
        self.select_button.pack(fill=tk.X)
        
        self.folder_label = tk.Label(folder_frame, text="未選擇資料夾", wraplength=550)
        self.folder_label.pack(fill=tk.X, pady=(5, 0))

        # 進度顯示區域
        progress_frame = tk.Frame(self.window, pady=10)
        progress_frame.pack(fill=tk.X, padx=20)
        self.progress_bar = ttk.Progressbar(
            progress_frame, 
            orient="horizontal", 
            mode="determinate"
        )
        self.progress_bar.pack(fill=tk.X)

        # 處理按鈕
        process_frame = tk.Frame(self.window, pady=10)
        process_frame.pack(fill=tk.X, padx=20)
        self.process_button = tk.Button(
            process_frame, 
            text="處理所有PDF", 
            command=self.process_all_pdfs,
            state=tk.DISABLED
        )
        self.process_button.pack(fill=tk.X)

        # 狀態顯示
        self.status_frame = tk.Frame(self.window, pady=10)
        self.status_frame.pack(fill=tk.X, padx=20)
        self.status_label = tk.Label(self.status_frame, text="", wraplength=550)
        self.status_label.pack(fill=tk.X)

        # 處理結果列表
        result_frame = tk.Frame(self.window, pady=10)
        result_frame.pack(fill=tk.BOTH, expand=True, padx=20)
        self.result_text = tk.Text(result_frame, height=10, wrap=tk.WORD)
        self.result_text.pack(fill=tk.BOTH, expand=True)
        scrollbar = tk.Scrollbar(result_frame, command=self.result_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text.config(yscrollcommand=scrollbar.set)

    def is_valid_folder_name(self, folder_name):
        pattern = r'XB1#\d+|XB[1-4][ABC]#\d+|6S21[1-7]#\d+|6S20[1256]#\d+$'
        return bool(re.match(pattern, folder_name))

    def find_target_pdfs(self, root_folder):
        target_pdfs = []
        
        for root, dirs, files in os.walk(root_folder):
            current_folder = os.path.basename(root)
            
            if current_folder in ["04 Welding Identification Summary", "02 Material Traceability"]:
                parent_folder = os.path.basename(os.path.dirname(root))
                if self.is_valid_folder_name(parent_folder):
                    for file in files:
                        if file.lower().endswith('.pdf'):
                            pdf_path = os.path.join(root, file)
                            doc_type = 'Welding' if "Welding" in current_folder else 'Material'
                            target_pdfs.append((pdf_path, doc_type))
        
        return target_pdfs

    def append_result(self, message):
        self.result_text.insert(tk.END, message + "\n")
        self.result_text.see(tk.END)

    def create_signature_image(self):
        try:
            date_text = self.date_entry.get()
            if not date_text:
                self.status_label.config(text="請輸入日期")
                return None

            img = Image.open("紹宇.jpg").convert('RGBA')
            
            data = img.getdata()
            new_data = []
            threshold = 100
            
            for item in data:
                if item[0] > threshold and item[1] > threshold and item[2] > threshold:
                    new_data.append((255, 255, 255, 0))
                else:
                    new_data.append(item)
                    
            img.putdata(new_data)
            
            font = ImageFont.truetype("JasonHandwriting2-Regular.ttf", 100)
            draw = ImageDraw.Draw(img)
            
            try:
                text_bbox = draw.textbbox((0, 0), date_text, font=font)
                text_width = text_bbox[2] - text_bbox[0]
                text_height = text_bbox[3] - text_bbox[1]
            except AttributeError:
                text_width, text_height = draw.textsize(date_text, font=font)

            padding = 20
            new_image_width = img.width + text_width + padding
            new_image_height = max(img.height, text_height)
            new_img = Image.new('RGBA', (new_image_width, new_image_height), (255, 255, 255, 0))
            
            new_img.paste(img, (0, 0), img)

            draw_new_img = ImageDraw.Draw(new_img)
            text_y = (new_image_height - text_height) // 2
            text_position = (img.width + padding, text_y)
            draw_new_img.text(text_position, date_text, font=font, fill=(0, 0, 0, 255))

            target_height = 50
            scale_factor = target_height / new_img.height
            target_width = int(new_image_width * scale_factor)
            
            resized_img = new_img.resize((target_width, target_height), Image.LANCZOS)

            temp_path = f"temp_signature_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            resized_img.save(temp_path, format='PNG')
            return temp_path

        except Exception as e:
            self.status_label.config(text=f"創建簽名圖片時發生錯誤: {e}")
            return None

    def add_signature_to_pdf(self, input_path, signature_path, x, y):
        temp_output_path = None
        try:
            # 創建臨時檔案路徑
            temp_dir = os.path.dirname(input_path)
            temp_output_path = os.path.join(temp_dir, f"temp_{os.path.basename(input_path)}")
            
            # 創建新的PDF頁面與簽名
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=A4)
            width = 40
            height = 15
            can.drawImage(signature_path, x - width / 2, y - height / 2, width, height, preserveAspectRatio=True, mask='auto')
            can.save()
            packet.seek(0)
            
            # 讀取新舊PDF
            new_pdf = PdfReader(packet)
            with open(input_path, "rb") as input_file:
                existing_pdf = PdfReader(input_file)
                output = PdfWriter()
                
                # 合併頁面
                page = existing_pdf.pages[0]
                page.merge_page(new_pdf.pages[0])
                output.add_page(page)
                
                # 寫入臨時檔案
                with open(temp_output_path, "wb") as output_file:
                    output.write(output_file)

            # 確保所有操作完成後替換檔案
            packet.close()
            time.sleep(0.1)  # 短暫等待確保檔案寫入完成

            if os.path.exists(temp_output_path):
                os.remove(input_path)  # 刪除原始檔案
                os.rename(temp_output_path, input_path)  # 重新命名臨時檔案
                self.append_result(f"成功更新: {os.path.basename(input_path)}")
                return True
            else:
                raise Exception("臨時檔案創建失敗")

        except Exception as e:
            self.append_result(f"處理PDF時發生錯誤 {os.path.basename(input_path)}: {str(e)}")
            if temp_output_path and os.path.exists(temp_output_path):
                try:
                    os.remove(temp_output_path)
                except:
                    pass
            return False

    def select_folder(self):
        self.root_folder = filedialog.askdirectory(title="選擇根目錄資料夾")
        if self.root_folder:
            self.folder_label.config(text=self.root_folder)
            self.process_button.config(state=tk.NORMAL)
            self.status_label.config(text="已選擇資料夾，請點擊處理所有PDF按鈕進行處理")
            self.result_text.delete(1.0, tk.END)
        else:
            self.folder_label.config(text="未選擇資料夾")
            self.process_button.config(state=tk.DISABLED)

    def process_pdf_with_position(self, pdf_info):
        pdf_path, doc_type = pdf_info
        if not self.positions.get(doc_type):
            self.append_result(f"跳過 {os.path.basename(pdf_path)}: 未設定簽名位置")
            return False

        signature_path = self.create_signature_image()
        if not signature_path:
            return False

        try:
            x, y = self.positions[doc_type]
            success = self.add_signature_to_pdf(pdf_path, signature_path, x, y)
            return success
        finally:
            if os.path.exists(signature_path):
                try:
                    os.remove(signature_path)
                except:
                    pass

    def process_all_pdfs(self):
        if self.processing:
            return

        if not self.root_folder:
            self.status_label.config(text="請先選擇根目錄資料夾")
            return

        self.processing = True
        self.process_button.config(state=tk.DISABLED)
        self.result_text.delete(1.0, tk.END)
        
        def process_thread():
            try:
                target_pdfs = self.find_target_pdfs(self.root_folder)
                if not target_pdfs:
                    self.window.after(0, lambda: self.status_label.config(text="未找到符合條件的PDF文件"))
                    return

                total = len(target_pdfs)
                self.window.after(0, lambda: self.progress_bar.configure(maximum=total))
                
                for i, pdf_info in enumerate(target_pdfs, 1):
                    pdf_path = pdf_info[0]
                    self.window.after(0, lambda p=pdf_path, i=i: self.status_label.config(
                        text=f"正在處理 {os.path.basename(p)} ({i}/{total})"
                    ))
                    self.window.after(0, lambda i=i: self.progress_bar.configure(value=i))
                    
                    self.process_pdf_with_position(pdf_info)
                    
                self.window.after(0, lambda: self.status_label.config(text="所有文件處理完成"))
            except Exception as e:
                self.window.after(0, lambda: se
