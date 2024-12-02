import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageDraw, ImageFont
from PyPDF2 import PdfWriter, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import io
import os
import shutil
from datetime import datetime
import threading
import time
import re
import fitz  # PyMuPDF

class SignatureTool:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("PDF簽名工具")
        self.selected_folder = None
        self.setup_gui()
        self.pending_files = []
        self.start_file_check_thread()
        self.signature_path = None

    def setup_gui(self):
        self.window.geometry("400x300")
        
        # 日期輸入框
        date_frame = tk.Frame(self.window, pady=10)
        date_frame.pack(fill=tk.X, padx=20)
        tk.Label(date_frame, text="輸入日期:").pack(side=tk.LEFT)
        self.date_entry = tk.Entry(date_frame)
        self.date_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(10, 0))
        # 設定預設日期
        self.date_entry.insert(0, "2024/7/1")

        # 選擇資料夾按鈕
        folder_frame = tk.Frame(self.window, pady=10)
        folder_frame.pack(fill=tk.X, padx=20)
        self.select_button = tk.Button(
            folder_frame, 
            text="選擇資料夾", 
            command=self.select_folder
        )
        self.select_button.pack(fill=tk.X)
        
        self.folder_label = tk.Label(folder_frame, text="未選擇資料夾", wraplength=350)
        self.folder_label.pack(fill=tk.X, pady=(5, 0))

        # 處理按鈕
        process_frame = tk.Frame(self.window, pady=10)
        process_frame.pack(fill=tk.X, padx=20)
        self.process_button = tk.Button(
            process_frame, 
            text="開始處理", 
            command=self.process_folder,
            state=tk.DISABLED
        )
        self.process_button.pack(fill=tk.X)

        # 狀態顯示
        status_frame = tk.Frame(self.window, pady=10)
        status_frame.pack(fill=tk.X, padx=20)
        self.status_label = tk.Label(status_frame, text="", wraplength=350)
        self.status_label.pack(fill=tk.X)

    def select_folder(self):
        self.selected_folder = filedialog.askdirectory(
            title="選擇資料夾"
        )
        if self.selected_folder:
            self.folder_label.config(text=self.selected_folder)
            self.process_button.config(state=tk.NORMAL)
            self.status_label.config(text="已選擇資料夾，請點擊開始處理按鈕進行處理")
        else:
            self.folder_label.config(text="未選擇資料夾")
            self.process_button.config(state=tk.DISABLED)

    def start_file_check_thread(self):
        self.file_check_thread = threading.Thread(target=self.check_pending_files, daemon=True)
        self.file_check_thread.start()

    def check_pending_files(self):
        while True:
            current_time = time.time()
            files_to_process = []
            files_to_remove = []

            for file_info in self.pending_files:
                if current_time - file_info['timestamp'] >= 1:
                    files_to_process.append(file_info)
                    files_to_remove.append(file_info)

            for file_info in files_to_process:
                try:
                    shutil.copy2(file_info['temp_path'], file_info['original_path'])
                    os.remove(file_info['temp_path'])
                    self.window.after(0, lambda: self.status_label.config(
                        text=f"檔案已成功更新: {os.path.basename(file_info['original_path'])}"
                    ))
                except Exception as e:
                    self.window.after(0, lambda error=str(e): self.status_label.config(
                        text=f"更新檔案時發生錯誤: {error}"
                    ))

            for file_info in files_to_remove:
                self.pending_files.remove(file_info)

            time.sleep(1)

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
            threshold = 100
            
            for item in data:
                if item[0] > threshold and item[1] > threshold and item[2] > threshold:
                    new_data.append((255, 255, 255, 0))
                else:
                    new_data.append(item)
            img.putdata(new_data)
            
            # 使用較大字體大小以保持清晰度
            font_size = 200
            font = ImageFont.truetype("JasonHandwriting2-Regular.ttf", font_size)
            draw = ImageDraw.Draw(img)
            
            try:
                text_bbox = draw.textbbox((0, 0), date_text, font=font)
                text_width = text_bbox[2] - text_bbox[0]
                text_height = text_bbox[3] - text_bbox[1]
            except AttributeError:
                text_width, text_height = draw.textsize(date_text, font=font)

            # 使用較大的 padding
            padding = 40
            new_image_width = img.width + text_width + padding
            new_image_height = max(img.height, text_height)
            new_img = Image.new('RGBA', (new_image_width, new_image_height), (255, 255, 255, 0))
            new_img.paste(img, (0, 0), img)

            draw_new_img = ImageDraw.Draw(new_img)
            text_y = (new_image_height - text_height) // 2
            text_position = (img.width + padding, text_y)
            draw_new_img.text(text_position, date_text, font=font, fill=(0, 0, 0, 255))

            # 最終輸出使用較大的尺寸
            target_height = 50  # 增加目標高度確保清晰度
            scale_factor = target_height / new_img.height
            target_width = int(new_image_width * scale_factor)
            
            resized_img = new_img.resize(
                (target_width, target_height), 
                Image.LANCZOS  # 使用高質量的重採樣方法
            )

            temp_path = f"temp_signature_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            resized_img.save(temp_path, format='PNG', quality=95, dpi=(300, 300))  # 設定較高的DPI和質量
            return temp_path

        except Exception as e:
            self.status_label.config(text=f"創建簽名圖片時發生錯誤: {e}")
            return None

    def find_text_position(self, pdf_path):
        try:
            doc = fitz.open(pdf_path)
            page = doc[0]  # 假設在第一頁
            text_instances = page.search_for("Reviewed by")
            
            if text_instances:
                # 取得最後一個 "Reviewed by" 的位置
                last_instance = text_instances[-1]
                return last_instance
            return None
        except Exception as e:
            self.status_label.config(text=f"搜尋文字位置時發生錯誤: {e}")
            return None
        finally:
            if 'doc' in locals():
                doc.close()

    def add_signature_to_pdf(self, input_path, signature_path, offset_y, offset_x=0):
        try:
            position = self.find_text_position(input_path)
            if not position:
                self.status_label.config(text=f"未找到'Reviewed by'文字位置: {os.path.basename(input_path)}")
                return False

            doc = fitz.open(input_path)
            page = doc[0]
            page_height = page.rect.height
            doc.close()

            # 設定簽名圖片在PDF中的固定大小
            width = 40  # 設定固定寬度
            height = 15  # 設定固定高度

            # 計算位置
            y = page_height - position.y1 - height - offset_y
            x = position.x0 + offset_x

            temp_output_path = f"{input_path}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.temp"
            
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=A4)
            can.drawImage(
                signature_path, 
                x, 
                y, 
                width, 
                height, 
                preserveAspectRatio=True, 
                mask='auto'
            )
            can.save()
            
            packet.seek(0)
            new_pdf = PdfReader(packet)
            existing_pdf = PdfReader(open(input_path, "rb"))
            output = PdfWriter()
            
            page = existing_pdf.pages[0]
            page.merge_page(new_pdf.pages[0])
            output.add_page(page)
            
            with open(temp_output_path, "wb") as outputStream:
                output.write(outputStream)
            
            self.pending_files.append({
                'temp_path': temp_output_path,
                'original_path': input_path,
                'timestamp': time.time()
            })
            return True
        except Exception as e:
            self.status_label.config(text=f"處理PDF時發生錯誤: {e}")
            if 'temp_output_path' in locals() and os.path.exists(temp_output_path):
                try:
                    os.remove(temp_output_path)
                except:
                    pass
            return False

    def process_folder(self):
        if not self.selected_folder:
            self.status_label.config(text="請先選擇資料夾")
            return

        regex_pattern = r'XB1#\d+|XB[1-4][ABC]#\d+|6S21[1-7]#\d+|6S20[12356]#\d+'
        pattern = re.compile(regex_pattern)

        for root, dirs, files in os.walk(self.selected_folder):
            for dir_name in dirs:
                if pattern.match(dir_name):
                    full_dir_path = os.path.join(root, dir_name)
                    target_subfolders = ["04 Welding Identification Summary", "02 Material Traceability"]
                    
                    for subfolder in target_subfolders:
                        if subfolder == "04 Welding Identification Summary":
                            offset_y = -10
                            offset_x = 50
                        elif subfolder == "02 Material Traceability":
                            offset_y = 3
                            offset_x = 0

                        subfolder_path = os.path.join(full_dir_path, subfolder)
                        if os.path.isdir(subfolder_path):
                            pdf_files = [f for f in os.listdir(subfolder_path) if f.endswith('.pdf')]
                            
                            for pdf_file in pdf_files:
                                pdf_path = os.path.join(subfolder_path, pdf_file)
                                self.signature_path = self.create_signature_image()
                                if self.signature_path:
                                    self.add_signature_to_pdf(pdf_path, self.signature_path, offset_y, offset_x)
                                    os.remove(self.signature_path)
                                else:
                                    self.status_label.config(text="簽名圖像生成失敗，跳過此文件")

        self.status_label.config(text="處理完畢，請查看資料夾中的PDF文件")

    def run(self):
        self.window.mainloop()

if __name__ == "__main__":
    app = SignatureTool()
    app.run()