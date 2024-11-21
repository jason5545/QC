import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageDraw, ImageFont
from PyPDF2 import PdfWriter, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import io
import os
import json
from datetime import datetime

class SignatureTool:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("PDF簽名工具")
        self.input_pdf_path = None
        self.positions = self.load_positions()
        self.setup_gui()

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

    def add_signature_to_pdf(self, input_path, output_path, signature_path, x, y):
        try:
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
            
            with open(output_path, "wb") as outputStream:
                output.write(outputStream)
            
            return True

        except Exception as e:
            self.status_label.config(text=f"處理PDF時發生錯誤: {e}")
            return False

    def select_pdf(self):
        self.input_pdf_path = filedialog.askopenfilename(
            title="選擇PDF文件",
            filetypes=[("PDF files", "*.pdf")]
        )
        if self.input_pdf_path:
            self.file_label.config(text=self.input_pdf_path)
            self.process_button.config(state=tk.NORMAL)
            self.status_label.config(text="已選擇文件，請點擊處理PDF按鈕進行處理")
        else:
            self.file_label.config(text="未選擇文件")
            self.process_button.config(state=tk.DISABLED)

    def process_pdf_with_position(self, x, y):
        try:
            signature_path = self.create_signature_image()
            if not signature_path:
                return

            output_dir = os.path.dirname(self.input_pdf_path)
            filename = os.path.basename(self.input_pdf_path)
            output_path = os.path.join(output_dir, f"signed_{filename}")

            if self.add_signature_to_pdf(self.input_pdf_path, output_path, signature_path, x, y):
                self.status_label.config(text=f"處理完成！\n儲存於: {output_path}")
            
            os.remove(signature_path)

        except Exception as e:
            self.status_label.config(text=f"處理過程中發生錯誤: {e}")

    def process_single_pdf(self):
        if not self.input_pdf_path:
            self.status_label.config(text="請先選擇PDF文件")
            return

        doc_type = self.get_document_type(os.path.basename(self.input_pdf_path))
        if not doc_type:
            self.status_label.config(text="無法識別文件類型")
            return

        # 檢查是否有儲存的位置
        if self.positions[doc_type]:
            x, y = self.positions[doc_type]
            self.process_pdf_with_position(x, y)
        else:
            self.status_label.config(text="尚未設定此類型文件的簽名位置")

    def run(self):
        self.window.mainloop()

if __name__ == "__main__":
    app = SignatureTool()
    app.run()