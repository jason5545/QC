import tkinter as tk
from tkinter import filedialog, Toplevel
from PIL import Image, ImageDraw, ImageFont, ImageTk
import fitz  # PyMuPDF
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

    def save_positions(self):
        with open('signature_positions.json', 'w') as f:
            json.dump(self.positions, f)

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

        # 重設位置按鈕
        reset_frame = tk.Frame(self.window, pady=10)
        reset_frame.pack(fill=tk.X, padx=20)
        self.reset_button = tk.Button(
            reset_frame,
            text="重設儲存的位置",
            command=self.reset_positions
        )
        self.reset_button.pack(fill=tk.X)

        # 狀態顯示
        status_frame = tk.Frame(self.window, pady=10)
        status_frame.pack(fill=tk.X, padx=20)
        self.status_label = tk.Label(status_frame, text="", wraplength=350)
        self.status_label.pack(fill=tk.X)

    def reset_positions(self):
        self.positions = {'Welding': None, 'Material': None}
        self.save_positions()
        self.status_label.config(text="已重設所有儲存的位置")

    def get_document_type(self, filename):
        if 'Welding' in filename:
            return 'Welding'
        elif 'Material' in filename:
            return 'Material'
        return None

    def show_pdf_picker(self):
        try:
            doc = fitz.open(self.input_pdf_path)
            page = doc[0]
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # 2x縮放以獲得更好的顯示質量

            # 創建新窗口
            picker_window = Toplevel(self.window)
            picker_window.title("選擇簽名位置")

            # 設定視窗大小為螢幕高度的80%
            screen_height = picker_window.winfo_screenheight()
            window_height = int(screen_height * 0.8)
            picker_window.geometry(f"{pix.width}x{window_height}")

            # 創建說明標籤
            tk.Label(picker_window, text="點擊要放置簽名的位置").pack()

            # 創建框架來容納Canvas和Scrollbar
            frame = tk.Frame(picker_window)
            frame.pack(fill=tk.BOTH, expand=True)

            # 創建垂直捲動條
            v_scrollbar = tk.Scrollbar(frame, orient=tk.VERTICAL)
            v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

            # 創建水平捲動條
            h_scrollbar = tk.Scrollbar(frame, orient=tk.HORIZONTAL)
            h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

            # 創建Canvas
            canvas = tk.Canvas(frame, 
                             yscrollcommand=v_scrollbar.set,
                             xscrollcommand=h_scrollbar.set)
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

            # 配置捲動條
            v_scrollbar.config(command=canvas.yview)
            h_scrollbar.config(command=canvas.xview)

            # 將PDF頁面轉換為PhotoImage
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            self.pdf_image = ImageTk.PhotoImage(img)

            # 在Canvas上顯示圖片
            canvas.create_image(0, 0, anchor='nw', image=self.pdf_image)

            # 設定Canvas的捲動區域
            canvas.config(scrollregion=canvas.bbox("all"))

            def on_click(event):
                try:
                    # 計算實際PDF中的位置（考慮捲動和縮放因素）
                    canvas_x = canvas.canvasx(event.x)  # 獲取相對於Canvas的座標
                    canvas_y = canvas.canvasy(event.y)
                    
                    # 轉換為PDF座標（考慮2x縮放）
                    pdf_x = canvas_x / 2
                    pdf_y = canvas_y / 2
                    
                    # 轉換為PDF坐標系統
                    pdf_x = pdf_x * 595.27 / page.rect.width
                    pdf_y = (page.rect.height - pdf_y) * 841.89 / page.rect.height
                    
                    doc_type = self.get_document_type(os.path.basename(self.input_pdf_path))
                    if doc_type:
                        self.positions[doc_type] = (pdf_x, pdf_y)
                        self.save_positions()
                        
                    picker_window.destroy()
                    self.process_pdf_with_position(pdf_x, pdf_y)
                finally:
                    doc.close()

            canvas.bind('<Button-1>', on_click)

            # 添加鍵盤捲動支援
            def on_mousewheel(event):
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            canvas.bind_all("<MouseWheel>", on_mousewheel)

            # 添加拖曳支援
            def on_press(event):
                canvas.scan_mark(event.x, event.y)
                
            def on_drag(event):
                canvas.scan_dragto(event.x, event.y, gain=1)
                
            canvas.bind("<ButtonPress-2>", on_press)  # 中鍵按下
            canvas.bind("<B2-Motion>", on_drag)       # 中鍵拖曳

        except Exception as e:
            self.status_label.config(text=f"打開PDF時發生錯誤: {e}")

    def create_signature_image(self):
        try:
            date_text = self.date_entry.get()
            if not date_text:
                self.status_label.config(text="請輸入日期")
                return None

            # 開啟原始簽名圖片
            img = Image.open("紹宇.jpg")
            
            # 設定文字和字體
            font = ImageFont.truetype("JasonHandwriting2-Regular.ttf", 100)

            # 初始化繪圖對象
            draw = ImageDraw.Draw(img)

            # 獲取文字尺寸
            try:
                text_bbox = draw.textbbox((0, 0), date_text, font=font)
                text_width = text_bbox[2] - text_bbox[0]
                text_height = text_bbox[3] - text_bbox[1]
            except AttributeError:
                text_width, text_height = draw.textsize(date_text, font=font)

            # 創建新圖片
            padding = 20
            new_image_width = img.width + text_width + padding
            new_image_height = max(img.height, text_height)
            new_img = Image.new('RGB', (new_image_width, new_image_height), color=(255, 255, 255))
            new_img.paste(img, (0, 0))

            # 在右側填充背景顏色 #d0d0d0
            draw_new_img = ImageDraw.Draw(new_img)
            right_area = (img.width, 0, new_image_width, new_image_height)
            draw_new_img.rectangle(right_area, fill=(208, 208, 208))

            # 添加日期文字
            text_y = (new_image_height - text_height) // 2
            text_position = (img.width + padding, text_y)
            draw_new_img.text(text_position, date_text, font=font, fill=(0, 0, 0))

            # 縮小整個圖片
            target_height = 50
            scale_factor = target_height / new_img.height
            target_width = int(new_image_width * scale_factor)
            
            resized_img = new_img.resize((target_width, target_height), Image.Resampling.LANCZOS)

            temp_path = f"temp_signature_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            resized_img.save(temp_path)
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
            
            can.drawImage(signature_path, x-width/2, y-height/2, width, height, preserveAspectRatio=True)
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

        # 檢查是否已有儲存的位置
        if self.positions[doc_type]:
            x, y = self.positions[doc_type]
            self.process_pdf_with_position(x, y)
        else:
            # 如果沒有儲存的位置，顯示選擇器
            self.show_pdf_picker()

    def run(self):
        self.window.mainloop()

if __name__ == "__main__":
    app = SignatureTool()
    app.run()